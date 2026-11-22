#Requires -Version 7.0
#Requires -Modules Microsoft.Graph.Authentication
<#
.SYNOPSIS
    Stage 1 of IntuneAsBuilt: captures Intune configuration from Microsoft Graph into a dated JSON snapshot.

.DESCRIPTION
    Read-only. Each run writes a new folder under <workspace>\snapshots (e.g. snapshots\2026-11-22_0330Z) containing one
    JSON file per area plus meta.json. Snapshots are the point-in-time baseline / audit evidence and the input
    for Build-IntuneAsBuilt.ps1. Script bodies are NOT stored unless -IncludeScriptContent is used; a SHA-256
    hash of each script is always stored so changes are still detectable.

.EXAMPLE
    # Unattended, app registration with certificate (recommended)
    .\Invoke-IntuneCollect.ps1 -Workspace D:\intune-as-built-config -TenantId contoso.onmicrosoft.com -ClientId <appId> -CertificateThumbprint <thumb>

.EXAMPLE
    # Lab / ad-hoc, delegated sign-in
    .\Invoke-IntuneCollect.ps1 -Interactive -IncludeScriptContent -IncludeAutopilotDevices

.EXAMPLE
    # Non-admin team member, through a dedicated delegated app registration
    .\Invoke-IntuneCollect.ps1 -Interactive -TenantId contoso.onmicrosoft.com -ClientId <delegatedAppId>
#>
[CmdletBinding(DefaultParameterSetName = 'Certificate')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Certificate')]
    [Parameter(ParameterSetName = 'Interactive')]
    [string]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'Certificate')]
    [Parameter(ParameterSetName = 'Interactive')]
    [string]$ClientId,
    [Parameter(Mandatory, ParameterSetName = 'Certificate')][string]$CertificateThumbprint,
    [Parameter(Mandatory, ParameterSetName = 'Interactive')][switch]$Interactive,

    # Private folder for tenant data (snapshots, notes, template, output). Default: $env:INTUNEASBUILT_WORKSPACE
    [string]$Workspace,
    # Override where snapshots are written (default: <workspace>\snapshots)
    [string]$OutputRoot,
    [switch]$IncludeScriptContent,
    [switch]$IncludeAutopilotDevices
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib/Graph.ps1')
. (Join-Path $PSScriptRoot 'lib/Workspace.ps1')
$ws = Resolve-AsBuiltWorkspace -Workspace $Workspace -ToolRoot $PSScriptRoot
if (-not $OutputRoot) { $OutputRoot = $ws.Snapshots }
$CollectorVersion = '1.0.0'
$dm  = 'beta/deviceManagement'
$dam = 'beta/deviceAppManagement'

#region Connect
$scopes = @(
    'DeviceManagementConfiguration.Read.All', 'DeviceManagementApps.Read.All',
    'DeviceManagementServiceConfig.Read.All', 'DeviceManagementScripts.Read.All',
    'DeviceManagementRBAC.Read.All', 'Group.Read.All', 'Organization.Read.All'
)
if ($PSCmdlet.ParameterSetName -eq 'Certificate') {
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
}
else {
    $p = @{ Scopes = $scopes; NoWelcome = $true }
    if ($TenantId) { $p.TenantId = $TenantId }
    if ($ClientId) { $p.ClientId = $ClientId }   # dedicated delegated app registration (recommended)
    Connect-MgGraph @p
}
$ctx = Get-MgContext
#endregion

$stamp   = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd_HHmm'Z'")
$snapDir = Join-Path $OutputRoot $stamp
New-Item -ItemType Directory -Path $snapDir -Force | Out-Null

$collected = [ordered]@{}
$warnings  = [System.Collections.Generic.List[string]]::new()
$settingDefinitions = @{}

#region Helpers
function Save-Section([string]$Name, $Data) {
    $collected[$Name] = $Data
    ConvertTo-Json -InputObject $Data -Depth 100 |
        Set-Content -LiteralPath (Join-Path $snapDir "$Name.json") -Encoding utf8NoBOM
}

function Invoke-Section([string]$Name, [scriptblock]$Script) {
    Write-Host "Collecting $Name..." -ForegroundColor Cyan
    try {
        $data = & $Script
        Save-Section $Name @($data)
        Write-Host "  $(@($data).Count) item(s)"
    }
    catch {
        $msg = "${Name}: $($_.Exception.Message)"
        Write-Warning $msg
        $warnings.Add($msg)
    }
}

function Add-Assignments($Items, [string]$BasePath) {
    foreach ($i in @($Items)) {
        $a = Get-AsBuiltGraphCollection "$BasePath/$($i.id)/assignments"
        $i | Add-Member -NotePropertyName assignments -NotePropertyValue @($a) -Force
    }
    @($Items) | Sort-Object { $_.displayName ?? $_.name }, id
}

function Set-ContentHash($Object, [string[]]$Properties) {
    foreach ($prop in $Properties) {
        $b64 = $Object.$prop
        if ($b64) {
            $bytes = [Convert]::FromBase64String($b64)
            $hash = [Convert]::ToHexString([System.Security.Cryptography.SHA256]::HashData($bytes))
            $Object | Add-Member -NotePropertyName "$($prop)Sha256" -NotePropertyValue $hash -Force
            if (-not $IncludeScriptContent) { $Object.$prop = $null }
        }
    }
}

function Find-GroupIds($Node, [System.Collections.Generic.HashSet[string]]$Set) {
    if ($null -eq $Node -or $Node -is [string] -or $Node -is [ValueType]) { return }
    if ($Node -is [System.Collections.IDictionary]) { foreach ($v in $Node.Values) { Find-GroupIds $v $Set }; return }
    if ($Node -is [System.Collections.IEnumerable]) { foreach ($v in $Node) { Find-GroupIds $v $Set }; return }
    if ($Node -isnot [System.Management.Automation.PSCustomObject]) { return }
    foreach ($p in $Node.PSObject.Properties) {
        if ($p.Name -eq 'groupId' -and $p.Value -is [string] -and $p.Value) { [void]$Set.Add($p.Value) }
        else { Find-GroupIds $p.Value $Set }
    }
}
#endregion

#region Tenant-wide lookups
Invoke-Section 'organization'   { Get-AsBuiltGraphCollection 'v1.0/organization?$select=id,displayName' }
Invoke-Section 'roleScopeTags'  { Get-AsBuiltGraphCollection "$dm/roleScopeTags" | Sort-Object displayName }
Invoke-Section 'assignmentFilters' { Get-AsBuiltGraphCollection "$dm/assignmentFilters" | Sort-Object displayName }
#endregion

#region Windows enrollment
Invoke-Section 'autopilotProfiles' {
    Add-Assignments (Get-AsBuiltGraphCollection "$dm/windowsAutopilotDeploymentProfiles") "$dm/windowsAutopilotDeploymentProfiles"
}
Invoke-Section 'enrollmentConfigurations' {
    Add-Assignments (Get-AsBuiltGraphCollection "$dm/deviceEnrollmentConfigurations") "$dm/deviceEnrollmentConfigurations"
}
if ($IncludeAutopilotDevices) {
    Invoke-Section 'autopilotDevices' {
        Get-AsBuiltGraphCollection "$dm/windowsAutopilotDeviceIdentities" |
            Select-Object id, serialNumber, groupTag, manufacturer, model, enrollmentState, deploymentProfileAssignmentStatus |
            Sort-Object serialNumber
    }
}
#endregion

#region Policies
# Settings catalog, endpoint security, baselines and Autopilot device preparation all live in configurationPolicies.
Invoke-Section 'configurationPolicies' {
    $policies = Get-AsBuiltGraphCollection "$dm/configurationPolicies"
    foreach ($p in $policies) {
        $settings = Get-AsBuiltGraphCollection "$dm/configurationPolicies/$($p.id)/settings?`$expand=settingDefinitions"
        foreach ($s in $settings) {
            foreach ($d in @($s.settingDefinitions)) {
                if ($d -and -not $settingDefinitions.ContainsKey($d.id)) {
                    $opts = @{}
                    foreach ($o in @($d.options)) { if ($o.itemId) { $opts[$o.itemId] = $o.displayName } }
                    $settingDefinitions[$d.id] = @{ name = $d.displayName; options = $opts }
                }
            }
        }
        $p | Add-Member -NotePropertyName settings -NotePropertyValue @($settings | ForEach-Object { $_.settingInstance }) -Force
    }
    Add-Assignments $policies "$dm/configurationPolicies"
}
Write-Host "Saving settingDefinitions ($($settingDefinitions.Count) definitions)..." -ForegroundColor Cyan
$sortedDefs = [ordered]@{}
foreach ($k in ($settingDefinitions.Keys | Sort-Object)) { $sortedDefs[$k] = $settingDefinitions[$k] }
Save-Section 'settingDefinitions' $sortedDefs

Invoke-Section 'deviceConfigurations' {
    Add-Assignments (Get-AsBuiltGraphCollection "$dm/deviceConfigurations") "$dm/deviceConfigurations"
}
Invoke-Section 'groupPolicyConfigurations' {
    $gps = Get-AsBuiltGraphCollection "$dm/groupPolicyConfigurations"
    foreach ($g in $gps) {
        $dvs = Get-AsBuiltGraphCollection "$dm/groupPolicyConfigurations/$($g.id)/definitionValues?`$expand=definition"
        foreach ($dv in $dvs) {
            $pvs = Get-AsBuiltGraphCollection "$dm/groupPolicyConfigurations/$($g.id)/definitionValues/$($dv.id)/presentationValues?`$expand=presentation"
            $dv | Add-Member -NotePropertyName presentationValues -NotePropertyValue @($pvs) -Force
        }
        $g | Add-Member -NotePropertyName definitionValues -NotePropertyValue @($dvs) -Force
    }
    Add-Assignments $gps "$dm/groupPolicyConfigurations"
}
Invoke-Section 'compliancePolicies' {
    $cps = Get-AsBuiltGraphCollection "$dm/deviceCompliancePolicies?`$expand=scheduledActionsForRule(`$expand=scheduledActionConfigurations)"
    Add-Assignments $cps "$dm/deviceCompliancePolicies"
}
Invoke-Section 'intents' {   # legacy endpoint security / baseline templates
    $intents = Get-AsBuiltGraphCollection "$dm/intents"
    foreach ($i in $intents) {
        $s = Get-AsBuiltGraphCollection "$dm/intents/$($i.id)/settings"
        $i | Add-Member -NotePropertyName settings -NotePropertyValue @($s) -Force
    }
    Add-Assignments $intents "$dm/intents"
}
Invoke-Section 'featureUpdateProfiles' { Add-Assignments (Get-AsBuiltGraphCollection "$dm/windowsFeatureUpdateProfiles") "$dm/windowsFeatureUpdateProfiles" }
Invoke-Section 'qualityUpdateProfiles' { Add-Assignments (Get-AsBuiltGraphCollection "$dm/windowsQualityUpdateProfiles") "$dm/windowsQualityUpdateProfiles" }
Invoke-Section 'driverUpdateProfiles'  { Add-Assignments (Get-AsBuiltGraphCollection "$dm/windowsDriverUpdateProfiles") "$dm/windowsDriverUpdateProfiles" }
#endregion

#region Apps & scripts
Invoke-Section 'mobileApps' {
    $list = Get-AsBuiltGraphCollection "$dam/mobileApps"
    $apps = foreach ($app in $list) {
        $full = Invoke-AsBuiltGraph "$dam/mobileApps/$($app.id)?`$expand=assignments"
        $rel = Get-AsBuiltGraphCollection "$dam/mobileApps/$($app.id)/relationships"
        $full | Add-Member -NotePropertyName relationships -NotePropertyValue @($rel) -Force
        $full.PSObject.Properties.Remove('largeIcon')
        $full
    }
    @($apps) | Sort-Object displayName, id
}

Invoke-Section 'platformScripts' {
    $scripts = foreach ($s in (Get-AsBuiltGraphCollection "$dm/deviceManagementScripts")) {
        $full = Invoke-AsBuiltGraph "$dm/deviceManagementScripts/$($s.id)?`$expand=assignments"
        Set-ContentHash $full 'scriptContent'
        $full
    }
    @($scripts) | Sort-Object displayName, id
}

Invoke-Section 'remediations' {
    $all = Get-AsBuiltGraphCollection "$dm/deviceHealthScripts"
    $kept = foreach ($s in $all) {
        $full = Invoke-AsBuiltGraph "$dm/deviceHealthScripts/$($s.id)"
        $a = Get-AsBuiltGraphCollection "$dm/deviceHealthScripts/$($s.id)/assignments"
        # Skip Microsoft-published scripts unless they are assigned
        if ($full.isGlobalScript -and @($a).Count -eq 0) { continue }
        $full | Add-Member -NotePropertyName assignments -NotePropertyValue @($a) -Force
        Set-ContentHash $full 'detectionScriptContent', 'remediationScriptContent'
        $full
    }
    @($kept) | Sort-Object displayName, id
}
#endregion

#region Resolve group names
Write-Host 'Resolving group names...' -ForegroundColor Cyan
$groupIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
Find-GroupIds $collected $groupIds
$groups = [ordered]@{}
foreach ($gid in ($groupIds | Sort-Object)) {
    try { $groups[$gid] = (Invoke-AsBuiltGraph "v1.0/groups/$gid`?`$select=displayName").displayName }
    catch { $groups[$gid] = "(group not found: $gid)" }
}
Save-Section 'groups' $groups
Write-Host "  $($groups.Count) group(s)"
#endregion

#region Meta
$counts = [ordered]@{}
foreach ($k in $collected.Keys) { if ($k -notin 'settingDefinitions', 'groups') { $counts[$k] = @($collected[$k]).Count } }
$meta = [ordered]@{
    tenantId         = $ctx.TenantId
    tenantName       = @($collected['organization'])[0].displayName
    collectedAtUtc   = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    collectedBy      = $ctx.Account ?? $ctx.AppName ?? $ctx.ClientId
    authType         = "$($ctx.AuthType)"
    collectorVersion = $CollectorVersion
    options          = [ordered]@{ includeScriptContent = [bool]$IncludeScriptContent; includeAutopilotDevices = [bool]$IncludeAutopilotDevices }
    counts           = $counts
    warnings         = @($warnings)
}
ConvertTo-Json -InputObject $meta -Depth 10 | Set-Content -LiteralPath (Join-Path $snapDir 'meta.json') -Encoding utf8NoBOM
#endregion

# The Graph SDK can warn that it failed to clear its token cache on sign-out (known SDK issue,
# more likely when signed in without -TenantId). The snapshot is already complete at this point.
try { Disconnect-MgGraph -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null }
catch { Write-Verbose "Disconnect-MgGraph: $($_.Exception.Message)" }
Write-Host "`nSnapshot written to $snapDir" -ForegroundColor Green
if ($warnings.Count) { Write-Warning "$($warnings.Count) section(s) had errors - see meta.json" }
$snapDir
