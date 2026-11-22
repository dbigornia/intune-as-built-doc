#Requires -Version 7.0
<#
.SYNOPSIS
    Stage 2 of IntuneAsBuilt: turns a snapshot from Invoke-IntuneCollect.ps1 into a Word as-built document.

.DESCRIPTION
    Works offline from the snapshot JSON - no Graph connection, no Word install, no extra modules.
    Covers Windows enrollment, device policies, and apps & scripts. Hand-written design notes
    (purpose, owner, rollback ...) are merged in from a CSV keyed by object ID.

.PARAMETER Workspace
    Private folder holding tenant data: snapshots\, output\, notes.csv and an optional template.docx/.dotx.
    Default: $env:INTUNEASBUILT_WORKSPACE. In a workspace, notes.csv and the template are picked up
    automatically and notes.csv is refreshed on every build. Explicit parameters override these defaults.
.PARAMETER SnapshotPath
    Snapshot folder (e.g. snapshots\2026-11-22_0330Z). Defaults to the newest folder under <workspace>\snapshots.
.PARAMETER TemplatePath
    Optional company .docx/.dotx. Its headers, footers, page setup and styles are kept. Put a paragraph
    containing {{ASBUILT}} where the generated content should go (otherwise the body is replaced).
.PARAMETER NotesPath
    CSV of design notes: ObjectId, Category, Name, Purpose, Owner, DependsOn, LastValidated, Rollback.
.PARAMETER ExportNotesTemplate
    Writes/refreshes a notes CSV listing every documented object. Existing notes in that file are kept.
.PARAMETER Pdf
    Also save a PDF next to the .docx. Uses Word if installed (builds the table of contents and
    bookmarks, and saves the .docx with the TOC filled in), otherwise LibreOffice.
.PARAMETER PdfEngine
    Auto (default), Word or LibreOffice.
.PARAMETER AllPlatforms
    Include iOS/Android/macOS objects. Default is Windows only.

.EXAMPLE
    .\Build-IntuneAsBuilt.ps1 -Workspace D:\intune-as-built-config -Pdf
.EXAMPLE
    .\Build-IntuneAsBuilt.ps1 -Pdf -NotesPath .\notes.csv -ExportNotesTemplate .\notes.csv -TemplatePath .\templates\Company-Template.docx
#>
[CmdletBinding()]
param(
    [string]$Workspace,
    [string]$SnapshotPath,
    [string]$OutputPath,
    [string]$TemplatePath,
    [string]$NotesPath,
    [string]$ExportNotesTemplate,
    [string]$Title = 'Intune As-Built',
    [switch]$AllPlatforms,
    [switch]$IncludeEmptyValues,
    [switch]$HideMissingNotes,
    [switch]$Pdf,
    [ValidateSet('Auto', 'Word', 'LibreOffice')][string]$PdfEngine = 'Auto'
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib/Docx.ps1')
. (Join-Path $PSScriptRoot 'lib/Pdf.ps1')
. (Join-Path $PSScriptRoot 'lib/Workspace.ps1')
$ws = Resolve-AsBuiltWorkspace -Workspace $Workspace -ToolRoot $PSScriptRoot
if (-not $NotesPath -and $ws.Notes) { $NotesPath = $ws.Notes }
if (-not $PSBoundParameters.ContainsKey('ExportNotesTemplate') -and $ws.Notes) { $ExportNotesTemplate = $ws.Notes }
if (-not $TemplatePath -and $ws.Template) { $TemplatePath = $ws.Template; Write-Host "Using template $TemplatePath" }

#region Load snapshot
if (-not $SnapshotPath) {
    if (Test-Path -LiteralPath $ws.Snapshots) {
        $SnapshotPath = Get-ChildItem -Directory -LiteralPath $ws.Snapshots | Sort-Object Name | Select-Object -Last 1 -ExpandProperty FullName
    }
    if (-not $SnapshotPath) { throw "No snapshot found in $($ws.Snapshots). Run Invoke-IntuneCollect.ps1 first or pass -SnapshotPath." }
}
Write-Host "Using snapshot $SnapshotPath" -ForegroundColor Cyan

function Read-Snapshot([string]$Name, [switch]$AsHashtable) {
    $p = Join-Path $SnapshotPath "$Name.json"
    if (-not (Test-Path -LiteralPath $p)) { if ($AsHashtable) { return @{} } else { return , @() } }
    $raw = Get-Content -Raw -LiteralPath $p
    if ($AsHashtable) { return ($raw | ConvertFrom-Json -AsHashtable) }
    return , @($raw | ConvertFrom-Json)
}

$meta          = Get-Content -Raw -LiteralPath (Join-Path $SnapshotPath 'meta.json') | ConvertFrom-Json
$script:Groups = Read-Snapshot 'groups' -AsHashtable
$script:Defs   = Read-Snapshot 'settingDefinitions' -AsHashtable
$script:Filters = @{}; foreach ($f in (Read-Snapshot 'assignmentFilters')) { $script:Filters[$f.id] = $f.displayName }
$script:ScopeTags = @{}; foreach ($t in (Read-Snapshot 'roleScopeTags')) { $script:ScopeTags[[string]$t.id] = $t.displayName }

$autopilotProfiles = Read-Snapshot 'autopilotProfiles'
$enrollmentConfigs = Read-Snapshot 'enrollmentConfigurations'
$autopilotDevices  = Read-Snapshot 'autopilotDevices'
$configPolicies    = Read-Snapshot 'configurationPolicies'
$deviceConfigs     = Read-Snapshot 'deviceConfigurations'
$gpConfigs         = Read-Snapshot 'groupPolicyConfigurations'
$compliance        = Read-Snapshot 'compliancePolicies'
$intents           = Read-Snapshot 'intents'
$featureUpdates    = Read-Snapshot 'featureUpdateProfiles'
$qualityUpdates    = Read-Snapshot 'qualityUpdateProfiles'
$driverUpdates     = Read-Snapshot 'driverUpdateProfiles'
$apps              = Read-Snapshot 'mobileApps'
$platformScripts   = Read-Snapshot 'platformScripts'
$remediations      = Read-Snapshot 'remediations'

$script:AppNames = @{}; foreach ($a in $apps) { $script:AppNames[$a.id] = $a.displayName }

$script:Notes = @{}
if ($NotesPath -and (Test-Path -LiteralPath $NotesPath)) {
    foreach ($n in (Import-Csv -LiteralPath $NotesPath)) { if ($n.ObjectId) { $script:Notes[$n.ObjectId] = $n } }
    Write-Host "Loaded $($script:Notes.Count) design note(s)"
}
$script:Inventory = [System.Collections.Generic.List[object]]::new()
#endregion

#region Formatting helpers
$script:CommonSkip = @(
    'id', 'displayName', 'name', 'description', 'createdDateTime', 'lastModifiedDateTime', 'version', 'assignments',
    'roleScopeTagIds', 'roleScopeTags', 'supportsScopeTags', 'settings', 'settingCount', 'templateReference', 'technologies',
    'platforms', 'creationSource', 'priorityMetaData', 'isAssigned', 'largeIcon', 'scheduledActionsForRule', 'definitionValues',
    'relationships', 'scriptContent', 'detectionScriptContent', 'remediationScriptContent', 'uploadState', 'publishingState',
    'committedContentVersion', 'size', 'dependentAppCount', 'supersedingAppCount', 'supersededAppCount', 'isFeatured',
    'deviceStatusOverview', 'userStatusOverview', 'deviceStatuses', 'userStatuses', 'isGlobalScript', 'highestAvailableVersion',
    'isMigratingToConfigurationPolicy', 'deviceSettingStateSummaries'
)

# Properties holding IDs that should be shown as names
$script:IdResolvers = @{
    selectedMobileAppIds = { param($v) @($v | ForEach-Object { $script:AppNames[$_] ?? $_ }) }
}

function Format-Label([string]$Name) {
    $s = $Name -creplace '(?<=[a-z0-9])(?=[A-Z])', ' ' -creplace '(?<=[A-Z])(?=[A-Z][a-z])', ' '
    if ($s) { $s.Substring(0, 1).ToUpper() + $s.Substring(1) } else { $s }
}

function Format-TypeName([string]$Type) {
    Format-Label ($Type -replace '^#microsoft\.graph\.', '' -replace '^win32LobApp(?=[A-Z])', '')
}

function Format-Scalar($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [bool]) { return $(if ($v) { 'Yes' } else { 'No' }) }
    if ($v -is [datetime]) { return $v.ToString('yyyy-MM-dd HH:mm') }
    [string]$v
}

function Format-Inline($v, [int]$Depth = 0) {
    if ($null -eq $v) { return '' }
    if ($v -is [System.Management.Automation.PSCustomObject]) {
        $parts = [System.Collections.Generic.List[string]]::new()
        $encrypted = [bool]($v.PSObject.Properties['isEncrypted'] -and $v.isEncrypted)
        foreach ($p in $v.PSObject.Properties) {
            if ($p.Name -like '*@odata*' -or $null -eq $p.Value -or $p.Name -eq 'isEncrypted') { continue }
            if ($p.Value -is [string] -and $p.Value -eq '') { continue }
            if ($encrypted -and $p.Name -eq 'value') { $parts.Add('Value: (encrypted)'); continue }
            $val = if ($Depth -ge 3) { ConvertTo-Json -InputObject $p.Value -Compress -Depth 10 } else { Format-Inline $p.Value ($Depth + 1) }
            $parts.Add("$(Format-Label $p.Name): $val")
        }
        $type = $v.'@odata.type'
        $prefix = if ($type) { "[$(Format-TypeName $type)] " } else { '' }
        return $prefix + ($parts -join '; ')
    }
    if ($v -is [System.Collections.IList]) { return (@($v | ForEach-Object { Format-Inline $_ ($Depth + 1) }) -join ', ') }
    Format-Scalar $v
}

function ConvertTo-PropertyRows {
    <# Generic object -> Setting/Value rows. Nested objects become "Parent > Child" rows. #>
    param($Object, [string[]]$Skip = @(), [string]$Prefix = '')
    $rows = [System.Collections.Generic.List[object]]::new()
    if ($null -eq $Object) { return , $rows }
    foreach ($p in $Object.PSObject.Properties) {
        $n = $p.Name
        if ($n -like '*@odata*') { continue }
        if (-not $Prefix -and ($script:CommonSkip -contains $n -or $Skip -contains $n)) { continue }
        $label = if ($Prefix) { "$Prefix > $(Format-Label $n)" } else { Format-Label $n }
        $v = $p.Value
        if ($script:IdResolvers.ContainsKey($n) -and $null -ne $v) { $v = & $script:IdResolvers[$n] $v }
        $empty = ($null -eq $v) -or ($v -is [string] -and $v -eq '') -or ($v -is [System.Collections.IList] -and $v.Count -eq 0)
        if ($empty) { if ($IncludeEmptyValues) { $rows.Add(@($label, '')) }; continue }
        if ($v -is [System.Management.Automation.PSCustomObject]) {
            foreach ($r in (ConvertTo-PropertyRows -Object $v -Prefix $label)) { $rows.Add($r) }
            continue
        }
        $text = if ($v -is [System.Collections.IList]) { @($v | ForEach-Object { Format-Inline $_ }) -join "`n" } else { Format-Scalar $v }
        $rows.Add(@($label, $text))
    }
    return , $rows
}

function Get-ObjectName($o) { $o.displayName ?? $o.name ?? $o.id }

function Resolve-Group([string]$Id) { $script:Groups[$Id] ?? "(unknown group $Id)" }

function Get-TargetInfo($Target) {
    $t = [string]$Target.'@odata.type'
    $mode, $name = if ($t -like '*exclusionGroupAssignmentTarget') { 'Exclude', (Resolve-Group $Target.groupId) }
    elseif ($t -like '*groupAssignmentTarget') { 'Include', (Resolve-Group $Target.groupId) }
    elseif ($t -like '*allDevicesAssignmentTarget') { 'Include', 'All devices' }
    elseif ($t -like '*allLicensedUsersAssignmentTarget') { 'Include', 'All users' }
    else { 'Include', (Format-TypeName $t) }
    $filter = ''
    $fid = $Target.deviceAndAppManagementAssignmentFilterId
    if ($fid -and $fid -ne '00000000-0000-0000-0000-000000000000') {
        $filter = "$($script:Filters[$fid] ?? $fid) ($($Target.deviceAndAppManagementAssignmentFilterType))"
    }
    [pscustomobject]@{ Mode = $mode; Name = $name; Filter = $filter }
}

function Get-AssignmentRows($Assignments) {
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($a in @($Assignments)) {
        if ($null -eq $a -or $null -eq $a.target) { continue }
        $ti = Get-TargetInfo $a.target
        $details = [System.Collections.Generic.List[string]]::new()
        if ($a.intent)      { $details.Add("Intent: $(Format-Label $a.intent)") }
        if ($a.runSchedule) { $details.Add("Schedule: $(Format-Inline $a.runSchedule)") }
        if ($a.settings)    { $details.Add((Format-Inline $a.settings)) }
        if ($a.source -and $a.source -ne 'direct') { $details.Add("Source: $($a.source)") }
        $rows.Add(@($ti.Mode, $ti.Name, $ti.Filter, ($details -join "`n")))
    }
    $sorted = [System.Collections.Generic.List[object]]::new()
    foreach ($r in ($rows | Sort-Object { $_[0] -eq 'Exclude' }, { $_[1] })) { $sorted.Add($r) }
    return , $sorted
}

function Test-InScope($Object) {
    if ($AllPlatforms) { return $true }
    if ([string]$Object.'@odata.type' -like '*AssignmentFilter' -and $Object.platform) { return ([string]$Object.platform -match 'windows') }
    if ($Object.PSObject.Properties['platforms'] -and $Object.platforms) { return ([string]$Object.platforms -match 'windows') }
    return -not ([string]$Object.'@odata.type' -match '^#microsoft\.graph\.(ios|macOS|android|aosp|managedIOS|managedAndroid|windowsPhone)')
}

function Get-ScopeTagText($o) {
    @($o.roleScopeTagIds | Where-Object { $_ } | ForEach-Object { $script:ScopeTags[[string]$_] ?? $_ }) -join ', '
}
#endregion

#region Settings renderers
function Resolve-ChoiceValue($Def, $Value) {
    if ($null -eq $Value) { return '' }
    if ($Def -and $Def.options -and $Def.options[$Value]) { return $Def.options[$Value] }
    [string]$Value
}

function Format-SimpleValue($Sv) {
    if ($null -eq $Sv) { return '' }
    if ([string]$Sv.'@odata.type' -like '*secret*') { return '(secret - not exported)' }
    Format-Scalar $Sv.value
}

function Add-SettingInstanceRows($Instance, [int]$Level, $Rows) {
    if ($null -eq $Instance) { return }
    $def = $script:Defs[$Instance.settingDefinitionId]
    $name = if ($def -and $def.name) { $def.name } else { $Instance.settingDefinitionId }
    $label = if ($Level -gt 0) { ('   ' * $Level) + '- ' + $name } else { $name }
    $type = [string]$Instance.'@odata.type'
    $next = $Level + 1

    if ($type -like '*choiceSettingCollectionInstance') {
        $vals = @($Instance.choiceSettingCollectionValue)
        $Rows.Add(@($label, (@($vals | ForEach-Object { Resolve-ChoiceValue $def $_.value }) -join "`n")))
        foreach ($cv in $vals) { foreach ($c in @($cv.children)) { Add-SettingInstanceRows $c $next $Rows } }
    }
    elseif ($type -like '*choiceSettingInstance') {
        $cv = $Instance.choiceSettingValue
        $Rows.Add(@($label, (Resolve-ChoiceValue $def $cv.value)))
        foreach ($c in @($cv.children)) { Add-SettingInstanceRows $c $next $Rows }
    }
    elseif ($type -like '*simpleSettingCollectionInstance') {
        $Rows.Add(@($label, (@($Instance.simpleSettingCollectionValue | ForEach-Object { Format-SimpleValue $_ }) -join "`n")))
    }
    elseif ($type -like '*simpleSettingInstance') {
        $Rows.Add(@($label, (Format-SimpleValue $Instance.simpleSettingValue)))
    }
    elseif ($type -like '*groupSettingCollectionInstance') {
        $i = 0
        foreach ($gv in @($Instance.groupSettingCollectionValue)) {
            $i++
            $Rows.Add(@("$label [$i]", ''))
            foreach ($c in @($gv.children)) { Add-SettingInstanceRows $c $next $Rows }
        }
    }
    elseif ($type -like '*groupSettingInstance') {
        $Rows.Add(@($label, ''))
        foreach ($c in @($Instance.groupSettingValue.children)) { Add-SettingInstanceRows $c $next $Rows }
    }
    else { $Rows.Add(@($label, (Format-Inline $Instance))) }
}

function Get-SettingsCatalogRows($Policy) {
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($inst in @($Policy.settings)) { Add-SettingInstanceRows $inst 0 $rows }
    return , $rows
}

function Get-AdmxRows($Policy) {
    $rows = [System.Collections.Generic.List[object]]::new()
    $dvs = @($Policy.definitionValues | Where-Object { $_ } | Sort-Object { $_.definition.categoryPath }, { $_.definition.displayName })
    foreach ($dv in $dvs) {
        $d = $dv.definition
        $label = if ($d.categoryPath) { "$($d.categoryPath) > $($d.displayName)" } else { $d.displayName }
        $lines = [System.Collections.Generic.List[string]]::new()
        $lines.Add($(if ($dv.enabled) { 'Enabled' } else { 'Disabled' }))
        foreach ($pv in @($dv.presentationValues | Where-Object { $_ })) {
            $val = if ($pv.PSObject.Properties['values'] -and $null -ne $pv.values) {
                @($pv.values | ForEach-Object { if ($_ -is [System.Management.Automation.PSCustomObject]) { "$($_.name)=$($_.value)" } else { [string]$_ } }) -join '; '
            }
            elseif ($pv.PSObject.Properties['value']) { Format-Scalar $pv.value }
            else { Format-Inline $pv }
            $lines.Add("$($pv.presentation.label) $val".Trim())
        }
        $rows.Add(@($label, (Format-Label ([string]$d.classType)), ($lines -join "`n")))
    }
    return , $rows
}

function Get-IntentRows($Intent) {
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($s in @($Intent.settings)) {
        if ($null -eq $s -or $s.valueJson -eq 'null') { continue }
        $label = Format-Label (([string]$s.definitionId -split '_')[-1])
        $val = $s.valueJson
        try { $val = Format-Inline ($s.valueJson | ConvertFrom-Json) } catch { }
        $rows.Add(@($label, $val))
    }
    return , $rows
}

function Get-ComplianceRows($Policy) {
    $rows = ConvertTo-PropertyRows $Policy
    $actions = foreach ($r in @($Policy.scheduledActionsForRule)) {
        foreach ($c in @($r.scheduledActionConfigurations)) {
            if ($c) { "$(Format-Label $c.actionType) after $($c.gracePeriodHours) h" }
        }
    }
    if ($actions) { $rows.Add(@('Actions for noncompliance', (@($actions) -join "`n"))) }
    return , $rows
}
#endregion

#region Document section writers
function Write-DesignNotes($Doc, [string]$Id) {
    $n = $script:Notes[$Id]
    $fields = [ordered]@{ Purpose = 'Purpose'; Owner = 'Owner'; DependsOn = 'Depends on'; LastValidated = 'Last validated in lab'; Rollback = 'Rollback' }
    $rows = @(foreach ($k in $fields.Keys) { if ($n -and $n.$k) { , @($fields[$k], $n.$k) } })
    if ($rows.Count) {
        Add-DocxTable -Doc $Doc -Headers 'Design note', 'Detail' -Rows $rows -Widths 0.25, 0.75
    }
    elseif (-not $HideMissingNotes) {
        Add-DocxParagraph -Doc $Doc -Text 'Design notes: not yet documented.' -Italic -Color '808080'
    }
}

function Write-ObjectSection {
    param(
        $Doc, $Object, [string]$Category, [string]$TypeLabel,
        $SettingRows, [string[]]$SettingHeaders = @('Setting', 'Value'), [double[]]$SettingWidths = @(0.42, 0.58),
        [switch]$NoAssignments
    )
    $name = Get-ObjectName $Object
    Add-DocxHeading -Doc $Doc -Text $name -Level 3

    $overview = [System.Collections.Generic.List[object]]::new()
    $overview.Add(@('Type', $TypeLabel))
    if ($Object.description) { $overview.Add(@('Description', $Object.description)) }
    if ($Object.PSObject.Properties['platforms'] -and $Object.platforms) { $overview.Add(@('Platform', [string]$Object.platforms)) }
    if ($Object.createdDateTime) { $overview.Add(@('Created', (Format-Scalar $Object.createdDateTime))) }
    if ($Object.lastModifiedDateTime) { $overview.Add(@('Last modified', (Format-Scalar $Object.lastModifiedDateTime))) }
    $tags = Get-ScopeTagText $Object
    if ($tags) { $overview.Add(@('Scope tags', $tags)) }
    $overview.Add(@('Object ID', $Object.id))
    Add-DocxTable -Doc $Doc -Headers 'Property', 'Value' -Rows $overview -Widths 0.25, 0.75

    Write-DesignNotes $Doc $Object.id

    if (-not $NoAssignments) {
        Add-DocxParagraph -Doc $Doc -Text 'Assignments' -Bold -KeepNext
        $ar = Get-AssignmentRows $Object.assignments
        if ($ar.Count) { Add-DocxTable -Doc $Doc -Headers 'Mode', 'Target', 'Filter', 'Details' -Rows $ar -Widths 0.12, 0.36, 0.24, 0.28 }
        else { Add-DocxParagraph -Doc $Doc -Text 'Not assigned.' -Italic -Color '808080' }
    }

    if ($SettingRows -and $SettingRows.Count) {
        Add-DocxParagraph -Doc $Doc -Text 'Configuration' -Bold -KeepNext
        Add-DocxTable -Doc $Doc -Headers $SettingHeaders -Rows $SettingRows -Widths $SettingWidths
    }
    $script:Inventory.Add([pscustomobject]@{ ObjectId = $Object.id; Category = $Category; Name = $name })
}

function Write-ObjectList {
    param(
        $Doc, [string]$Heading, $Objects, [string]$Category,
        [scriptblock]$TypeLabel, [scriptblock]$Rows,
        [string[]]$Headers = @('Setting', 'Value'), [double[]]$Widths = @(0.42, 0.58), [switch]$NoAssignments
    )
    Add-DocxHeading -Doc $Doc -Text $Heading -Level 2
    $objs = @($Objects | Where-Object { $_ } | Sort-Object { Get-ObjectName $_ })
    if ($objs.Count -eq 0) { Add-DocxParagraph -Doc $Doc -Text 'None configured.' -Italic -Color '808080'; return }
    foreach ($o in $objs) {
        Write-ObjectSection -Doc $Doc -Object $o -Category $Category -TypeLabel (& $TypeLabel $o) -SettingRows (& $Rows $o) `
            -SettingHeaders $Headers -SettingWidths $Widths -NoAssignments:$NoAssignments
    }
}
#endregion

#region Classify
$apdpPolicies = @(); $catalogPolicies = @(); $esPolicies = @(); $baselinePolicies = @()
foreach ($p in ($configPolicies | Where-Object { Test-InScope $_ })) {
    $family = [string]$p.templateReference.templateFamily
    if ($family -eq 'enrollmentConfiguration' -or [string]$p.technologies -match 'enrollment') { $apdpPolicies += $p }
    elseif ($family -like 'endpointSecurity*') { $esPolicies += $p }
    elseif ($family -eq 'baseline') { $baselinePolicies += $p }
    else { $catalogPolicies += $p }
}
$deviceConfigsInScope = @($deviceConfigs | Where-Object { Test-InScope $_ })
$updateRings  = @($deviceConfigsInScope | Where-Object { $_.'@odata.type' -like '*windowsUpdateForBusinessConfiguration' })
$templates    = @($deviceConfigsInScope | Where-Object { $_.'@odata.type' -notlike '*windowsUpdateForBusinessConfiguration' })
$complianceIn = @($compliance | Where-Object { Test-InScope $_ })
$espConfigs   = @($enrollmentConfigs | Where-Object { $_.'@odata.type' -like '*windows10EnrollmentCompletionPageConfiguration' })
$otherEnroll  = @($enrollmentConfigs | Where-Object { $_.'@odata.type' -notlike '*windows10EnrollmentCompletionPageConfiguration' })
$appsIn       = @($apps | Where-Object { Test-InScope $_ })
$scriptsIn    = @($platformScripts)
$remIn        = @($remediations)
#endregion

#region Build document
$doc = New-AsBuiltDocument -TemplatePath $TemplatePath
$tenantName = $meta.tenantName ?? $meta.tenantId
$collectedAt = ([datetime]$meta.collectedAtUtc).ToUniversalTime().ToString('yyyy-MM-dd HH:mm')
$typeLabel = { param($o) Format-TypeName $o.'@odata.type' }
$propRows  = { param($o) ConvertTo-PropertyRows $o }
$catalogRows = { param($o) Get-SettingsCatalogRows $o }
$templateLabel = { param($o) $o.templateReference.templateDisplayName ?? 'Settings catalog' }

# Title page
Add-DocxParagraph -Doc $doc -Text $Title -Style 'Title'
Add-DocxParagraph -Doc $doc -Text $tenantName -Style 'Subtitle'
Add-DocxParagraph -Doc $doc -Text "Configuration snapshot: $collectedAt UTC" -Style 'Subtitle'
Add-DocxParagraph -Doc $doc -Text "Document generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')   |   Scope: $(if ($AllPlatforms) { 'All platforms' } else { 'Windows' })" -Color '808080'
Add-DocxParagraph -Doc $doc -Text ''
Add-DocxParagraph -Doc $doc -Text 'Contents' -Bold -Size 28
Add-DocxToc -Doc $doc

# 1. About
Add-DocxHeading -Doc $doc -Text 'About this document' -Level 1
Add-DocxParagraph -Doc $doc -Text 'This document is generated from a read-only Microsoft Graph snapshot of the Intune tenant. Settings tables reflect the tenant exactly as captured; design notes (purpose, owner, dependencies, lab validation and rollback) are maintained separately by the endpoint team and merged in at build time.'
Add-DocxTable -Doc $doc -Headers 'Snapshot', 'Value' -Widths 0.3, 0.7 -Rows @(
    , @('Tenant', $tenantName)
    , @('Tenant ID', $meta.tenantId)
    , @('Captured (UTC)', $collectedAt)
    , @('Captured by', $meta.collectedBy)
    , @('Snapshot folder', (Split-Path $SnapshotPath -Leaf))
    , @('Collector version', $meta.collectorVersion)
    , @('Script content included', (Format-Scalar ([bool]$meta.options.includeScriptContent)))
)
if (@($meta.warnings).Count) {
    Add-DocxParagraph -Doc $doc -Text 'Collection warnings - the following areas could not be read and may be incomplete:' -Bold -Color 'C00000'
    Add-DocxTable -Doc $doc -Headers 'Warning' -Rows @($meta.warnings | ForEach-Object { , @($_) })
}

Add-DocxHeading -Doc $doc -Text 'Summary' -Level 2
Add-DocxTable -Doc $doc -Headers 'Area', 'Objects' -Widths 0.7, 0.3 -Rows @(
    , @('Autopilot deployment profiles (v1)', $autopilotProfiles.Count)
    , @('Autopilot device preparation policies (v2)', $apdpPolicies.Count)
    , @('Enrollment Status Page profiles', $espConfigs.Count)
    , @('Settings catalog policies', $catalogPolicies.Count)
    , @('Endpoint security policies', $esPolicies.Count)
    , @('Security baselines', ($baselinePolicies.Count + $intents.Count))
    , @('Configuration profiles (templates)', $templates.Count)
    , @('Administrative templates', $gpConfigs.Count)
    , @('Compliance policies', $complianceIn.Count)
    , @('Windows Update rings and profiles', ($updateRings.Count + $featureUpdates.Count + $qualityUpdates.Count + $driverUpdates.Count))
    , @('Apps', $appsIn.Count)
    , @('Platform scripts', $scriptsIn.Count)
    , @('Remediations', $remIn.Count)
)

# 2. Windows enrollment
Add-DocxHeading -Doc $doc -Text 'Windows enrollment' -Level 1
Write-ObjectList -Doc $doc -Heading 'Autopilot deployment profiles (v1)' -Objects $autopilotProfiles -Category 'Autopilot profile' `
    -TypeLabel { param($o) if ($o.'@odata.type' -like '*activeDirectory*') { 'Autopilot profile - Hybrid Entra join' } else { 'Autopilot profile - Entra join' } } -Rows $propRows
Write-ObjectList -Doc $doc -Heading 'Autopilot device preparation policies (v2)' -Objects $apdpPolicies -Category 'Autopilot device preparation' `
    -TypeLabel { param($o) 'Autopilot device preparation' } -Rows $catalogRows
Add-DocxParagraph -Doc $doc -Text 'Note: confirm the device preparation device security group in the admin center; it is not exposed with the policy settings above.' -Italic -Color '808080'
Write-ObjectList -Doc $doc -Heading 'Enrollment Status Page' -Objects $espConfigs -Category 'Enrollment Status Page' -TypeLabel { param($o) 'Enrollment Status Page' } -Rows $propRows
Write-ObjectList -Doc $doc -Heading 'Other enrollment configurations' -Objects $otherEnroll -Category 'Enrollment configuration' -TypeLabel $typeLabel -Rows $propRows

if ($autopilotDevices.Count) {
    Add-DocxHeading -Doc $doc -Text 'Autopilot device register' -Level 2
    Add-DocxParagraph -Doc $doc -Text "$($autopilotDevices.Count) device(s) registered at snapshot time."
    Add-DocxTable -Doc $doc -Headers 'Group tag', 'Devices' -Widths 0.7, 0.3 -Rows @(
        $autopilotDevices | Group-Object { if ($_.groupTag) { $_.groupTag } else { '(none)' } } | Sort-Object Name | ForEach-Object { , @($_.Name, $_.Count) })
    Add-DocxTable -Doc $doc -Headers 'Manufacturer / model', 'Devices' -Widths 0.7, 0.3 -Rows @(
        $autopilotDevices | Group-Object { "$($_.manufacturer) $($_.model)".Trim() } | Sort-Object Count -Descending | ForEach-Object { , @($_.Name, $_.Count) })
}

# 3. Policies
Add-DocxHeading -Doc $doc -Text 'Device configuration and policies' -Level 1
Write-ObjectList -Doc $doc -Heading 'Settings catalog' -Objects $catalogPolicies -Category 'Settings catalog' -TypeLabel $templateLabel -Rows $catalogRows
Write-ObjectList -Doc $doc -Heading 'Endpoint security' -Objects $esPolicies -Category 'Endpoint security' `
    -TypeLabel { param($o) "$(Format-Label (([string]$o.templateReference.templateFamily) -replace '^endpointSecurity', '')): $($o.templateReference.templateDisplayName)" } -Rows $catalogRows
Write-ObjectList -Doc $doc -Heading 'Security baselines' -Objects $baselinePolicies -Category 'Security baseline' -TypeLabel $templateLabel -Rows $catalogRows
if ($intents.Count) {
    Write-ObjectList -Doc $doc -Heading 'Legacy template policies (intents)' -Objects $intents -Category 'Intent' `
        -TypeLabel { param($o) 'Legacy endpoint security / baseline template' } -Rows { param($o) Get-IntentRows $o }
}
Write-ObjectList -Doc $doc -Heading 'Configuration profiles (templates)' -Objects $templates -Category 'Configuration profile' -TypeLabel $typeLabel -Rows $propRows
Write-ObjectList -Doc $doc -Heading 'Administrative templates' -Objects $gpConfigs -Category 'Administrative template' `
    -TypeLabel { param($o) 'Administrative templates (ADMX)' } -Rows { param($o) Get-AdmxRows $o } `
    -Headers 'Setting', 'Scope', 'Value' -Widths 0.5, 0.12, 0.38
Write-ObjectList -Doc $doc -Heading 'Compliance policies' -Objects $complianceIn -Category 'Compliance policy' -TypeLabel $typeLabel -Rows { param($o) Get-ComplianceRows $o }
Write-ObjectList -Doc $doc -Heading 'Windows Update rings' -Objects $updateRings -Category 'Update ring' -TypeLabel { param($o) 'Update ring (WUfB)' } -Rows $propRows
Write-ObjectList -Doc $doc -Heading 'Feature update profiles' -Objects $featureUpdates -Category 'Feature update' -TypeLabel { param($o) 'Feature update profile' } -Rows $propRows
Write-ObjectList -Doc $doc -Heading 'Quality update profiles' -Objects $qualityUpdates -Category 'Quality update' -TypeLabel { param($o) 'Quality update profile' } -Rows $propRows
Write-ObjectList -Doc $doc -Heading 'Driver update profiles' -Objects $driverUpdates -Category 'Driver update' -TypeLabel { param($o) 'Driver update profile' } -Rows $propRows

Add-DocxHeading -Doc $doc -Text 'Assignment filters' -Level 2
$allFilters = Read-Snapshot 'assignmentFilters'
$filterObjs = @($allFilters | Where-Object { Test-InScope $_ })
if ($filterObjs.Count) {
    Add-DocxTable -Doc $doc -Headers 'Filter', 'Platform', 'Rule', 'Description' -Widths 0.22, 0.12, 0.46, 0.2 -Rows @(
        $filterObjs | ForEach-Object { , @($_.displayName, [string]$_.platform, $_.rule, $_.description) })
}
else { Add-DocxParagraph -Doc $doc -Text 'None configured.' -Italic -Color '808080' }

# 4. Apps & scripts
Add-DocxHeading -Doc $doc -Text 'Apps and scripts' -Level 1
Add-DocxHeading -Doc $doc -Text 'App assignment summary' -Level 2
$summary = [System.Collections.Generic.List[object]]::new()
foreach ($a in ($appsIn | Sort-Object displayName)) {
    $byIntent = @{ required = @(); available = @(); uninstall = @() }
    foreach ($as in @($a.assignments)) {
        if (-not $as -or -not $as.target) { continue }
        $ti = Get-TargetInfo $as.target
        $txt = if ($ti.Mode -eq 'Exclude') { "Excl: $($ti.Name)" } else { $ti.Name }
        $key = [string]$as.intent
        if ($key -like 'available*') { $key = 'available' }
        if ($byIntent.ContainsKey($key)) { $byIntent[$key] += $txt }
    }
    $summary.Add(@($a.displayName, (Format-TypeName $a.'@odata.type'), ($a.displayVersion ?? $a.productVersion ?? ''),
        ($byIntent.required -join "`n"), ($byIntent.available -join "`n"), ($byIntent.uninstall -join "`n")))
}
if ($summary.Count) {
    Add-DocxTable -Doc $doc -Headers 'App', 'Type', 'Version', 'Required', 'Available', 'Uninstall' -Widths 0.22, 0.13, 0.09, 0.2, 0.2, 0.16 -FontSize 16 -Rows $summary.ToArray()
}
else { Add-DocxParagraph -Doc $doc -Text 'None configured.' -Italic -Color '808080' }

Write-ObjectList -Doc $doc -Heading 'App details' -Objects $appsIn -Category 'App' -TypeLabel $typeLabel -Rows {
    param($o)
    $rows = ConvertTo-PropertyRows $o
    foreach ($r in @($o.relationships | Where-Object { $_ })) {
        $kind = Format-TypeName $r.'@odata.type'
        $rel = if ($r.dependencyType) { "dependency ($($r.dependencyType))" } elseif ($r.supersedenceType) { "supersedence ($($r.supersedenceType))" } else { '' }
        $rows.Add(@("$kind ($($r.targetType))", "$($r.targetDisplayName) $($r.targetDisplayVersion) $rel".Trim()))
    }
    , $rows
}
Write-ObjectList -Doc $doc -Heading 'Platform scripts' -Objects $scriptsIn -Category 'Platform script' -TypeLabel { param($o) 'Platform script (PowerShell)' } -Rows $propRows
Write-ObjectList -Doc $doc -Heading 'Remediations' -Objects $remIn -Category 'Remediation' -TypeLabel { param($o) 'Remediation' } -Rows $propRows

# Appendix: script bodies (only when the snapshot included them)
$withContent = @(
    $scriptsIn | Where-Object { $_.scriptContent } | ForEach-Object { [pscustomobject]@{ Name = $_.displayName; Part = 'Script'; B64 = $_.scriptContent } }
    $remIn | ForEach-Object {
        if ($_.detectionScriptContent)   { [pscustomobject]@{ Name = $_.displayName; Part = 'Detection'; B64 = $_.detectionScriptContent } }
        if ($_.remediationScriptContent) { [pscustomobject]@{ Name = $_.displayName; Part = 'Remediation'; B64 = $_.remediationScriptContent } }
    }
)
if ($withContent.Count) {
    Add-DocxHeading -Doc $doc -Text 'Appendix A - Script content' -Level 1
    foreach ($s in $withContent) {
        Add-DocxHeading -Doc $doc -Text "$($s.Name) - $($s.Part)" -Level 3
        $text = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($s.B64)).TrimStart([char]0xFEFF)
        Add-DocxParagraph -Doc $doc -Text $text -Mono -Size 16
    }
}
#endregion

#region Save
if (-not $OutputPath) {
    $OutputPath = Join-Path $ws.Output ("Intune-AsBuilt_{0}_{1}.docx" -f (($tenantName -replace '[^\w\-]', '')), (Split-Path $SnapshotPath -Leaf))
}
$saved = Save-AsBuiltDocument -Doc $doc -Path $OutputPath -TemplatePath $TemplatePath -Title "$Title - $tenantName" -FooterText "$Title - $tenantName - snapshot $collectedAt UTC"
Write-Host "As-built written to $saved" -ForegroundColor Green

if ($Pdf) {
    $pdfPath = [IO.Path]::ChangeExtension($saved, '.pdf')
    Write-Host 'Exporting PDF...' -ForegroundColor Cyan
    try {
        $engine = Export-AsBuiltPdf -DocxPath $saved -PdfPath $pdfPath -Engine $PdfEngine
        Write-Host "PDF written to $pdfPath (engine: $engine)" -ForegroundColor Green
        if ($engine -eq 'LibreOffice') {
            Write-Warning 'LibreOffice does not build the table of contents - the PDF shows the TOC placeholder. Use Word for a complete TOC.'
        }
    }
    catch { Write-Warning "PDF export failed - the .docx was still created. $($_.Exception.Message)" }
}

if ($ExportNotesTemplate) {
    $existing = @{}
    if (Test-Path -LiteralPath $ExportNotesTemplate) { foreach ($n in (Import-Csv -LiteralPath $ExportNotesTemplate)) { $existing[$n.ObjectId] = $n } }
    foreach ($k in $script:Notes.Keys) { if (-not $existing.ContainsKey($k)) { $existing[$k] = $script:Notes[$k] } }
    $seen = @{}
    $out = foreach ($i in ($script:Inventory | Sort-Object Category, Name)) {
        if ($seen[$i.ObjectId]) { continue }; $seen[$i.ObjectId] = $true
        $e = $existing[$i.ObjectId]
        [pscustomobject]@{
            ObjectId = $i.ObjectId; Category = $i.Category; Name = $i.Name
            Purpose = $e.Purpose; Owner = $e.Owner; DependsOn = $e.DependsOn; LastValidated = $e.LastValidated; Rollback = $e.Rollback
        }
    }
    # Keep notes for objects no longer in the tenant so nothing written by hand is lost
    $orphans = foreach ($k in $existing.Keys) {
        if (-not $seen[$k]) {
            $e = $existing[$k]
            [pscustomobject]@{ ObjectId = $k; Category = '(not in this snapshot)'; Name = $e.Name; Purpose = $e.Purpose; Owner = $e.Owner; DependsOn = $e.DependsOn; LastValidated = $e.LastValidated; Rollback = $e.Rollback }
        }
    }
    @($out) + @($orphans) | Export-Csv -LiteralPath $ExportNotesTemplate -NoTypeInformation -Encoding utf8
    Write-Host "Notes template written to $ExportNotesTemplate ($(@($out).Count) objects)" -ForegroundColor Green
}
#endregion
