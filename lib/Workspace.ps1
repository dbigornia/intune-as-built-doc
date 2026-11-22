# Workspace resolution for IntuneAsBuilt.
# A workspace is a private folder (ideally a private git repo) that holds everything tenant-specific:
#   snapshots\   output\   notes.csv   template.docx|.dotx
# Keeping it outside the tool folder means tenant data never ends up in the public tool repo.

$script:NotesHeader = '"ObjectId","Category","Name","Purpose","Owner","DependsOn","LastValidated","Rollback"'

function Resolve-AsBuiltWorkspace {
    param([string]$Workspace, [Parameter(Mandatory)][string]$ToolRoot)

    if (-not $Workspace) { $Workspace = $env:INTUNEASBUILT_WORKSPACE }

    if (-not $Workspace) {
        Write-Warning ("No -Workspace or INTUNEASBUILT_WORKSPACE set - tenant data will be written to the tool folder ($ToolRoot). " +
            "Use a private workspace folder so it can't end up in the public repo (see README > Workspace).")
        return [pscustomobject]@{
            Root = $ToolRoot; IsToolFolder = $true
            Snapshots = (Join-Path $ToolRoot 'snapshots'); Output = $ToolRoot
            Notes = $null; Template = $null
        }
    }

    $root = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Workspace)
    $toolFull = [IO.Path]::GetFullPath($ToolRoot).TrimEnd('\', '/')
    if ([IO.Path]::GetFullPath($root).TrimEnd('\', '/').StartsWith($toolFull, [StringComparison]::OrdinalIgnoreCase)) {
        Write-Warning "The workspace is inside the tool folder ($toolFull). Put it in a separate, private location."
    }

    $snapshots = Join-Path $root 'snapshots'
    $output = Join-Path $root 'output'
    foreach ($d in $root, $snapshots, $output) {
        if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
    }

    $notes = Join-Path $root 'notes.csv'
    if (-not (Test-Path -LiteralPath $notes)) {
        Set-Content -LiteralPath $notes -Value $script:NotesHeader -Encoding utf8
        Write-Host "Created $notes"
    }

    $template = $null
    foreach ($name in 'template.docx', 'template.dotx') {
        $candidate = Join-Path $root $name
        if (Test-Path -LiteralPath $candidate) { $template = $candidate; break }
    }

    [pscustomobject]@{
        Root = $root; IsToolFolder = $false
        Snapshots = $snapshots; Output = $output
        Notes = $notes; Template = $template
    }
}
