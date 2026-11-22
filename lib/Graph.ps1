# Graph helpers for IntuneAsBuilt. Requires Microsoft.Graph.Authentication (Connect-MgGraph / Invoke-MgGraphRequest).

$script:GraphBase = 'https://graph.microsoft.com'

function Invoke-AsBuiltGraph {
    <#
      GET a Graph path ('beta/deviceManagement/...') or an absolute nextLink, with retry on throttling/transient errors.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$MaxRetries = 5
    )
    $uri = if ($Path -match '^https://') { $Path } else { "$script:GraphBase/$($Path.TrimStart('/'))" }
    for ($attempt = 1; ; $attempt++) {
        try {
            return Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject -ErrorAction Stop
        }
        catch {
            $status = $null
            try { $status = [int]$_.Exception.Response.StatusCode } catch { }
            $transient = ($null -eq $status) -or ($status -in 429, 500, 502, 503, 504)
            if ($attempt -lt $MaxRetries -and $transient) {
                $delay = [math]::Min(60, [math]::Pow(2, $attempt))
                Write-Verbose "Graph returned $status for $uri - retry $attempt in $delay s"
                Start-Sleep -Seconds $delay
                continue
            }
            throw
        }
    }
}

function Get-AsBuiltGraphCollection {
    <#
      GET a collection and follow @odata.nextLink until all pages are read. Always returns an array.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    $items = [System.Collections.Generic.List[object]]::new()
    $next = $Path
    while ($next) {
        $resp = Invoke-AsBuiltGraph -Path $next
        if ($resp.PSObject.Properties['value']) {
            foreach ($i in @($resp.value)) { if ($null -ne $i) { $items.Add($i) } }
            $next = $resp.'@odata.nextLink'
        }
        else {
            $items.Add($resp)
            $next = $null
        }
    }
    return , $items.ToArray()
}
