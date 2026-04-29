<#
.SYNOPSIS
    Shared helper for issuing OData $batch requests against the Dataverse Web API.

.DESCRIPTION
    Dot-source this file to gain access to Invoke-ODataBatch, which posts a
    multipart/mixed batch of GET sub-requests in a single HTTP round trip and
    returns the parsed JSON body of each response, in request order.

    A Dataverse $batch counts as ONE request against the per-user
    service-protection limit (6,000 requests / 5 min by default), so batching is
    the most effective way to keep call volume well below the throttling
    threshold while still issuing many logical queries.

.NOTES
    The helper is read-only by design - it only emits GET sub-requests.
#>

function Invoke-ODataBatch {
    <#
    .SYNOPSIS
        Posts a multipart/mixed OData $batch containing GET sub-requests and returns
        the parsed JSON body of each response, in request order. Failed sub-requests
        return $null in their slot.

    .PARAMETER OrgUrl
        The base Dataverse organization URL (no trailing slash, no /api/data/v9.2).

    .PARAMETER Headers
        Hashtable containing the Authorization header and any other request-level
        headers (Accept, OData-Version, etc.). The Content-Type header is replaced
        internally with the multipart boundary content type.

    .PARAMETER RelativeRequests
        Array of OData URLs RELATIVE to /api/data/v9.2/ (no leading slash).
        Example: "accounts?`$filter=statecode eq 0&`$count=true&`$top=1"
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$OrgUrl,

        [Parameter(Mandatory = $true)]
        [hashtable]$Headers,

        [Parameter(Mandatory = $true)]
        [string[]]$RelativeRequests
    )

    if (-not $RelativeRequests -or $RelativeRequests.Count -eq 0) {
        return @()
    }

    $boundary = "batch_$([guid]::NewGuid().ToString('N'))"
    $sb = [System.Text.StringBuilder]::new()
    foreach ($req in $RelativeRequests) {
        # CRLF line endings are required by RFC 2046 / OData batch spec
        [void]$sb.Append("--$boundary`r`n")
        [void]$sb.Append("Content-Type: application/http`r`n")
        [void]$sb.Append("Content-Transfer-Encoding: binary`r`n")
        [void]$sb.Append("`r`n")
        [void]$sb.Append("GET /api/data/v9.2/$req HTTP/1.1`r`n")
        [void]$sb.Append("Accept: application/json`r`n")
        [void]$sb.Append("OData-Version: 4.0`r`n")
        [void]$sb.Append("OData-MaxVersion: 4.0`r`n")
        [void]$sb.Append("`r`n")
    }
    [void]$sb.Append("--$boundary--`r`n")
    $body = $sb.ToString()

    # Build a request-specific header set: keep auth + odata headers, swap Content-Type
    $batchHeaders = @{}
    foreach ($k in $Headers.Keys) {
        if ($k -ne 'Content-Type') { $batchHeaders[$k] = $Headers[$k] }
    }
    $batchHeaders['Content-Type'] = "multipart/mixed; boundary=$boundary"

    try {
        $resp = Invoke-WebRequest -Uri "$OrgUrl/api/data/v9.2/`$batch" `
            -Method Post -Headers $batchHeaders -Body $body -UseBasicParsing
    }
    catch {
        # Trim to first 200 chars to keep warning readable (the exception body can include
        # the entire response payload, which is very noisy). Use Verbose so callers that do
        # internal retry logic (e.g. binary-split) don't spam the console.
        $errMsg = ($_.Exception.Message -replace '\r?\n', ' ').Trim()
        if ($errMsg.Length -gt 200) { $errMsg = $errMsg.Substring(0, 200) + '...' }
        Write-Verbose "OData batch request failed (will return null per request): $errMsg"
        return @($null) * $RelativeRequests.Count
    }

    # PowerShell 7's Invoke-WebRequest returns Content as a byte[]; Windows PowerShell
    # returns a string. Decode UTF-8 explicitly when needed.
    $respText = if ($resp.Content -is [byte[]]) {
        [System.Text.Encoding]::UTF8.GetString($resp.Content)
    }
    else {
        [string]$resp.Content
    }

    # Discover the response boundary from Content-Type
    $respCT = $resp.Headers['Content-Type']
    if ($respCT -is [array]) { $respCT = $respCT[0] }
    $respBoundary = $null
    if ($respCT -match 'boundary=([^;]+)') {
        $respBoundary = $matches[1].Trim('"')
    }
    if (-not $respBoundary) {
        Write-Warning "Could not parse response boundary from batch response."
        return @($null) * $RelativeRequests.Count
    }

    # Split into parts. The leading boundary marker plus trailing terminator are non-content.
    $parts = $respText -split [regex]::Escape("--$respBoundary")

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($part in $parts) {
        $trimmed = $part.Trim()
        if ($trimmed -eq '' -or $trimmed -eq '--') { continue }

        # Each part body looks like:
        #   Content-Type: application/http
        #   Content-Transfer-Encoding: binary
        #
        #   HTTP/1.1 200 OK
        #   Header: value
        #   ...
        #
        #   {"@odata.context":"...","@odata.count":150,"value":[...]}
        #
        # Extract the JSON between the first '{' and the last '}' of the part.
        $startIdx = $part.IndexOf('{')
        $endIdx   = $part.LastIndexOf('}')
        if ($startIdx -lt 0 -or $endIdx -lt $startIdx) {
            $results.Add($null)
            continue
        }
        $json = $part.Substring($startIdx, $endIdx - $startIdx + 1)
        try {
            $obj = $json | ConvertFrom-Json -ErrorAction Stop
            $results.Add($obj)
        }
        catch {
            $results.Add($null)
        }
    }

    if ($results.Count -ne $RelativeRequests.Count) {
        Write-Warning "Batch response part count ($($results.Count)) does not match request count ($($RelativeRequests.Count))."
    }

    return $results.ToArray()
}
