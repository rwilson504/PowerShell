<#
.SYNOPSIS
    Downloads all note attachments directly related to a Dataverse record.

.DESCRIPTION
    Queries the Dataverse annotation table for notes whose object lookup matches the
    supplied record GUID and whose isdocument value is true. Each attachment's base64
    documentbody is decoded and written to the output directory.

    The script follows OData paging, sanitizes filenames, and gives duplicate filenames
    a suffix based on the annotation ID. Existing files are skipped unless -Overwrite is
    specified. One result object is returned for every attachment found.

    This script downloads note (annotation) attachments directly related to the record.
    It does not include file/image column data or attachments on related email activities.

.PARAMETER OrganizationUrl
    The URL of the Dataverse organization (for example,
    https://your-org.crm.dynamics.com).

.PARAMETER AccessToken
    The access token for authenticating with the Dataverse Web API.

.PARAMETER RecordId
    The GUID of the record whose note attachments should be downloaded.

.PARAMETER OutputPath
    Directory where attachments are written. Defaults to a folder named
    DataverseAttachments-<RecordId> beneath the current directory.

.PARAMETER Overwrite
    Replace files that already exist at the resolved output path. By default, existing
    files are skipped.

.EXAMPLE
    .\Save-RecordAttachments.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -RecordId "00000000-0000-0000-0000-000000000001"

    Downloads all note attachments for the record into the default output directory.

.EXAMPLE
    .\Save-RecordAttachments.ps1 -OrganizationUrl "https://your-org.crm.dynamics.com" -AccessToken $token -RecordId "00000000-0000-0000-0000-000000000001" -OutputPath "C:\Exports\CaseAttachments" -Overwrite

    Downloads the attachments to a specific directory and replaces existing files.
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory = $true)]
    [string]$OrganizationUrl,

    [Parameter(Mandatory = $true)]
    [string]$AccessToken,

    [Parameter(Mandatory = $true)]
    [guid]$RecordId,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = (Join-Path (Get-Location).Path "DataverseAttachments-$RecordId"),

    [Parameter(Mandatory = $false)]
    [switch]$Overwrite
)

$OrganizationUrl = $OrganizationUrl.TrimEnd('/')
$resolvedOutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)

$headers = @{
    "Authorization"    = "Bearer $AccessToken"
    "Accept"           = "application/json"
    "OData-MaxVersion" = "4.0"
    "OData-Version"    = "4.0"
}

function ConvertTo-SafeFileName {
    param (
        [string]$FileName,
        [guid]$AnnotationId
    )

    $safeName = [System.IO.Path]::GetFileName($FileName)
    if ([string]::IsNullOrWhiteSpace($safeName)) {
        $safeName = "attachment-$AnnotationId"
    }

    $invalidCharacters = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($character in $invalidCharacters) {
        $safeName = $safeName.Replace([string]$character, '_')
    }

    $safeName = $safeName.Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($safeName)) {
        return "attachment-$AnnotationId"
    }

    return $safeName
}

function Get-UniqueAttachmentPath {
    param (
        [string]$Directory,
        [string]$FileName,
        [guid]$AnnotationId,
        [System.Collections.Generic.HashSet[string]]$ReservedPaths
    )

    $candidatePath = Join-Path $Directory $FileName
    if ($ReservedPaths.Add($candidatePath)) {
        return $candidatePath
    }

    $extension = [System.IO.Path]::GetExtension($FileName)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $shortId = $AnnotationId.ToString('N').Substring(0, 8)
    $candidatePath = Join-Path $Directory "$baseName-$shortId$extension"
    $suffix = 2

    while (-not $ReservedPaths.Add($candidatePath)) {
        $candidatePath = Join-Path $Directory "$baseName-$shortId-$suffix$extension"
        $suffix++
    }

    return $candidatePath
}

$recordIdFilter = $RecordId.ToString('D')
$select = 'annotationid,filename,mimetype,filesize,subject,createdon,documentbody'
$url = "$OrganizationUrl/api/data/v9.2/annotations?`$select=$select&`$filter=_objectid_value eq $recordIdFilter and isdocument eq true"
$attachments = New-Object System.Collections.Generic.List[object]

Write-Host "Finding note attachments for record $recordIdFilter..." -ForegroundColor Cyan

do {
    $response = Invoke-RestMethod -Method Get -Uri $url -Headers $headers
    foreach ($attachment in $response.value) {
        $attachments.Add($attachment) | Out-Null
    }
    $url = $response.'@odata.nextLink'
} while ($url)

if ($attachments.Count -eq 0) {
    Write-Host "No note attachments were found." -ForegroundColor Yellow
    return @()
}

if (-not (Test-Path -LiteralPath $resolvedOutputPath)) {
    if ($PSCmdlet.ShouldProcess($resolvedOutputPath, 'Create attachment output directory')) {
        New-Item -ItemType Directory -Path $resolvedOutputPath -Force | Out-Null
    }
}

$reservedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$results = New-Object System.Collections.Generic.List[object]

foreach ($attachment in $attachments) {
    $annotationId = [guid]$attachment.annotationid
    $safeFileName = ConvertTo-SafeFileName -FileName $attachment.filename -AnnotationId $annotationId
    $filePath = Get-UniqueAttachmentPath -Directory $resolvedOutputPath -FileName $safeFileName -AnnotationId $annotationId -ReservedPaths $reservedPaths
    $resolvedFileName = Split-Path -Path $filePath -Leaf
    $status = 'Downloaded'
    $errorMessage = $null

    try {
        if ((Test-Path -LiteralPath $filePath) -and -not $Overwrite) {
            $status = 'SkippedExisting'
            Write-Warning "Skipping existing file '$filePath'. Use -Overwrite to replace it."
        }
        elseif ($PSCmdlet.ShouldProcess($filePath, 'Write Dataverse note attachment')) {
            if ([string]::IsNullOrWhiteSpace([string]$attachment.documentbody)) {
                throw "The attachment has no documentbody content."
            }

            $bytes = [Convert]::FromBase64String([string]$attachment.documentbody)
            [System.IO.File]::WriteAllBytes($filePath, $bytes)
            Write-Host "Downloaded: $resolvedFileName" -ForegroundColor Green
        }
        else {
            $status = 'NotWritten'
        }
    }
    catch {
        $status = 'Error'
        $errorMessage = $_.Exception.Message
        Write-Warning "Failed to download attachment '$resolvedFileName': $errorMessage"
    }

    $results.Add([PSCustomObject]@{
        RecordId     = $recordIdFilter
        AnnotationId = $annotationId
        FileName     = $resolvedFileName
        OriginalName = $attachment.filename
        MimeType     = $attachment.mimetype
        FileSize     = $attachment.filesize
        Subject      = $attachment.subject
        CreatedOn    = $attachment.createdon
        OutputPath   = $filePath
        Status       = $status
        Error        = $errorMessage
    }) | Out-Null
}

Write-Host "Processed $($attachments.Count) attachment(s)." -ForegroundColor Cyan
return $results.ToArray()
