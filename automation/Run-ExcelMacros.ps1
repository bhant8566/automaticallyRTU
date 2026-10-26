[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ReportRoot,
    [Parameter(Mandatory = $true)][string]$WorkbookPath,
    [uri]$NotifyUrl = 'http://localhost:5080/api/reports/refresh',
    [switch]$Hidden,
    [switch]$ShowExcel
)

$ErrorActionPreference = 'Stop'

# Source folders: change these to the actual locations.
$sourceRTU = 'D:\ReportSources\Source1'
$sourceOHTM = 'D:\ReportSources\Source2'
$sourceWendy = 'D:\ReportSources\Source3'
$sourceFilter = @('.csv', '.xlsx')

function Get-LatestSourceFile {
    param([string]$Folder, [string[]]$Filter)
    $file = Get-ChildItem -LiteralPath $Folder -File |
        Where-Object { $_.Extension -in $Filter -and $_.Name -notlike '~$*' } |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($null -eq $file) { throw "No matching source file in: $Folder" }
    return $file.FullName
}

$phase = 'Initialize'
$relatedPath = $null
$reportDate = $null
$reportFolder = $null
$source1 = $null
$source2 = $null
$source3 = $null
$primaryError = $null
$cleanupErrors = [Collections.Generic.List[object]]::new()

try {
    $taiwanNow = [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::UtcNow, 'Taipei Standard Time')
    $reportDate = $taiwanNow.AddDays(-1).ToString('yyyy-MM-dd')
    $reportFolder = Join-Path $ReportRoot $reportDate

    $phase = 'ReadSourceRTU'
    $relatedPath = $sourceRTU
    $source1 = Get-LatestSourceFile -Folder $sourceRTU -Filter $sourceFilter
    $phase = 'ReadSourceOHTM'
    $relatedPath = $sourceOHTM
    $source2 = Get-LatestSourceFile -Folder $sourceOHTM -Filter $sourceFilter
    $phase = 'ReadSourceWendy'
    $relatedPath = $sourceWendy
    $source3 = Get-LatestSourceFile -Folder $sourceWendy -Filter $sourceFilter

    $phase = 'CreateReportFolder'
    $relatedPath = $reportFolder
    New-Item -ItemType Directory -Path $reportFolder -Force | Out-Null
    $completeFlag = Join-Path $reportFolder 'complete'
    $phase = 'RemoveOldCompleteFlag'
    $relatedPath = $completeFlag
    if (Test-Path -LiteralPath $completeFlag) { Remove-Item -LiteralPath $completeFlag -Force }

    $excel = $null
    $workbooks = $null
    $workbook = $null
    try {
        $phase = 'StartExcel'
        $relatedPath = $WorkbookPath
        $excel = New-Object -ComObject Excel.Application
        # Hidden is retained for older callers. If both switches are set, Hidden wins.
        $excel.Visible = $ShowExcel.IsPresent -and -not $Hidden.IsPresent
        $workbooks = $excel.Workbooks
        $phase = 'OpenWorkbook'
        $workbook = $workbooks.Open($WorkbookPath, 0, $false)
        if ($workbook.ReadOnly) { throw 'Workbook is read-only. Check file permissions and other Excel instances.' }
        $macroPrefix = "'$($workbook.Name.Replace("'", "''"))'!"

        # Replace these module/procedure names with the macros in your workbook.
        $phase = 'ImportRTU'
        $relatedPath = $source1
        $excel.Run("${macroPrefix}Module1.ImportSource1", $source1) | Out-Null
        $phase = 'ImportOHTM'
        $relatedPath = $source2
        $excel.Run("${macroPrefix}Module1.ImportSource2", $source2) | Out-Null
        $phase = 'ImportWendy'
        $relatedPath = $source3
        $excel.Run("${macroPrefix}Module1.ImportSource3", $source3) | Out-Null
        $phase = 'GenerateReports'
        $relatedPath = $reportFolder
        $excel.Run("${macroPrefix}Module1.GenerateReports", $reportFolder) | Out-Null

        $phase = 'ValidateReports'
        foreach ($filename in @('main.html', 'no-reason.html', 'summary.html')) {
            $path = Join-Path $reportFolder $filename
            $relatedPath = $path
            if (-not (Test-Path -LiteralPath $path -PathType Leaf) -or (Get-Item -LiteralPath $path).Length -eq 0) {
                throw "Missing or empty report: $filename"
            }
        }
        $phase = 'SaveWorkbook'
        $relatedPath = $WorkbookPath
        $workbook.Save() | Out-Null
    }
    catch { $primaryError = $_ }
    finally {
        # Always attempt all cleanup steps, without overwriting the primary failure.
        if ($null -ne $workbook) {
            try { $workbook.Close($false) | Out-Null }
            catch { $cleanupErrors.Add([pscustomobject]@{ Stage = 'CloseWorkbook'; Error = $_ }) }
        }
        if ($null -ne $excel) {
            try { $excel.Quit() | Out-Null }
            catch { $cleanupErrors.Add([pscustomobject]@{ Stage = 'QuitExcel'; Error = $_ }) }
        }
        foreach ($object in @($workbook, $workbooks, $excel)) {
            try {
                if ($null -ne $object -and [Runtime.InteropServices.Marshal]::IsComObject($object)) {
                    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($object)
                }
            }
            catch { $cleanupErrors.Add([pscustomobject]@{ Stage = 'ReleaseComObject'; Error = $_ }) }
        }
    }
    if ($null -ne $primaryError) { throw $primaryError }
    if ($cleanupErrors.Count -gt 0) {
        $phase = $cleanupErrors[0].Stage
        $relatedPath = $WorkbookPath
        throw $cleanupErrors[0].Error
    }

    $phase = 'WriteCompleteFlag'
    $relatedPath = $completeFlag
    [IO.File]::WriteAllText($completeFlag, '')

    # Only source files are removed; report date folders are left untouched.
    $phase = 'CleanSourceFiles'
    foreach ($folder in @($sourceRTU, $sourceOHTM, $sourceWendy)) {
        $relatedPath = $folder
        $oldFiles = Get-ChildItem -LiteralPath $folder -File |
            Where-Object { $_.Extension -in $sourceFilter -and $_.Name -notlike '~$*' } |
            Sort-Object LastWriteTime -Descending | Select-Object -Skip 7
        foreach ($file in $oldFiles) {
            $relatedPath = $file.FullName
            Remove-Item -LiteralPath $file.FullName -Force
        }
    }

    # Notification is always the last operation after successful source cleanup.
    $phase = 'NotifyWebsite'
    $relatedPath = $NotifyUrl.GetLeftPart([UriPartial]::Path)
    try {
        $body = @{ reportDate = $reportDate } | ConvertTo-Json -Compress
        Invoke-RestMethod -Uri $NotifyUrl -Method Post -ContentType 'application/json' -Body $body -TimeoutSec 10 | Out-Null
    }
    catch {
        throw [InvalidOperationException]::new('Report generation and source cleanup succeeded, but website notification failed. Retry only the notification; do not rerun the macro.', $_.Exception)
    }
}
catch {
    $failure = $_
    try {
        $time = [TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTimeOffset]::UtcNow, 'Taipei Standard Time')
        $entry = [ordered]@{
            time = $time.ToString('o')
            stage = $phase
            reportDate = $reportDate
            reportFolder = $reportFolder
            workbook = $WorkbookPath
            relatedPath = $relatedPath
            sources = @{ RTU = $source1; OHTM = $source2; Wendy = $source3 }
            message = $failure.Exception.Message
            exception = $failure.Exception.ToString()
            location = $failure.InvocationInfo.PositionMessage
            stack = $failure.ScriptStackTrace
            cleanupErrors = @($cleanupErrors | ForEach-Object {
                @{ stage = $_.Stage; message = $_.Error.Exception.Message; exception = $_.Error.Exception.ToString(); stack = $_.Error.ScriptStackTrace }
            })
        }
        $record = $entry | ConvertTo-Json -Depth 6 -Compress
        $logDirectory = Join-Path $PSScriptRoot 'logs'
        [void][IO.Directory]::CreateDirectory($logDirectory)
        $logPath = Join-Path $logDirectory ('error-' + $time.ToString('yyyy-MM-dd') + '.log')
        [IO.File]::AppendAllText($logPath, $record + [Environment]::NewLine, [Text.UTF8Encoding]::new($false))
    }
    catch {
        [Console]::Error.WriteLine('Unable to write error log: ' + $_.Exception.Message)
        [Console]::Error.WriteLine('Original failure [' + $phase + ']: ' + $failure.Exception.ToString())
        foreach ($cleanupFailure in $cleanupErrors) {
            [Console]::Error.WriteLine('Cleanup failure [' + $cleanupFailure.Stage + ']: ' + $cleanupFailure.Error.Exception.ToString())
        }
    }
    exit 1
}
exit 0
