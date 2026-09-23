# The Accurate routine on the laptop, the way it is done by hand:
#   1. download prev + current month from Accurate into "Accurate Report 2025\"
#      as "<MM>. BBL Aura WhiteInc Sales.xlsx" (old copies backed up first)
#   2. refresh "Accurate 2026.xlsx" in Excel (its Power Query reads those files)
#   3. refresh "0. 2025 Accurate.xlsx" in Excel (Accurate 2025 + Accurate 2026)
# The workbooks keep their Power Query - Excel refreshes them, nothing rewrites them.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File local\run_accurate_local.ps1
param([string[]] $Month = @())

$ErrorActionPreference = "Stop"
$repo   = Split-Path -Parent $PSScriptRoot
$folder = "D:\SCARLETT_512\SCARLETT-329\SOM\Data\Report\Sales\Accurate Report 2025"
$python = "D:\SCARLETT_512\SCARLETT-329\SOM\GitHub-Automation\som-sell-in-report-automation\.venv\Scripts\python.exe"
$backup = Join-Path $repo "work\local_backups"
$books  = @("$folder\Accurate 2026.xlsx", "$folder\0. 2025 Accurate.xlsx")   # order matters
$env:PYTHONIOENCODING = "utf-8"

# A scheduled run has no window, so keep a dated log (last 30 days).
$logDir = Join-Path $repo "work\logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
Get-ChildItem $logDir -Filter "run_*.log" | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } | Remove-Item -Force
Start-Transcript -Path (Join-Path $logDir ("run_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))) | Out-Null

# --- 1. download ----------------------------------------------------------
$pyArgs = @("$repo\local\download_to_folder.py")
foreach ($m in ($Month | ForEach-Object { $_ -split "[,;]" } | Where-Object { $_ })) { $pyArgs += @("--month", $m.Trim()) }
& $python @pyArgs
if ($LASTEXITCODE -ne 0) { "Download failed (exit $LASTEXITCODE) - workbooks not refreshed."; Stop-Transcript | Out-Null; exit 1 }

# --- 2+3. refresh in Excel ------------------------------------------------
New-Item -ItemType Directory -Force -Path $backup | Out-Null
Get-ChildItem $backup -File | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } | Remove-Item -Force
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
foreach ($b in $books) { Copy-Item $b (Join-Path $backup ("{0} {1}.xlsx" -f [IO.Path]::GetFileNameWithoutExtension($b), $stamp)) }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
$failed = $false
try {
    foreach ($path in $books) {
        $name = Split-Path $path -Leaf
        "$(Get-Date -Format 'HH:mm:ss')  Refreshing $name"
        $wb = $excel.Workbooks.Open($path, 0, $false)
        # RefreshAll returns while a background query is still running, which
        # would save a half-refreshed book - make every connection synchronous.
        foreach ($conn in $wb.Connections) {
            try {
                if ($conn.Type -eq 1) { $conn.OLEDBConnection.BackgroundQuery = $false }
                elseif ($conn.Type -eq 2) { $conn.ODBCConnection.BackgroundQuery = $false }
            } catch { }
        }
        $wb.RefreshAll()
        $excel.CalculateUntilAsyncQueriesDone()
        $rows = 0
        foreach ($ws in $wb.Worksheets) { foreach ($lo in $ws.ListObjects) { $rows += $lo.ListRows.Count } }
        $wb.Save()
        $wb.Close($true)
        "$(Get-Date -Format 'HH:mm:ss')  Refreshed $name - $('{0:N0}' -f $rows) rows"
    }
}
catch {
    $failed = $true
    "$(Get-Date -Format 'HH:mm:ss')  ERROR $($_.Exception.Message)"
}
finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
Stop-Transcript | Out-Null
if ($failed) { exit 1 }
exit 0
