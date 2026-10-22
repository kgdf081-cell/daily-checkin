param([switch]$NoOpen)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\run.ps1')

$result = $null
try {
    $result = Invoke-CheckinEvaluate -Write $true
}
catch {
    Write-Log "打开表格失败：$($_.Exception.Message)" 'ERROR'
}

if (-not $result -or -not $result.Ok) {
    if ((Test-Path -LiteralPath $script:DataFile) -and (-not (Test-WorkbookWritable))) {
        # 表格正被 Excel 打开：直接激活/打开它，避免误报错误。
        try {
            Start-Process -FilePath $script:DataFile
            exit 0
        }
        catch { }
    }
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "打卡表读取失败，请先查看日志：`n$($script:LogFile)`n`n$($result.Error)",
        '每日打卡', 'OK', 'Error') | Out-Null
    exit 1
}

if (-not $NoOpen) {
    try {
        Start-Process -FilePath $result.Path
    }
    catch {
        Start-Process -FilePath (Split-Path $result.Path -Parent)
    }
}
