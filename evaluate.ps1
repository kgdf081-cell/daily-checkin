param(
    [switch]$NoWrite,
    [switch]$Json,
    [string]$AsOfDate
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\run.ps1')

function Format-Rate {
    param($Value)
    if ($null -eq $Value) { return '—' }
    return ('{0:P0}' -f [double]$Value)
}

$asOf = if ($AsOfDate) { ([datetime]::Parse($AsOfDate)).Date } else { (Get-Date).Date }
$result = Invoke-CheckinEvaluate -AsOf $asOf -Write (-not $NoWrite)
$payload = Get-ReminderPayload $result

if ($Json) {
    $payload | ConvertTo-Json -Depth 6
    exit 0
}

if (-not $result.Ok) {
    Write-Host "评估失败：$($result.Error)" -ForegroundColor Red
    Write-Host "如果表格正在 Excel 中打开，请先保存并关闭，再运行。" -ForegroundColor Yellow
    exit 1
}

$s = $result.Stats
Write-Host ''
Write-Host "  每日打卡 · $($asOf.ToString('yyyy-MM-dd'))" -ForegroundColor Cyan
Write-Host ('  ' + ('-' * 52))
Write-Host ("  今日：应做 {0} / 完成 {1} / 补签 {2} / 待办 {3}" -f $s.Today.Due, $s.Today.Done, $s.Today.Patch, $s.Today.Pending)
Write-Host ("  昨日：原始率 {0} / 含补签 {1} / 加权 {2}" -f (Format-Rate $s.Yesterday.RawRate), (Format-Rate $s.Yesterday.PatchRate), (Format-Rate $s.Yesterday.WeightedRate))
Write-Host ("  近7天：原始率 {0} / 含补签 {1} / 加权 {2}" -f (Format-Rate $s.Week.RawRate), (Format-Rate $s.Week.PatchRate), (Format-Rate $s.Week.WeightedRate))
Write-Host ''
Write-Host ("  稳定 {0} · 正常 {1} · 懈怠 {2} · 基本没做 {3} · 待观察 {4}" -f $s.Counts['稳定坚持'], $s.Counts['正常'], $s.Counts['懈怠'], $s.Counts['基本没做'], $s.Counts['待观察'])
if (@($s.PendingList).Count -gt 0) { Write-Host ("  待补签：" + (Get-ListText $s.PendingList 6)) -ForegroundColor Yellow }
if (@($s.ExpiringList).Count -gt 0) { Write-Host ("  即将超期：" + (Get-ListText $s.ExpiringList 6)) -ForegroundColor Yellow }
if (@($s.LockedList).Count -gt 0) { Write-Host ("  超期锁定：" + (Get-ListText $s.LockedList 6)) -ForegroundColor DarkYellow }
if (@($s.SlackList).Count -gt 0) { Write-Host ("  懈怠预警：" + (Get-ListText $s.SlackList 6)) -ForegroundColor Yellow }
if (@($s.NoneList).Count -gt 0) { Write-Host ("  基本没做：" + (Get-ListText $s.NoneList 6)) -ForegroundColor Red }
Write-Host ''
if ($result.Written) { Write-Host '  已刷新表格与评估页签。' -ForegroundColor Green }
else { Write-Host ("  未写回：" + $result.WriteReason) -ForegroundColor DarkGray }
Write-Host ''
