$ErrorActionPreference = 'Continue'
$root = $PSScriptRoot
$out = Join-Path $root '诊断结果.txt'
$lines = New-Object System.Collections.Generic.List[string]

$lines.Add('检查时间：' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
$lines.Add('')
$lines.Add('=== 计划任务 ===')
$q = & schtasks.exe /Query /TN '每日打卡提醒' /V /FO LIST 2>&1
if ($LASTEXITCODE -ne 0) {
    $lines.Add('未找到任务“每日打卡提醒”。')
    $lines.Add('请双击“安装提醒任务.cmd”重新安装。')
    $lines.Add(($q | Out-String).Trim())
}
else {
    $lines.Add(($q | Out-String).Trim())
}

$lines.Add('')
$lines.Add('=== 最近运行日志 ===')
$log = Join-Path $root 'log.txt'
if (Test-Path -LiteralPath $log) {
    $lines.Add(((Get-Content -LiteralPath $log -Encoding UTF8 -Tail 40) -join "`r`n"))
}
else {
    $lines.Add('log.txt 不存在')
}

$lines.Add('')
$lines.Add('=== 提醒状态 ===')
$state = Join-Path $root 'state.json'
if (Test-Path -LiteralPath $state) {
    $lines.Add((Get-Content -LiteralPath $state -Raw -Encoding UTF8))
}
else {
    $lines.Add('state.json 不存在')
}

$text = $lines -join "`r`n"
[IO.File]::WriteAllText($out, $text, (New-Object System.Text.UTF8Encoding($false)))
Write-Host $text
Write-Host ''
Write-Host ("结果已保存到：{0}" -f $out) -ForegroundColor Green
