param([int]$SetVolume = -1)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$configPath = Join-Path $root 'config.json'

. (Join-Path $root 'lib\engine.ps1')
. (Join-Path $root 'lib\sound.ps1')

$cfg = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
$current = if ($null -ne $cfg.voiceVolume) { [int]$cfg.voiceVolume } else { 70 }

Write-Host ''
Write-Host ('  当前语音音量：{0}%' -f $current) -ForegroundColor Cyan
Write-Host '  输入 0-100 的数字，回车保存并试听；直接回车保持不变。'
$input = if ($SetVolume -ge 0) { [string]$SetVolume } else { Read-Host '  新音量' }

$vol = $current
if ($input -match '^\s*\d+\s*$') {
    $vol = [int]$input
    if ($vol -lt 0) { $vol = 0 }
    if ($vol -gt 100) { $vol = 100 }
    $cfg.voiceVolume = $vol
    $json = $cfg | ConvertTo-Json -Depth 6
    [IO.File]::WriteAllText($configPath, $json, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host ('  已保存：{0}%' -f $vol) -ForegroundColor Green
}

$script:CheckinVoiceFile = if ($cfg.voiceFile) { [string]$cfg.voiceFile } else { 'voice.wav' }
$script:CheckinVoiceVolume = $vol
$script:CheckinSoundPlayed = $false
Write-Host '  正在试听...' -ForegroundColor Cyan
Play-CheckinSound
Start-Sleep -Seconds 4
Write-Host '  完成。' -ForegroundColor Green
