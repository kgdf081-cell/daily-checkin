# 提醒音：音量缩放（0-100）与播放

function Get-ScaledWav {
    param([string]$Source, [int]$Volume)
    $vol = [Math]::Max(0, [Math]::Min(100, $Volume))
    if ($vol -ge 100) { return $Source }
    if (-not (Test-Path -LiteralPath $Source)) { return $null }

    $cacheDir = Join-Path $env:TEMP 'DailyCheckin'
    if (-not (Test-Path -LiteralPath $cacheDir)) {
        New-Item -ItemType Directory -Force -Path $cacheDir | Out-Null
    }
    $base = [IO.Path]::GetFileNameWithoutExtension($Source)
    $cache = Join-Path $cacheDir ("{0}-v{1}.wav" -f $base, $vol)
    if ((Test-Path -LiteralPath $cache) -and ((Get-Item -LiteralPath $cache).LastWriteTime -ge (Get-Item -LiteralPath $Source).LastWriteTime)) {
        return $cache
    }

    $bytes = [IO.File]::ReadAllBytes($Source)
    $pos = 12
    $fmtPos = -1
    $fmtSize = 0
    $dataStart = -1
    $dataSize = 0
    while ($pos -le $bytes.Length - 8) {
        $id = [Text.Encoding]::ASCII.GetString($bytes, $pos, 4)
        $size = [BitConverter]::ToInt32($bytes, $pos + 4)
        if ($id -eq 'fmt ') {
            $fmtPos = $pos
            $fmtSize = $size
        }
        if ($id -eq 'data') {
            $dataStart = $pos + 8
            $dataSize = $size
            break
        }
        $pos += 8 + $size + ($size % 2)
    }
    if ($fmtPos -lt 0 -or $dataStart -lt 0 -or $dataSize -le 0) { return $Source }

    $audioFormat = [BitConverter]::ToInt16($bytes, $fmtPos + 8)
    $bits = [BitConverter]::ToInt16($bytes, $fmtPos + 22)
    if ($audioFormat -ne 1 -or $bits -ne 16) { return $Source }

    $sampleCount = [int]($dataSize / 2)
    $outData = New-Object byte[] $dataSize
    $factor = $vol / 100.0
    for ($i = 0; $i -lt $sampleCount; $i++) {
        $s = [BitConverter]::ToInt16($bytes, $dataStart + 2 * $i)
        $v = [int]($s * $factor)
        if ($v -gt 32767) { $v = 32767 }
        if ($v -lt -32768) { $v = -32768 }
        $outData[2 * $i] = [byte]($v -band 0xFF)
        $outData[2 * $i + 1] = [byte](($v -shr 8) -band 0xFF)
    }

    $fmtBytes = New-Object byte[] (8 + $fmtSize)
    [Array]::Copy($bytes, $fmtPos, $fmtBytes, 0, 8 + $fmtSize)
    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)
    try {
        $bw.Write([Text.Encoding]::ASCII.GetBytes('RIFF'))
        $bw.Write([int](4 + $fmtBytes.Length + 8 + $dataSize))
        $bw.Write([Text.Encoding]::ASCII.GetBytes('WAVE'))
        $bw.Write($fmtBytes)
        $bw.Write([Text.Encoding]::ASCII.GetBytes('data'))
        $bw.Write([int]$dataSize)
        $bw.Write($outData)
        $bw.Flush()
        [IO.File]::WriteAllBytes($cache, $ms.ToArray())
    }
    finally {
        $bw.Dispose()
        $ms.Dispose()
    }
    return $cache
}

function Play-CheckinSound {
    if ($script:CheckinSoundPlayed) { return }
    $script:CheckinSoundPlayed = $true
    $root = $script:RootDir
    if (-not $root) { $root = Split-Path $PSScriptRoot -Parent }
    $voiceName = if ($script:CheckinVoiceFile) { [string]$script:CheckinVoiceFile } else { 'voice.wav' }
    $voicePath = Join-Path $root $voiceName
    $dingPath = Join-Path $root 'ding.wav'
    $path = $null
    if (Test-Path -LiteralPath $voicePath) { $path = $voicePath }
    elseif (Test-Path -LiteralPath $dingPath) { $path = $dingPath }
    $vol = if ($null -ne $script:CheckinVoiceVolume) { [int]$script:CheckinVoiceVolume } else { 70 }
    $vol = [Math]::Max(0, [Math]::Min(100, $vol))
    try {
        if ($path) {
            $playPath = Get-ScaledWav -Source $path -Volume $vol
            $script:CheckinSoundPlayer = New-Object System.Media.SoundPlayer($playPath)
            $script:CheckinSoundPlayer.Play()
            Write-Log ("提醒音已播放：{0}（音量 {1}%）" -f (Split-Path $path -Leaf), $vol)
        }
        else {
            [System.Media.SystemSounds]::Asterisk.Play()
            Write-Log '未找到语音/提示音文件，改用系统提示音' 'WARN'
        }
    }
    catch {
        Write-Log "提示音播放失败：$($_.Exception.Message)" 'WARN'
    }
}
