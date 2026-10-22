# 读取 → 规范化 → 统计 → 写回 的统一入口。

. (Join-Path $PSScriptRoot 'engine.ps1')
. (Join-Path $PSScriptRoot 'render.ps1')
. (Join-Path $PSScriptRoot 'report.ps1')

function Invoke-CheckinEvaluate {
    param(
        [datetime]$AsOf = (Get-Date).Date,
        [bool]$Write = $true,
        [string]$Path = $script:DataFile
    )
    $config = Get-AppConfig
    $AsOf = $AsOf.Date
    $readOk = $true
    $readError = ''
    $model = $null
    $exists = Test-Path -LiteralPath $Path
    if ($exists) {
        try {
            $model = Read-Model -Path $Path -AsOf $AsOf
        }
        catch {
            $readOk = $false
            $readError = $_.Exception.Message
            Write-Log "读取表格失败：$readError" 'ERROR'
        }
    }
    else {
        $model = New-InitialModel -Today $AsOf
    }
    if (-not $model) {
        return [pscustomobject]@{
            Ok = $false; Error = $readError; Written = $false; Stats = $null; Model = $null; Config = $config; Path = $Path
        }
    }

    $changed = $false
    for ($i = 0; $i -lt $model.Items.Count; $i++) {
        $item = $model.Items[$i]
        if (-not $item.StartDate) {
            $earliest = $null
            foreach ($d in ($model.Dates | Sort-Object)) {
                if ($d -gt $AsOf) { continue }
                $key = "$i|$($d.ToString('yyyy-MM-dd'))"
                if ($model.Marks.ContainsKey($key)) { $earliest = $d; break }
            }
            if ($earliest) { $item.StartDate = $earliest } else { $item.StartDate = $AsOf }
            $changed = $true
        }
        if ($item.Status -eq $script:StatusDone) {
            if (-not $item.FinishDate) {
                $item.FinishDate = $AsOf
                $changed = $true
            }
            if ($item.FinishDate -gt $AsOf) {
                $item.FinishDate = $AsOf
                $changed = $true
            }
            if ($item.DeleteDate) { $item.DeleteDate = $null; $changed = $true }
        }
        elseif ($item.Status -eq $script:StatusDeleted) {
            if (-not $item.DeleteDate) {
                $item.DeleteDate = $AsOf
                $changed = $true
            }
            if ($item.DeleteDate -gt $AsOf) {
                $item.DeleteDate = $AsOf
                $changed = $true
            }
            if ($item.FinishDate) { $item.FinishDate = $null; $changed = $true }
        }
        else {
            if ($item.FinishDate) { $item.FinishDate = $null; $changed = $true }
            if ($item.DeleteDate) { $item.DeleteDate = $null; $changed = $true }
        }
        if ($item.Type -notin @('习惯', '任务')) { $item.Type = '习惯'; $changed = $true }
        if ($item.Weight -notin @('高', '中', '低')) { $item.Weight = '中'; $changed = $true }
    }

    $beforeCount = $model.Dates.Count
    Add-MissingDates -Model $model -Today $AsOf -Config $config | Out-Null
    if ($model.Dates.Count -ne $beforeCount) { $changed = $true }

    $stats = Get-AllStats -Model $model -AsOf $AsOf -Config $config
    try {
        Write-HtmlReport -Stats $stats -AsOf $AsOf -Path (Join-Path $script:RootDir 'report.html')
    }
    catch {
        Write-Log "生成 report.html 失败：$($_.Exception.Message)" 'WARN'
    }

    $written = $false
    $writeReason = ''
    if ($Write -and $readOk) {
        if (-not $exists -or (Test-WorkbookWritable -Path $Path)) {
            try {
                if ($exists) { Backup-Workbook -Path $Path }
                New-CheckinWorkbook -Model $model -Stats $stats -Config $config -AsOf $AsOf -Path $Path
                $written = $true
                Write-Log "已写回表格（事项 $($model.Items.Count) 个，日期列 $($model.Dates.Count) 个）"
            }
            catch {
                $writeReason = $_.Exception.Message
                Write-Log "写回表格失败：$writeReason" 'ERROR'
            }
        }
        else {
            $writeReason = '表格正被 Excel 打开，本次只读评估'
            Write-Log $writeReason 'WARN'
        }
    }
    elseif (-not $readOk) {
        $writeReason = '表格读取失败，已跳过写回以保护原文件'
    }

    $state = Get-StateData
    $state.lastEvaluateAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Save-StateData $state

    if ($readOk) {
        try {
            $cachePayload = Get-ReminderPayload ([pscustomobject]@{ Stats = $stats; Written = $written })
            Save-PayloadCache -Payload $cachePayload -AsOf $AsOf
        }
        catch {
            Write-Log "写入提醒快照失败：$($_.Exception.Message)" 'WARN'
        }
    }

    return [pscustomobject]@{
        Ok         = $readOk
        Error      = $readError
        Written    = $written
        WriteReason = $writeReason
        Stats      = $stats
        Model      = $model
        Config     = $config
        Path       = $Path
        AsOf       = $AsOf
    }
}

function Save-PayloadCache {
    param($Payload, [datetime]$AsOf)
    $obj = [pscustomobject]@{
        savedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        asOf    = $AsOf.ToString('yyyy-MM-dd')
        payload = $Payload
    }
    $path = Join-Path $script:RootDir 'cache.json'
    [IO.File]::WriteAllText($path, ($obj | ConvertTo-Json -Depth 6), (New-Object System.Text.UTF8Encoding($false)))
}

function Get-PayloadCache {
    $path = Join-Path $script:RootDir 'cache.json'
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    try {
        return (Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json)
    }
    catch { return $null }
}

function Get-ReminderPayload {
    param($Result)
    if (-not $Result -or -not $Result.Stats) {
        return [pscustomobject]@{
            ok = $false
            title = '每日打卡'
            text = '打卡表读取失败，请手动打开检查。'
            todayDue = 0; todayDone = 0; todayPending = 0; yesterdayPending = 0; expiring = 0
        }
    }
    $s = $Result.Stats
    $yPending = $s.Yesterday.Pending
    $expiring = @($s.ExpiringList).Count
    $lines = @()
    if ($s.Today.Due -eq 0) {
        $lines += '今天没有需要打卡的事项'
    }
    elseif ($s.Today.Pending -le 0) {
        $lines += "今天 $($s.Today.Due) 项已全部打勾 🎉"
    }
    else {
        $lines += "今日待办 $($s.Today.Pending) 项 / 共 $($s.Today.Due) 项"
    }
    if ($yPending -gt 0) { $lines += "昨日待补签 $yPending 项" }
    if ($expiring -gt 0) { $lines += "$expiring 项即将超期" }
    $lines += "近7天：稳定 $($s.Counts['稳定坚持']) · 正常 $($s.Counts['正常']) · 懈怠 $($s.Counts['懈怠']) · 没做 $($s.Counts['基本没做'])"
    if (@($s.NoneList).Count -gt 0) { $lines += '基本没做：' + (Get-ListText $s.NoneList 2) }
    elseif (@($s.SlackList).Count -gt 0) { $lines += '懈怠：' + (Get-ListText $s.SlackList 2) }
    if (-not $Result.Written) { $lines += '本次未写回（表格可能正打开）' }
    return [pscustomobject]@{
        ok              = $true
        title           = '每日打卡'
        text            = ($lines -join "`n")
        todayDue        = $s.Today.Due
        todayDone       = $s.Today.Done
        todayPending    = $s.Today.Pending
        yesterdayPending = $yPending
        expiring        = $expiring
        written         = $Result.Written
        counts          = $s.Counts
    }
}
