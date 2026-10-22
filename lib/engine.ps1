# 打卡数据模型、读取、统计与写回。

. (Join-Path $PSScriptRoot 'xlsx.ps1')

$script:StDone = '✅完成'
$script:StMiss = '❌未做'
$script:StNA = '➖不适用'
$script:StPatch = '🩹补签'
$script:StatusActive = '进行中'
$script:StatusDone = '已完成'
$script:StatusDeleted = '已删除'
$script:RootDir = Split-Path $PSScriptRoot -Parent
$script:DataFile = Join-Path $script:RootDir 'checkin.xlsx'
$script:ConfigFile = Join-Path $script:RootDir 'config.json'
$script:StateFile = Join-Path $script:RootDir 'state.json'
$script:LogFile = Join-Path $script:RootDir 'log.txt'
$script:BackupDir = Join-Path $script:RootDir 'backup'

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Level, $Message
    try {
        [IO.File]::AppendAllText($script:LogFile, $line + [Environment]::NewLine, (New-Object System.Text.UTF8Encoding($false)))
    }
    catch { }
}

function Get-AppConfig {
    $defaults = [ordered]@{
        graceDays      = 7
        maxItems       = 12
        minItemRows    = 20
        maxDateColumns = 730
        weightMap      = @{ '高' = 3; '中' = 2; '低' = 1 }
        reminderTime   = '10:00'
        keepSeconds    = 45
        notifyMode     = 'window'
        playSound      = $true
        voiceFile      = 'voice.wav'
        voiceVolume    = 70
    }
    if (Test-Path -LiteralPath $script:ConfigFile) {
        try {
            $raw = Get-Content -LiteralPath $script:ConfigFile -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($k in @($defaults.Keys)) {
                if ($null -ne $raw.$k) { $defaults[$k] = $raw.$k }
            }
        }
        catch {
            Write-Log "配置文件读取失败，使用默认值: $($_.Exception.Message)" 'WARN'
        }
    }
    return [pscustomobject]$defaults
}

function Get-StateData {
    $state = [ordered]@{ lastReminderDate = ''; lastReminderAt = ''; lastEvaluateAt = '' }
    if (Test-Path -LiteralPath $script:StateFile) {
        try {
            $raw = Get-Content -LiteralPath $script:StateFile -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($p in $raw.PSObject.Properties) { $state[$p.Name] = $p.Value }
        }
        catch { }
    }
    return [pscustomobject]$state
}

function Save-StateData {
    param($State)
    $json = $State | ConvertTo-Json -Depth 5
    [IO.File]::WriteAllText($script:StateFile, $json, (New-Object System.Text.UTF8Encoding($false)))
}

function Get-WeightValue {
    param($Weight, $Config)
    $w = [string]$Weight
    $map = $Config.weightMap
    if ($map -is [System.Collections.IDictionary]) {
        if ($map.Contains($w)) { return [int]$map[$w] }
    }
    elseif ($map.PSObject.Properties.Name -contains $w) {
        return [int]$map.$w
    }
    return 1
}

function Get-NormalizedStatus {
    param($Value)
    $s = ([string]$Value).Trim()
    if ($s -eq $script:StatusDone) { return $script:StatusDone }
    if ($s -eq $script:StatusDeleted) { return $script:StatusDeleted }
    return $script:StatusActive
}

function Get-MarkKind {
    param($Value)
    if ($null -eq $Value) { return '' }
    $s = ([string]$Value).Trim()
    if ($s -eq '') { return '' }
    if ($s -match '^✅' -or $s -match '^✔' -or $s -match '^✓' -or $s -match '^√') { return 'done' }
    if ($s -in @('完成', '做了', '已做', '已完成')) { return 'done' }
    if ($s -in @('❌', '✘', '×', 'x', 'X', '❎')) { return 'miss' }
    if ($s -match '^❌') { return 'miss' }
    if ($s -in @('未做', '没做', '不做')) { return 'miss' }
    if ($s -in @('➖', '－', '-', '—', '－不适用')) { return 'na' }
    if ($s -match '^➖') { return 'na' }
    if ($s -in @('不适用', '跳过')) { return 'na' }
    if ($s -match '^🩹' -or $s -eq '补签') { return 'patch' }
    return 'unknown'
}

function Get-MarkText {
    param([string]$Kind)
    switch ($Kind) {
        'done' { return $script:StDone }
        'miss' { return $script:StMiss }
        'na' { return $script:StNA }
        'patch' { return $script:StPatch }
        default { return '' }
    }
}

function ConvertFrom-XValueDate {
    param($Value)
    if ($null -eq $Value) { return $null }
    $s = ([string]$Value).Trim()
    if ($s -eq '') { return $null }
    $num = 0.0
    if ([double]::TryParse($s, [ref]$num)) {
        if ($num -gt 20000 -and $num -lt 90000) {
            return (ConvertFrom-XSerial $num).Date
        }
        return $null
    }
    $formats = [string[]]@('yyyy-MM-dd', 'yyyy/M/d', 'yyyy-M-d', 'M/d/yyyy', 'MM/dd/yyyy', 'yyyy.MM.dd')
    $d = [datetime]::MinValue
    if ([datetime]::TryParseExact($s, $formats, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$d)) {
        return $d.Date
    }
    if ([datetime]::TryParse($s, [ref]$d)) { return $d.Date }
    return $null
}

function New-InitialModel {
    param([datetime]$Today)
    $items = @(
        @{ Name = '早起'; Type = '习惯'; Weight = '高'; Status = $script:StatusActive; StartDate = $Today; FinishDate = $null; DeleteDate = $null; Note = '示例，可直接改名' },
        @{ Name = '运动 30 分钟'; Type = '习惯'; Weight = '中'; Status = $script:StatusActive; StartDate = $Today; FinishDate = $null; DeleteDate = $null; Note = '示例，可直接改名' },
        @{ Name = '阅读 20 分钟'; Type = '习惯'; Weight = '低'; Status = $script:StatusActive; StartDate = $Today; FinishDate = $null; DeleteDate = $null; Note = '示例，可直接改名' }
    )
    return @{
        Items  = $items
        Dates  = @($Today)
        Marks  = @{}
        Source = 'new'
    }
}

function Read-Model {
    param([string]$Path = $script:DataFile, [datetime]$AsOf = (Get-Date).Date)
    if (-not (Test-Path -LiteralPath $Path)) {
        return (New-InitialModel -Today $AsOf)
    }
    $x = Read-XlsxFile -Path $Path
    if (-not $x.Sheets.ContainsKey('打卡')) {
        throw '表格缺少“打卡”页签，无法读取。'
    }
    $sheet = $x.Sheets['打卡']

    $dateCols = @()
    for ($c = 9; $c -le $sheet.MaxCol; $c++) {
        $ref = (Get-XColName $c) + '1'
        $cell = $sheet.Cells[$ref]
        if (-not $cell) {
            $ref = (Get-XColName $c) + '2'
            $cell = $sheet.Cells[$ref]
        }
        if (-not $cell) { continue }
        $d = ConvertFrom-XValueDate $cell.Value
        if ($d) { $dateCols += @{ Col = $c; Date = $d } }
    }
    if ($dateCols.Count -eq 0) {
        $dateCols += @{ Col = 9; Date = $AsOf }
    }
    $dateCols = @($dateCols | Sort-Object { $_.Date })
    $dates = @($dateCols | ForEach-Object { $_.Date })

    $items = @()
    $marks = @{}
    for ($r = 3; $r -le $sheet.MaxRow; $r++) {
        $nameCell = $sheet.Cells["A$r"]
        $name = if ($nameCell) { ([string]$nameCell.Value).Trim() } else { '' }
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $idx = $items.Count
        $type = ([string]($sheet.Cells["B$r"].Value)).Trim()
        if ($type -ne '习惯' -and $type -ne '任务') { $type = '习惯' }
        $weight = ([string]($sheet.Cells["C$r"].Value)).Trim()
        if ($weight -notin @('高', '中', '低')) { $weight = '中' }
        $status = Get-NormalizedStatus ($sheet.Cells["D$r"].Value)
        $start = ConvertFrom-XValueDate ($sheet.Cells["E$r"].Value)
        $finish = ConvertFrom-XValueDate ($sheet.Cells["F$r"].Value)
        $deleted = ConvertFrom-XValueDate ($sheet.Cells["G$r"].Value)
        $note = if ($sheet.Cells["H$r"]) { ([string]$sheet.Cells["H$r"].Value).Trim() } else { '' }
        $items += @{
            Name       = $name
            Type       = $type
            Weight     = $weight
            Status     = $status
            StartDate  = $start
            FinishDate = $finish
            DeleteDate = $deleted
            Note       = $note
        }
        foreach ($dc in $dateCols) {
            $ref = (Get-XColName $dc.Col) + $r
            $cell = $sheet.Cells[$ref]
            if ($cell) {
                $kind = Get-MarkKind $cell.Value
                if ($kind -ne '' -and $kind -ne 'unknown') {
                    $marks["$idx|$($dc.Date.ToString('yyyy-MM-dd'))"] = $kind
                }
            }
        }
    }

    return @{ Items = $items; Dates = $dates; Marks = $marks; Source = 'file' }
}

function Add-MissingDates {
    param($Model, [datetime]$Today, $Config)
    $dates = New-Object System.Collections.Generic.List[datetime]
    foreach ($d in $Model.Dates) { if (-not $dates.Contains($d.Date)) { $dates.Add($d.Date) } }
    if ($dates.Count -eq 0) { $dates.Add($Today.Date) }
    $last = $dates[$dates.Count - 1]
    if ($last -lt $Today.Date) {
        $guard = 0
        while ($last -lt $Today.Date -and $dates.Count -lt [int]$Config.maxDateColumns -and $guard -lt 4000) {
            $last = $last.AddDays(1)
            $dates.Add($last)
            $guard++
        }
    }
    $Model.Dates = @($dates | Sort-Object)
    return $Model
}

function Get-ItemMarkMap {
    param($Model, [int]$Index)
    $map = @{}
    foreach ($d in $Model.Dates) {
        $key = "$Index|$($d.ToString('yyyy-MM-dd'))"
        $kind = ''
        if ($Model.Marks.ContainsKey($key)) { $kind = $Model.Marks[$key] }
        $map[$d.ToString('yyyy-MM-dd')] = $kind
    }
    return $map
}

function Get-Streaks {
    param($Item, $MarkMap, [datetime]$AsOf, [int]$MaxBack = 365)
    $todayKey = $AsOf.ToString('yyyy-MM-dd')
    $start = $Item.StartDate.Date
    $finish = if ($Item.FinishDate) { $Item.FinishDate.Date } else { $null }

    $doneStart = $AsOf.Date
    $todayKind = if ($MarkMap.ContainsKey($todayKey)) { $MarkMap[$todayKey] } else { '' }
    if ($todayKind -ne 'done' -and $todayKind -ne 'patch') { $doneStart = $AsOf.Date.AddDays(-1) }
    if ($finish -and $doneStart -gt $finish) { $doneStart = $finish }
    $streakDone = 0
    $d = $doneStart
    $n = 0
    while ($d -ge $start -and $n -lt $MaxBack) {
        if ($finish -and $d -gt $finish) { $d = $d.AddDays(-1); $n++; continue }
        $k = if ($MarkMap.ContainsKey($d.ToString('yyyy-MM-dd'))) { $MarkMap[$d.ToString('yyyy-MM-dd')] } else { '__missing__' }
        if ($k -eq '__missing__') { break }
        if ($k -eq 'done' -or $k -eq 'patch') { $streakDone++; $d = $d.AddDays(-1); $n++; continue }
        if ($k -eq 'na') { $d = $d.AddDays(-1); $n++; continue }
        break
    }

    $breakStart = $AsOf.Date.AddDays(-1)
    if ($todayKind -eq 'miss') { $breakStart = $AsOf.Date }
    $streakBreak = 0
    $d = $breakStart
    $n = 0
    while ($d -ge $start -and $n -lt $MaxBack) {
        if ($finish -and $d -gt $finish) { break }
        $k = if ($MarkMap.ContainsKey($d.ToString('yyyy-MM-dd'))) { $MarkMap[$d.ToString('yyyy-MM-dd')] } else { '__missing__' }
        if ($k -eq '__missing__') { break }
        if ($k -eq 'done' -or $k -eq 'patch') { break }
        if ($k -eq 'na') { $d = $d.AddDays(-1); $n++; continue }
        $streakBreak++
        $d = $d.AddDays(-1)
        $n++
    }
    return @{ Done = $streakDone; Break = $streakBreak }
}

function Get-PeriodStats {
    param($Item, $MarkMap, [datetime]$Start, [datetime]$End, $Config, [datetime]$AsOf = (Get-Date).Date)
    $due = 0; $raw = 0; $patch = 0; $na = 0; $pending = 0; $locked = 0
    $grace = [int]$Config.graceDays
    $finish = if ($Item.FinishDate) { $Item.FinishDate.Date } else { $null }
    $d = $Start.Date
    while ($d -le $End.Date) {
        if ($d -lt $Item.StartDate.Date) { $d = $d.AddDays(1); continue }
        if ($finish -and $d -gt $finish) { break }
        $key = $d.ToString('yyyy-MM-dd')
        $k = if ($MarkMap.ContainsKey($key)) { $MarkMap[$key] } else { '' }
        if ($k -eq 'na') {
            $na++
        }
        else {
            $due++
            switch ($k) {
                'done' { $raw++ }
                'patch' { $patch++ }
                'miss' { }
                default {
                    $pending++
                    $elapsed = ($AsOf.Date - $d).Days
                    if ($elapsed -gt $grace) { $locked++ }
                }
            }
        }
        $d = $d.AddDays(1)
    }
    return @{ Due = $due; Raw = $raw; Patch = $patch; NA = $na; Pending = $pending; Locked = $locked }
}

function Get-Judge {
    param($Due, $Raw, $Patch, $Break)
    if ($Due -le 0) { return '暂无数据' }
    if ($Due -lt 4) {
        if (($Raw + $Patch) -ge 1) { return '正常' }
        return '待观察'
    }
    if ($Break -ge 3 -or $Raw -le 1) { return '基本没做' }
    if ($Break -ge 2) { return '懈怠' }
    if ($Raw -ge 6) { return '稳定坚持' }
    if ($Raw -ge 4) { return '正常' }
    return '懈怠'
}

function Get-LockedDays {
    param($Item, $MarkMap, [datetime]$AsOf, [int]$Grace, [int]$Lookback = 30)
    $start = $AsOf.Date.AddDays(-$Lookback)
    if ($Item.StartDate.Date -gt $start) { $start = $Item.StartDate.Date }
    $end = $AsOf.Date.AddDays(-1)
    $count = 0
    $d = $start
    while ($d -le $end) {
        $key = $d.ToString('yyyy-MM-dd')
        $k = if ($MarkMap.ContainsKey($key)) { $MarkMap[$key] } else { '' }
        if ($k -eq '') {
            $elapsed = ($AsOf.Date - $d).Days
            if ($elapsed -gt $Grace) { $count++ }
        }
        $d = $d.AddDays(1)
    }
    return $count
}

function Get-Advice {
    param([string]$Judge)
    switch ($Judge) {
        '稳定坚持' { return '保持当前节奏' }
        '正常' { return '继续保持，偶尔中断不用焦虑' }
        '懈怠' { return '最近有中断，今天优先做这一项' }
        '基本没做' { return '连续多天未做，考虑降低难度或标记完成/删除' }
        '待观察' { return '刚开始记录，先坚持几天' }
        default { return '还没有足够记录' }
    }
}

function Get-AllStats {
    param($Model, [datetime]$AsOf, $Config)
    $weekEnd = $AsOf.Date.AddDays(-1)
    $weekStart = $weekEnd.AddDays(-6)
    $grace = [int]$Config.graceDays

    $itemStats = @()
    $weights = @{ '高' = 3; '中' = 2; '低' = 1 }
    if ($Config.weightMap) {
        if ($Config.weightMap -is [System.Collections.IDictionary]) {
            foreach ($k in $Config.weightMap.Keys) { $weights[[string]$k] = [int]$Config.weightMap[$k] }
        }
        else {
            foreach ($p in $Config.weightMap.PSObject.Properties) { $weights[$p.Name] = [int]$p.Value }
        }
    }

    for ($i = 0; $i -lt $Model.Items.Count; $i++) {
        $item = $Model.Items[$i]
        $map = Get-ItemMarkMap -Model $Model -Index $i
        if ($item.Status -eq $script:StatusDone) {
            $start = $item.StartDate.Date
            $end = if ($item.FinishDate) { $item.FinishDate.Date } else { $AsOf.Date }
            if ($end -lt $start) { $end = $start }
            $full = Get-PeriodStats -Item $item -MarkMap $map -Start $start -End $end -Config $Config -AsOf $AsOf
            $days = ($end - $start).Days + 1
            $itemStats += [pscustomobject]@{
                Index = $i; Item = $item; Kind = 'done'
                Due = $full.Due; Raw = $full.Raw; Patch = $full.Patch; NA = $full.NA
                RawRate = if ($full.Due -gt 0) { $full.Raw / $full.Due } else { $null }
                PatchRate = if ($full.Due -gt 0) { ($full.Raw + $full.Patch) / $full.Due } else { $null }
                StreakDone = 0; StreakBreak = 0; Pending = 0; Locked = 0
                StatusText = '已完成'; Advice = ''; DurationDays = $days
            }
            continue
        }
        if ($item.Status -eq $script:StatusDeleted) {
            $start = $item.StartDate.Date
            $end = if ($item.DeleteDate) { $item.DeleteDate.Date } else { $AsOf.Date }
            if ($end -lt $start) { $end = $start }
            $full = Get-PeriodStats -Item $item -MarkMap $map -Start $start -End $end -Config $Config -AsOf $AsOf
            $days = ($end - $start).Days + 1
            $itemStats += [pscustomobject]@{
                Index = $i; Item = $item; Kind = 'deleted'
                Due = $full.Due; Raw = $full.Raw; Patch = $full.Patch; NA = $full.NA
                RawRate = if ($full.Due -gt 0) { $full.Raw / $full.Due } else { $null }
                PatchRate = if ($full.Due -gt 0) { ($full.Raw + $full.Patch) / $full.Due } else { $null }
                StreakDone = 0; StreakBreak = 0; Pending = 0; Locked = 0
                StatusText = '已删除'; Advice = ''; DurationDays = $days
            }
            continue
        }

        $period = Get-PeriodStats -Item $item -MarkMap $map -Start $weekStart -End $weekEnd -Config $Config -AsOf $AsOf
        $streaks = Get-Streaks -Item $item -MarkMap $map -AsOf $AsOf
        $lockedDays = Get-LockedDays -Item $item -MarkMap $map -AsOf $AsOf -Grace $grace
        $judge = Get-Judge -Due $period.Due -Raw $period.Raw -Patch $period.Patch -Break $streaks.Break
        $itemStats += [pscustomobject]@{
            Index = $i; Item = $item; Kind = 'active'
            Due = $period.Due; Raw = $period.Raw; Patch = $period.Patch; NA = $period.NA
            RawRate = if ($period.Due -gt 0) { $period.Raw / $period.Due } else { $null }
            PatchRate = if ($period.Due -gt 0) { ($period.Raw + $period.Patch) / $period.Due } else { $null }
            StreakDone = $streaks.Done; StreakBreak = $streaks.Break
            Pending = $period.Pending; Locked = $lockedDays
            StatusText = $judge; Advice = (Get-Advice $judge); DurationDays = 0
        }
    }

    $active = @($itemStats | Where-Object { $_.Kind -eq 'active' })

    $todayDue = 0; $todayDone = 0; $todayPatch = 0; $todayNA = 0
    foreach ($s in $active) {
        $key = $AsOf.ToString('yyyy-MM-dd')
        if ($s.Item.StartDate -gt $AsOf) { continue }
        $k = if ($Model.Marks.ContainsKey("$($s.Index)|$key")) { $Model.Marks["$($s.Index)|$key"] } else { '' }
        if ($k -eq 'na') { $todayNA++; continue }
        $todayDue++
        if ($k -eq 'done') { $todayDone++ }
        elseif ($k -eq 'patch') { $todayPatch++ }
    }

    $ystart = $AsOf.AddDays(-1)
    $yDue = 0; $yRaw = 0; $yPatch = 0; $yNA = 0; $yW = 0; $yD = 0; $yP = 0; $yPending = 0
    foreach ($s in $active) {
        $key = $ystart.ToString('yyyy-MM-dd')
        if ($s.Item.StartDate -gt $ystart) { continue }
        $k = if ($Model.Marks.ContainsKey("$($s.Index)|$key")) { $Model.Marks["$($s.Index)|$key"] } else { '' }
        if ($k -eq 'na') { $yNA++; continue }
        $yDue++
        $w = $weights[$s.Item.Weight]
        if ($null -eq $w) { $w = 1 }
        $yW += $w
        if ($k -eq 'done') { $yRaw++; $yD += $w }
        elseif ($k -eq 'patch') { $yPatch++; $yP += $w }
        elseif ($k -eq '') { $yPending++ }
    }

    $wDue = 0; $wDoneW = 0; $wPatchW = 0; $wRaw = 0; $wPatch = 0; $wNA = 0; $wDueTotal = 0
    foreach ($s in $active) {
        $w = $weights[$s.Item.Weight]
        if ($null -eq $w) { $w = 1 }
        $wDue += $s.Due * $w
        $wDoneW += $s.Raw * $w
        $wPatchW += ($s.Raw + $s.Patch) * $w
        $wRaw += $s.Raw
        $wPatch += $s.Patch
        $wNA += $s.NA
        $wDueTotal += $s.Due
    }

    $trend = @()
    for ($n = 6; $n -ge 0; $n--) {
        $d = $weekStart.AddDays(6 - $n)
        $due = 0; $raw = 0; $patch = 0; $na = 0; $wd = 0; $ww = 0; $wp = 0
        foreach ($s in $active) {
            $key = $d.ToString('yyyy-MM-dd')
            if ($s.Item.StartDate -gt $d) { continue }
            $k = if ($Model.Marks.ContainsKey("$($s.Index)|$key")) { $Model.Marks["$($s.Index)|$key"] } else { '' }
            if ($k -eq 'na') { $na++; continue }
            $due++
            $w = $weights[$s.Item.Weight]
            if ($null -eq $w) { $w = 1 }
            $ww += $w
            if ($k -eq 'done') { $raw++; $wd += $w }
            elseif ($k -eq 'patch') { $patch++; $wp += $w }
        }
        $trend += [pscustomobject]@{
            Date = $d; Due = $due; Raw = $raw; Patch = $patch; NA = $na
            RawRate = if ($due -gt 0) { $raw / $due } else { $null }
            PatchRate = if ($due -gt 0) { ($raw + $patch) / $due } else { $null }
            WeightedRate = if ($ww -gt 0) { $wd / $ww } else { $null }
        }
    }

    $pendingList = @($active | Where-Object { $_.Pending -gt 0 } | ForEach-Object { "$($_.Item.Name)（$($_.Pending)天）" })
    $lockedList = @($active | Where-Object { $_.Locked -gt 0 } | ForEach-Object { "$($_.Item.Name)（$($_.Locked)天）" })

    $expiring = @()
    foreach ($s in $active) {
        $left = 0
        foreach ($d in $Model.Dates) {
            if ($d -ge $AsOf.Date) { continue }
            if ($d -lt $s.Item.StartDate.Date) { continue }
            $key = $d.ToString('yyyy-MM-dd')
            $k = if ($Model.Marks.ContainsKey("$($s.Index)|$key")) { $Model.Marks["$($s.Index)|$key"] } else { '' }
            if ($k -ne '') { continue }
            $daysLeft = $grace - ($AsOf.Date - $d).Days
            if ($daysLeft -eq 1) { $left++ }
        }
        if ($left -gt 0) { $expiring += "$($s.Item.Name)（$left 天）" }
    }

    $counts = @{ '稳定坚持' = 0; '正常' = 0; '懈怠' = 0; '基本没做' = 0; '待观察' = 0; '暂无数据' = 0 }
    foreach ($s in $active) {
        if ($counts.ContainsKey($s.StatusText)) { $counts[$s.StatusText]++ }
    }

    $slackList = @($active | Where-Object { $_.StatusText -eq '懈怠' } | ForEach-Object { $_.Item.Name })
    $noneList = @($active | Where-Object { $_.StatusText -eq '基本没做' } | ForEach-Object { $_.Item.Name })

    return [pscustomobject]@{
        AsOf         = $AsOf.Date
        WeekStart    = $weekStart
        WeekEnd      = $weekEnd
        Items        = $itemStats
        Active       = $active
        Today        = [pscustomobject]@{ Due = $todayDue; Done = $todayDone; Patch = $todayPatch; NA = $todayNA; Pending = ($todayDue - $todayDone - $todayPatch) }
        Yesterday    = [pscustomobject]@{ Date = $ystart; Due = $yDue; Raw = $yRaw; Patch = $yPatch; NA = $yNA; Pending = $yPending; RawRate = if ($yDue -gt 0) { $yRaw / $yDue } else { $null }; PatchRate = if ($yDue -gt 0) { ($yRaw + $yPatch) / $yDue } else { $null }; WeightedRate = if ($yW -gt 0) { $yD / $yW } else { $null } }
        Week         = [pscustomobject]@{ Due = $wDueTotal; Raw = $wRaw; Patch = $wPatch; NA = $wNA; RawRate = if ($wDueTotal -gt 0) { $wRaw / $wDueTotal } else { $null }; PatchRate = if ($wDueTotal -gt 0) { ($wRaw + $wPatch) / $wDueTotal } else { $null }; WeightedRate = if ($wDue -gt 0) { $wDoneW / $wDue } else { $null } }
        Trend        = $trend
        PendingList  = $pendingList
        LockedList   = $lockedList
        ExpiringList = $expiring
        SlackList    = $slackList
        NoneList     = $noneList
        Counts       = $counts
        GraceDays    = $grace
    }
}

function Backup-Workbook {
    param([string]$Path = $script:DataFile)
    if (-not (Test-Path -LiteralPath $Path)) { return }
    if (-not (Test-Path -LiteralPath $script:BackupDir)) {
        New-Item -ItemType Directory -Force -Path $script:BackupDir | Out-Null
    }
    $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
    $dest = Join-Path $script:BackupDir ("checkin-$stamp.xlsx")
    Copy-Item -LiteralPath $Path -Destination $dest -Force
    $files = Get-ChildItem -LiteralPath $script:BackupDir -Filter 'checkin-*.xlsx' | Sort-Object LastWriteTime -Descending
    if ($files.Count -gt 14) {
        $files | Select-Object -Skip 14 | ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force }
    }
}

function Test-WorkbookWritable {
    param([string]$Path = $script:DataFile)
    if (-not (Test-Path -LiteralPath $Path)) { return $true }
    try {
        $fs = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $fs.Close()
        return $true
    }
    catch { return $false }
}
