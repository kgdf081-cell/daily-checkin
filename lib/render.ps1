# 把数据模型渲染为 xlsx 的六个页签。

. (Join-Path $PSScriptRoot 'engine.ps1')

function Get-ListText {
    param($List, [int]$Max = 3)
    $arr = @($List | Where-Object { $_ })
    if ($arr.Count -eq 0) { return '无' }
    if ($arr.Count -le $Max) { return ($arr -join '、') }
    return (($arr | Select-Object -First $Max) -join '、') + " 等 $($arr.Count) 项"
}

function Add-TitleRows {
    param($Sheet, [string]$Title, [string]$Subtitle, [int]$LastCol)
    Add-XCell $Sheet 'A1' $Title 2
    $Sheet.Merges += "A1:$(Get-XColName $LastCol)1"
    Add-XRow $Sheet 1 -Height 30
    Add-XCell $Sheet 'A2' $Subtitle 13
    $Sheet.Merges += "A2:$(Get-XColName $LastCol)2"
    Add-XRow $Sheet 2 -Height 22
}

function New-CheckinSheet {
    param($Model, $Stats, $Config, [datetime]$AsOf)
    $sh = New-XSheet '打卡'
    $sh.TabColor = 'FF1F4E79'
    $sh.Freeze = 'I3'
    $sh.TabSelected = $true

    $dates = @($Model.Dates | Sort-Object)
    if ($dates.Count -eq 0) { $dates = @($AsOf.Date) }
    $dateCount = $dates.Count
    $lastCol = 8 + $dateCount
    $lastColName = Get-XColName $lastCol
    $rowCount = [Math]::Max([int]$Config.minItemRows, $Model.Items.Count + 4)
    $lastRow = 2 + $rowCount

    $cols = @(
        @{ Min = 1; Max = 1; Width = 24 },
        @{ Min = 2; Max = 2; Width = 9 },
        @{ Min = 3; Max = 3; Width = 7 },
        @{ Min = 4; Max = 4; Width = 9 },
        @{ Min = 5; Max = 5; Width = 11 },
        @{ Min = 6; Max = 6; Width = 11 },
        @{ Min = 7; Max = 7; Width = 11 },
        @{ Min = 8; Max = 8; Width = 24 },
        @{ Min = 9; Max = $lastCol; Width = 9.5 }
    )
    if ($dateCount -eq 0) { $cols = $cols[0..7] }
    $sh.Cols = $cols

    Add-XRow $sh 1 -Hidden $true
    Add-XCell $sh 'A1' '内部日期行（请勿删除）' 13
    Add-XCell $sh 'E1' 'start' 13
    Add-XCell $sh 'F1' 'finish' 13
    Add-XCell $sh 'G1' 'delete' 13
    for ($i = 0; $i -lt $dateCount; $i++) {
        $cn = Get-XColName (9 + $i)
        Add-XCell $sh ($cn + '1') (ConvertTo-XSerial $dates[$i]) 3 'n'
    }

    Add-XRow $sh 2 -Height 34
    $headers = @('事项', '类型', '权重', '状态', '开始日期', '完成日期', '删除日期', '备注')
    for ($i = 0; $i -lt $headers.Count; $i++) {
        Add-XCell $sh ((Get-XColName ($i + 1)) + '2') $headers[$i] 1
    }
    for ($i = 0; $i -lt $dateCount; $i++) {
        $cn = Get-XColName (9 + $i)
        $style = 1
        if ($dates[$i].Date -eq $AsOf.Date) { $style = 23 }
        Add-XCell $sh ($cn + '2') $dates[$i].ToString('MM-dd') $style
    }
    $todayCol = $lastCol
    for ($i = 0; $i -lt $dateCount; $i++) {
        if ($dates[$i].Date -eq $AsOf.Date) { $todayCol = 9 + $i; break }
    }
    $sh.ViewCell = (Get-XColName $todayCol) + '3'

    for ($i = 0; $i -lt $rowCount; $i++) {
        $r = 3 + $i
        Add-XRow $sh $r -Height 22
        $item = $null
        if ($i -lt $Model.Items.Count) { $item = $Model.Items[$i] }
        if ($item) {
            Add-XCell $sh "A$r" $item.Name 5
            Add-XCell $sh "B$r" $item.Type 4
            Add-XCell $sh "C$r" $item.Weight 4
            Add-XCell $sh "D$r" $item.Status 4
            if ($item.StartDate) { Add-XCell $sh "E$r" (ConvertTo-XSerial $item.StartDate) 3 'n' }
            else { Add-XCell $sh "E$r" '' 3 }
            if ($item.FinishDate) { Add-XCell $sh "F$r" (ConvertTo-XSerial $item.FinishDate) 3 'n' }
            else { Add-XCell $sh "F$r" '' 3 }
            if ($item.DeleteDate) { Add-XCell $sh "G$r" (ConvertTo-XSerial $item.DeleteDate) 3 'n' }
            else { Add-XCell $sh "G$r" '' 3 }
            Add-XCell $sh "H$r" $item.Note 9
        }
        else {
            Add-XCell $sh "A$r" '' 5
            Add-XCell $sh "B$r" '' 4
            Add-XCell $sh "C$r" '' 4
            Add-XCell $sh "D$r" '' 4
            Add-XCell $sh "E$r" '' 3
            Add-XCell $sh "F$r" '' 3
            Add-XCell $sh "G$r" '' 3
            Add-XCell $sh "H$r" '' 9
        }
        for ($j = 0; $j -lt $dateCount; $j++) {
            $cn = Get-XColName (9 + $j)
            $key = "$i|$($dates[$j].ToString('yyyy-MM-dd'))"
            $kind = if ($Model.Marks.ContainsKey($key)) { $Model.Marks[$key] } else { '' }
            $text = Get-MarkText $kind
            Add-XCell $sh ($cn + $r) $text 4
        }
    }

    $sh.Validations = @(
        @{ Sqref = "B3:B$lastRow"; Formula = '"习惯,任务"' },
        @{ Sqref = "C3:C$lastRow"; Formula = '"高,中,低"' },
        @{ Sqref = "D3:D$lastRow"; Formula = '"进行中,已完成,已删除"' },
        @{ Sqref = "I3:$lastColName$lastRow"; Formula = '"✅完成,❌未做,➖不适用,🩹补签"' }
    )

    $grace = [int]$Config.graceDays
    $dateRange = "I3:$lastColName$lastRow"
    $baseRange = "A3:H$lastRow"
    $sh.CondFormats = @(
        @{
            Sqref = $dateRange
            Rules = @(
                @{ Type = 'cellIs'; Operator = 'equal'; Dxf = 0; Priority = 1; Stop = $true; Formula = '"✅完成"' },
                @{ Type = 'cellIs'; Operator = 'equal'; Dxf = 1; Priority = 2; Stop = $true; Formula = '"❌未做"' },
                @{ Type = 'cellIs'; Operator = 'equal'; Dxf = 2; Priority = 3; Stop = $true; Formula = '"➖不适用"' },
                @{ Type = 'cellIs'; Operator = 'equal'; Dxf = 3; Priority = 4; Stop = $true; Formula = '"🩹补签"' },
                @{ Type = 'expression'; Dxf = 4; Priority = 5; Stop = $false; Formula = "AND(`$D3=`"进行中`",I`$1<TODAY(),I`$1>=TODAY()-$grace,I3=`"`")" },
                @{ Type = 'expression'; Dxf = 7; Priority = 6; Stop = $false; Formula = "AND(`$D3=`"进行中`",I`$1<TODAY()-$grace,I`$1>=TODAY()-30,I3=`"`")" },
                @{ Type = 'expression'; Dxf = 5; Priority = 7; Stop = $false; Formula = '$D3="已完成"' },
                @{ Type = 'expression'; Dxf = 6; Priority = 8; Stop = $false; Formula = '$D3="已删除"' }
            )
        },
        @{
            Sqref = $baseRange
            Rules = @(
                @{ Type = 'expression'; Dxf = 5; Priority = 9; Stop = $false; Formula = '$D3="已完成"' },
                @{ Type = 'expression'; Dxf = 6; Priority = 10; Stop = $false; Formula = '$D3="已删除"' }
            )
        }
    )
    return $sh
}

function New-EvalSheet {
    param($Model, $Stats, $AsOf)
    $sh = New-XSheet '事项评估'
    $sh.TabColor = 'FF70AD47'
    $sh.Freeze = 'A4'
    $sh.Cols = @(
        @{ Min = 1; Max = 1; Width = 24 },
        @{ Min = 2; Max = 2; Width = 8 },
        @{ Min = 3; Max = 3; Width = 7 },
        @{ Min = 4; Max = 4; Width = 9 },
        @{ Min = 5; Max = 5; Width = 9 },
        @{ Min = 6; Max = 6; Width = 7 },
        @{ Min = 7; Max = 7; Width = 8 },
        @{ Min = 8; Max = 8; Width = 10 },
        @{ Min = 9; Max = 9; Width = 12 },
        @{ Min = 10; Max = 10; Width = 9 },
        @{ Min = 11; Max = 11; Width = 9 },
        @{ Min = 12; Max = 12; Width = 8 },
        @{ Min = 13; Max = 13; Width = 8 },
        @{ Min = 14; Max = 14; Width = 8 },
        @{ Min = 15; Max = 15; Width = 10 },
        @{ Min = 16; Max = 16; Width = 34 }
    )
    $sub = '统计窗口：{0} ~ {1}（近 7 个完整日，不含今天）｜生成：{2}' -f $Stats.WeekStart.ToString('yyyy-MM-dd'), $Stats.WeekEnd.ToString('yyyy-MM-dd'), (Get-Date).ToString('yyyy-MM-dd HH:mm')
    Add-TitleRows -Sheet $sh -Title '事项评估 · 近 7 天' -Subtitle $sub -LastCol 16
    Add-XRow $sh 3 -Height 32
    $headers = @('事项', '类型', '权重', '近7天应做', '原始完成', '补签', '不适用', '原始完成率', '含补签完成率', '连续完成', '连续中断', '待补签', '已补签', '超期锁定', '判定', '建议')
    for ($i = 0; $i -lt $headers.Count; $i++) {
        Add-XCell $sh ((Get-XColName ($i + 1)) + '3') $headers[$i] 1
    }
    $row = 4
    foreach ($s in $Stats.Active) {
        Add-XRow $sh $row -Height 24
        Add-XCell $sh "A$row" $s.Item.Name 5
        Add-XCell $sh "B$row" $s.Item.Type 4
        Add-XCell $sh "C$row" $s.Item.Weight 4
        Add-XCell $sh "D$row" $s.Due 7
        Add-XCell $sh "E$row" $s.Raw 7
        Add-XCell $sh "F$row" $s.Patch 7
        Add-XCell $sh "G$row" $s.NA 7
        if ($null -ne $s.RawRate) { Add-XCell $sh "H$row" ([double]$s.RawRate) 6 } else { Add-XCell $sh "H$row" '' 6 }
        if ($null -ne $s.PatchRate) { Add-XCell $sh "I$row" ([double]$s.PatchRate) 6 } else { Add-XCell $sh "I$row" '' 6 }
        Add-XCell $sh "J$row" $s.StreakDone 7
        Add-XCell $sh "K$row" $s.StreakBreak 7
        Add-XCell $sh "L$row" $s.Pending 7
        Add-XCell $sh "M$row" $s.Patch 7
        Add-XCell $sh "N$row" $s.Locked 7
        $js = 4
        switch ($s.StatusText) {
            '稳定坚持' { $js = 24 }
            '正常' { $js = 4 }
            '懈怠' { $js = 26 }
            '基本没做' { $js = 25 }
            '待观察' { $js = 26 }
            default { $js = 27 }
        }
        Add-XCell $sh "O$row" $s.StatusText $js
        Add-XCell $sh "P$row" $s.Advice 9
        $row++
    }
    if ($Stats.Active.Count -eq 0) {
        Add-XRow $sh 4 -Height 30
        Add-XCell $sh 'A4' '暂无进行中的事项。请在「打卡」页填写事项名称。' 9
        $sh.Merges += 'A4:P4'
    }
    return $sh
}

function New-SummarySheet {
    param($Model, $Stats, $AsOf, $Config)
    $sh = New-XSheet '每日汇总'
    $sh.TabColor = 'FFED7D31'
    $sh.Freeze = 'A4'
    $sh.Cols = @(
        @{ Min = 1; Max = 1; Width = 24 },
        @{ Min = 2; Max = 2; Width = 22 },
        @{ Min = 3; Max = 7; Width = 14 }
    )
    $sub = '生成时间：{0}｜补签窗口：{1} 天｜权重：高 3 / 中 2 / 低 1' -f (Get-Date).ToString('yyyy-MM-dd HH:mm'), $Stats.GraceDays
    Add-TitleRows -Sheet $sh -Title '每日汇总' -Subtitle $sub -LastCol 7

    $row = 4
    Add-XRow $sh $row -Height 26
    Add-XCell $sh "A$row" '今日（进行中）' 12
    $sh.Merges += "A$row:G$row"
    $row++

    $todayRows = @(
        @{ L = '今天日期'; V = $AsOf.ToString('yyyy-MM-dd'); S = 4 },
        @{ L = '今日应做'; V = $Stats.Today.Due; S = 7 },
        @{ L = '今日已完成'; V = $Stats.Today.Done; S = 24 },
        @{ L = '今日补签'; V = $Stats.Today.Patch; S = 26 },
        @{ L = '今日待办'; V = $Stats.Today.Pending; S = 25 }
    )
    foreach ($tr in $todayRows) {
        Add-XRow $sh $row -Height 22
        Add-XCell $sh "A$row" $tr.L 5
        Add-XCell $sh "B$row" $tr.V $tr.S
        $row++
    }

    $row++
    Add-XRow $sh $row -Height 26
    Add-XCell $sh "A$row" ("昨日结算（{0}）" -f $Stats.Yesterday.Date.ToString('MM-dd')) 12
    $sh.Merges += "A$row:G$row"
    $row++
    $yRows = @(
        @{ L = '应做'; V = $Stats.Yesterday.Due; S = 7 },
        @{ L = '原始完成'; V = $Stats.Yesterday.Raw; S = 7 },
        @{ L = '补签'; V = $Stats.Yesterday.Patch; S = 7 },
        @{ L = '不适用'; V = $Stats.Yesterday.NA; S = 7 },
        @{ L = '原始完成率'; V = $Stats.Yesterday.RawRate; S = 6 },
        @{ L = '含补签完成率'; V = $Stats.Yesterday.PatchRate; S = 6 },
        @{ L = '加权完成率'; V = $Stats.Yesterday.WeightedRate; S = 16 }
    )
    foreach ($yr in $yRows) {
        Add-XRow $sh $row -Height 22
        Add-XCell $sh "A$row" $yr.L 5
        if ($null -ne $yr.V) { Add-XCell $sh "B$row" $yr.V $yr.S } else { Add-XCell $sh "B$row" '—' 4 }
        $row++
    }

    $row++
    Add-XRow $sh $row -Height 26
    Add-XCell $sh "A$row" ("近 7 天（{0} ~ {1}）" -f $Stats.WeekStart.ToString('MM-dd'), $Stats.WeekEnd.ToString('MM-dd')) 12
    $sh.Merges += "A$row:G$row"
    $row++
    $wRows = @(
        @{ L = '应做'; V = $Stats.Week.Due; S = 7 },
        @{ L = '原始完成'; V = $Stats.Week.Raw; S = 7 },
        @{ L = '补签'; V = $Stats.Week.Patch; S = 7 },
        @{ L = '不适用'; V = $Stats.Week.NA; S = 7 },
        @{ L = '原始完成率'; V = $Stats.Week.RawRate; S = 6 },
        @{ L = '含补签完成率'; V = $Stats.Week.PatchRate; S = 6 },
        @{ L = '加权完成率'; V = $Stats.Week.WeightedRate; S = 16 }
    )
    foreach ($wr in $wRows) {
        Add-XRow $sh $row -Height 22
        Add-XCell $sh "A$row" $wr.L 5
        if ($null -ne $wr.V) { Add-XCell $sh "B$row" $wr.V $wr.S } else { Add-XCell $sh "B$row" '—' 4 }
        $row++
    }

    $row++
    Add-XRow $sh $row -Height 26
    Add-XCell $sh "A$row" '状态与预警' 12
    $sh.Merges += "A$row:G$row"
    $row++
    $alertRows = @(
        @{ L = '稳定坚持'; V = ('{0} 项' -f $Stats.Counts['稳定坚持']); S = 24 },
        @{ L = '正常'; V = ('{0} 项' -f $Stats.Counts['正常']); S = 4 },
        @{ L = '懈怠'; V = ('{0} 项' -f $Stats.Counts['懈怠']); S = 26 },
        @{ L = '基本没做'; V = ('{0} 项' -f $Stats.Counts['基本没做']); S = 25 },
        @{ L = '待观察 / 无数据'; V = ('{0} / {1} 项' -f $Stats.Counts['待观察'], $Stats.Counts['暂无数据']); S = 27 },
        @{ L = '待补签事项'; V = (Get-ListText $Stats.PendingList 5); S = 9 },
        @{ L = '即将超期'; V = (Get-ListText $Stats.ExpiringList 5); S = 9 },
        @{ L = '超期锁定'; V = (Get-ListText $Stats.LockedList 5); S = 9 },
        @{ L = '懈怠预警'; V = (Get-ListText $Stats.SlackList 5); S = 9 },
        @{ L = '基本没做'; V = (Get-ListText $Stats.NoneList 5); S = 9 }
    )
    foreach ($ar in $alertRows) {
        Add-XRow $sh $row -Height 24
        Add-XCell $sh "A$row" $ar.L 5
        Add-XCell $sh "B$row" $ar.V $ar.S
        if ($ar.S -eq 9) {
            $sh.Merges += "B$row:G$row"
        }
        $row++
    }

    $row++
    Add-XRow $sh $row -Height 26
    Add-XCell $sh "A$row" '近 7 天趋势' 12
    $sh.Merges += "A$row:G$row"
    $row++
    Add-XRow $sh $row -Height 28
    $th = @('日期', '应做', '原始完成', '补签', '原始率', '含补签率', '加权率')
    for ($i = 0; $i -lt $th.Count; $i++) {
        Add-XCell $sh ((Get-XColName ($i + 1)) + "$row") $th[$i] 1
    }
    $row++
    foreach ($t in $Stats.Trend) {
        Add-XRow $sh $row -Height 22
        Add-XCell $sh "A$row" $t.Date.ToString('MM-dd') 4
        Add-XCell $sh "B$row" $t.Due 7
        Add-XCell $sh "C$row" $t.Raw 7
        Add-XCell $sh "D$row" $t.Patch 7
        if ($null -ne $t.RawRate) { Add-XCell $sh "E$row" ([double]$t.RawRate) 6 } else { Add-XCell $sh "E$row" '' 6 }
        if ($null -ne $t.PatchRate) { Add-XCell $sh "F$row" ([double]$t.PatchRate) 6 } else { Add-XCell $sh "F$row" '' 6 }
        if ($null -ne $t.WeightedRate) { Add-XCell $sh "G$row" ([double]$t.WeightedRate) 6 } else { Add-XCell $sh "G$row" '' 6 }
        $row++
    }
    return $sh
}

function New-CompletedSheet {
    param($Model, $Stats, $AsOf)
    $sh = New-XSheet '完成记录'
    $sh.TabColor = 'FF7030A0'
    $sh.Freeze = 'A4'
    $sh.Cols = @(
        @{ Min = 1; Max = 1; Width = 6 },
        @{ Min = 2; Max = 2; Width = 26 },
        @{ Min = 3; Max = 3; Width = 8 },
        @{ Min = 4; Max = 4; Width = 7 },
        @{ Min = 5; Max = 5; Width = 12 },
        @{ Min = 6; Max = 6; Width = 12 },
        @{ Min = 7; Max = 7; Width = 10 },
        @{ Min = 8; Max = 8; Width = 10 },
        @{ Min = 9; Max = 9; Width = 8 },
        @{ Min = 10; Max = 10; Width = 8 },
        @{ Min = 11; Max = 11; Width = 12 },
        @{ Min = 12; Max = 12; Width = 28 }
    )
    Add-TitleRows -Sheet $sh -Title '完成记录' -Subtitle '在「打卡」页把状态改为「已完成」，该事项会自动结束并记录在这里。' -LastCol 12
    Add-XRow $sh 3 -Height 30
    $headers = @('序号', '事项', '类型', '权重', '开始日期', '完成日期', '持续天数', '原始完成', '补签', '不适用', '原始完成率', '备注')
    for ($i = 0; $i -lt $headers.Count; $i++) { Add-XCell $sh ((Get-XColName ($i + 1)) + '3') $headers[$i] 1 }
    $done = @($Stats.Items | Where-Object { $_.Kind -eq 'done' } | Sort-Object { $_.Item.FinishDate } -Descending)
    $row = 4
    $n = 1
    foreach ($s in $done) {
        Add-XRow $sh $row -Height 24
        Add-XCell $sh "A$row" $n 7
        Add-XCell $sh "B$row" $s.Item.Name 5
        Add-XCell $sh "C$row" $s.Item.Type 4
        Add-XCell $sh "D$row" $s.Item.Weight 4
        if ($s.Item.StartDate) { Add-XCell $sh "E$row" (ConvertTo-XSerial $s.Item.StartDate) 3 'n' } else { Add-XCell $sh "E$row" '' 3 }
        if ($s.Item.FinishDate) { Add-XCell $sh "F$row" (ConvertTo-XSerial $s.Item.FinishDate) 3 'n' } else { Add-XCell $sh "F$row" '' 3 }
        Add-XCell $sh "G$row" $s.DurationDays 7
        Add-XCell $sh "H$row" $s.Raw 7
        Add-XCell $sh "I$row" $s.Patch 7
        Add-XCell $sh "J$row" $s.NA 7
        if ($null -ne $s.RawRate) { Add-XCell $sh "K$row" ([double]$s.RawRate) 6 } else { Add-XCell $sh "K$row" '' 6 }
        Add-XCell $sh "L$row" $s.Item.Note 9
        $row++
        $n++
    }
    if ($done.Count -eq 0) {
        Add-XRow $sh 4 -Height 30
        Add-XCell $sh 'A4' '暂无已完成事项。' 9
        $sh.Merges += 'A4:L4'
    }
    return $sh
}

function New-RecycleSheet {
    param($Model, $Stats, $AsOf)
    $sh = New-XSheet '回收站'
    $sh.TabColor = 'FF808080'
    $sh.Freeze = 'A4'
    $sh.Cols = @(
        @{ Min = 1; Max = 1; Width = 6 },
        @{ Min = 2; Max = 2; Width = 26 },
        @{ Min = 3; Max = 3; Width = 8 },
        @{ Min = 4; Max = 4; Width = 7 },
        @{ Min = 5; Max = 5; Width = 12 },
        @{ Min = 6; Max = 6; Width = 12 },
        @{ Min = 7; Max = 7; Width = 10 },
        @{ Min = 8; Max = 8; Width = 10 },
        @{ Min = 9; Max = 9; Width = 10 },
        @{ Min = 10; Max = 10; Width = 32 }
    )
    Add-TitleRows -Sheet $sh -Title '回收站' -Subtitle '恢复方法：在「打卡」页把该事项的状态改回「进行中」，下次刷新即恢复。' -LastCol 10
    Add-XRow $sh 3 -Height 30
    $headers = @('序号', '事项', '类型', '权重', '开始日期', '删除日期', '坚持天数', '完成次数', '补签次数', '备注')
    for ($i = 0; $i -lt $headers.Count; $i++) { Add-XCell $sh ((Get-XColName ($i + 1)) + '3') $headers[$i] 1 }
    $del = @($Stats.Items | Where-Object { $_.Kind -eq 'deleted' } | Sort-Object { $_.Item.DeleteDate } -Descending)
    $row = 4
    $n = 1
    foreach ($s in $del) {
        Add-XRow $sh $row -Height 24
        Add-XCell $sh "A$row" $n 7
        Add-XCell $sh "B$row" $s.Item.Name 5
        Add-XCell $sh "C$row" $s.Item.Type 4
        Add-XCell $sh "D$row" $s.Item.Weight 4
        if ($s.Item.StartDate) { Add-XCell $sh "E$row" (ConvertTo-XSerial $s.Item.StartDate) 3 'n' } else { Add-XCell $sh "E$row" '' 3 }
        if ($s.Item.DeleteDate) { Add-XCell $sh "F$row" (ConvertTo-XSerial $s.Item.DeleteDate) 3 'n' } else { Add-XCell $sh "F$row" '' 3 }
        Add-XCell $sh "G$row" $s.DurationDays 7
        Add-XCell $sh "H$row" $s.Raw 7
        Add-XCell $sh "I$row" $s.Patch 7
        Add-XCell $sh "J$row" $s.Item.Note 9
        $row++
        $n++
    }
    if ($del.Count -eq 0) {
        Add-XRow $sh 4 -Height 30
        Add-XCell $sh 'A4' '回收站是空的。' 9
        $sh.Merges += 'A4:J4'
    }
    return $sh
}

function New-HelpSheet {
    param($AsOf, $Config)
    $sh = New-XSheet '说明'
    $sh.TabColor = 'FFBFBFBF'
    $sh.Cols = @(
        @{ Min = 1; Max = 1; Width = 26 },
        @{ Min = 2; Max = 7; Width = 13 },
        @{ Min = 8; Max = 8; Width = 26 }
    )
    $items = @(
        @{ T = '每日打卡与完成评估 · 使用说明'; S = 2; H = 34 },
        @{ T = '表格是唯一需要手动操作的文件；提醒、补日期列、评估都由脚本自动完成。'; S = 13; H = 22 },
        @{ T = '一、每天怎么用'; S = 12; H = 26 },
        @{ T = '双击「打开今日打卡」（或等 22:00 的提醒）→ 脚本先补上今天的日期列并刷新评估 → 表格自动打开。在今天这一列逐项选择状态，保存即可。'; S = 13; H = 60 },
        @{ T = '二、四种打卡状态'; S = 12; H = 26 },
        @{ T = '✅完成 = 做了；❌未做 = 明确没做；➖不适用 = 今天不该做（不计入应做天数）；🩹补签 = 事后补记。空白 = 还没打勾，过去日期会显示黄色虚线框提示待补签。'; S = 13; H = 75 },
        @{ T = '三、增加 / 减少事项'; S = 12; H = 26 },
        @{ T = '在「打卡」页的空白行直接输入事项名称即可新增；不想用了就清空事项名称（不要删除整行，以免破坏日期列）。建议同时不超过 12 项。'; S = 13; H = 60 },
        @{ T = '四、完成与删除'; S = 12; H = 26 },
        @{ T = '一件事感觉差不多了：把「状态」改为「已完成」，它停止统计，并自动出现在「完成记录」。实在没时间做：状态改为「已删除」，进入「回收站」并记录删除日期。想恢复：把状态改回「进行中」即可。'; S = 13; H = 90 },
        @{ T = '五、补签与锁定'; S = 12; H = 26 },
        @{ T = '过去 7 天内的空白可以直接选 🩹补签（窗口可在 config.json 改）。超过 7 天仍为空白，评估时按未做计算，并列入「超期锁定」。补签不会改变原始完成率，只会提高含补签完成率。'; S = 13; H = 78 },
        @{ T = '六、评估在哪里看'; S = 12; H = 26 },
        @{ T = '「事项评估」：每项近 7 天完成率、连续完成/中断、判定和建议。「每日汇总」：今天、昨天、近 7 天的完成率与预警名单。「完成记录」「回收站」：已结束、已删除事项。'; S = 13; H = 78 },
        @{ T = '七、判定规则'; S = 12; H = 26 },
        @{ T = '稳定坚持：近 7 天原始完成 ≥6 天；正常：4–5 天；懈怠：2–3 天或连续中断 2 天；基本没做：≤1 天或连续中断 3 天。新事项不足 4 天时标为「待观察」，避免误判。'; S = 13; H = 80 },
        @{ T = '八、提醒与补发'; S = 12; H = 26 },
        @{ T = '每天 22:00 右下角卡片提醒，不抢焦点，点击可打开表格。电脑关机或休眠错过时，登录后会自动补发一次；同一天只提醒一次。电脑需保持开机且未休眠。'; S = 13; H = 65 },
        @{ T = '九、手机查看'; S = 12; H = 26 },
        @{ T = '把 checkin.xlsx 或自动生成的 report.html 发到手机（微信、网盘等）即可查看。手机端建议只读，避免与电脑同时编辑产生冲突。'; S = 13; H = 60 },
        @{ T = '十、备份与注意'; S = 12; H = 26 },
        @{ T = '每次脚本写回前会自动备份到 backup 文件夹（保留最近 14 份）。不要删除第 1 行隐藏的日期行，不要在 I 列之后插入非日期列。表格正被 Excel 打开时，脚本会跳过写回，下次打开或提醒时再补。'; S = 13; H = 95 },
        @{ T = '十一、修改补签窗口'; S = 12; H = 26 },
        @{ T = '用记事本打开 config.json，把 graceDays 改成 3 或 10，保存后下次运行生效。'; S = 13; H = 30 }
    )
    $row = 1
    foreach ($it in $items) {
        Add-XRow $sh $row -Height $it.H
        Add-XCell $sh "A$row" $it.T $it.S
        $sh.Merges += "A$row:H$row"
        $row++
    }
    return $sh
}

function New-CheckinWorkbook {
    param($Model, $Stats, $Config, [datetime]$AsOf, [string]$Path)
    $sheets = @(
        (New-CheckinSheet -Model $Model -Stats $Stats -Config $Config -AsOf $AsOf),
        (New-EvalSheet -Model $Model -Stats $Stats -AsOf $AsOf),
        (New-SummarySheet -Model $Model -Stats $Stats -AsOf $AsOf -Config $Config),
        (New-CompletedSheet -Model $Model -Stats $Stats -AsOf $AsOf),
        (New-RecycleSheet -Model $Model -Stats $Stats -AsOf $AsOf),
        (New-HelpSheet -AsOf $AsOf -Config $Config)
    )
    New-XlsxFile -Path $Path -Sheets $sheets -Title '每日打卡与完成评估'
}
