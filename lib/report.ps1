# 生成手机友好的只读报告 report.html

function Get-Html {
    param($Text)
    if ($null -eq $Text) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Text)
}

function Get-RateText {
    param($Value)
    if ($null -eq $Value) { return '—' }
    return ('{0:P0}' -f [double]$Value)
}

function Get-RateWidth {
    param($Value)
    if ($null -eq $Value) { return 0 }
    $w = [int][math]::Round([double]$Value * 100)
    if ($w -lt 0) { return 0 }
    if ($w -gt 100) { return 100 }
    return $w
}

function Write-HtmlReport {
    param($Stats, [datetime]$AsOf, [string]$Path)

    $payload = Get-ReminderPayload ([pscustomobject]@{ Stats = $Stats; Written = $true })
    $weekday = @('周日', '周一', '周二', '周三', '周四', '周五', '周六')[[int]$AsOf.DayOfWeek]
    $donePct = if ($Stats.Today.Due -gt 0) { [int][math]::Round($Stats.Today.Done * 100 / $Stats.Today.Due) } else { 0 }
    $statusClass = @{
        '稳定坚持' = 'ok'
        '正常'     = 'good'
        '懈怠'     = 'warn'
        '基本没做' = 'bad'
        '待观察'   = 'muted'
        '暂无数据' = 'muted'
    }

    $itemsHtml = New-Object System.Text.StringBuilder
    foreach ($s in $Stats.Active) {
        $cls = $statusClass[$s.StatusText]
        if (-not $cls) { $cls = 'muted' }
        $rate = Get-RateText $s.RawRate
        $patchRate = Get-RateText $s.PatchRate
        [void]$itemsHtml.Append('<div class="row">')
        [void]$itemsHtml.Append('<div class="row-main"><div class="name">' + (Get-Html $s.Item.Name) + '</div>')
        [void]$itemsHtml.Append('<div class="meta">' + (Get-Html $s.Item.Type) + ' · 权重' + (Get-Html $s.Item.Weight) + ' · 应做 ' + $s.Due + ' 天</div></div>')
        [void]$itemsHtml.Append('<div class="row-side"><span class="pill ' + $cls + '">' + (Get-Html $s.StatusText) + '</span>')
        [void]$itemsHtml.Append('<div class="rate">' + $rate + '<small>原始</small></div>')
        [void]$itemsHtml.Append('<div class="rate small">含补签 ' + $patchRate + '</div></div></div>')
        [void]$itemsHtml.Append('<div class="mini-bar"><span style="width:' + (Get-RateWidth $s.RawRate) + '%"></span></div>')
    }
    if ($Stats.Active.Count -eq 0) { [void]$itemsHtml.Append('<p class="empty">暂无进行中的事项。</p>') }

    $trendHtml = New-Object System.Text.StringBuilder
    foreach ($t in $Stats.Trend) {
        [void]$trendHtml.Append('<div class="trend-row"><div class="trend-date">' + $t.Date.ToString('MM-dd') + '</div>')
        [void]$trendHtml.Append('<div class="trend-bar"><span style="width:' + (Get-RateWidth $t.RawRate) + '%"></span></div>')
        [void]$trendHtml.Append('<div class="trend-val">' + (Get-RateText $t.RawRate) + '</div></div>')
    }

    $alertHtml = New-Object System.Text.StringBuilder
    foreach ($pair in @(
            @{ T = '待补签'; L = $Stats.PendingList; C = 'warn' },
            @{ T = '即将超期'; L = $Stats.ExpiringList; C = 'warn' },
            @{ T = '超期锁定'; L = $Stats.LockedList; C = 'bad' },
            @{ T = '懈怠预警'; L = $Stats.SlackList; C = 'warn' },
            @{ T = '基本没做'; L = $Stats.NoneList; C = 'bad' }
        )) {
        $arr = @($pair.L | Where-Object { $_ })
        if ($arr.Count -eq 0) { continue }
        [void]$alertHtml.Append('<div class="alert ' + $pair.C + '"><b>' + $pair.T + '</b><span>' + (Get-Html ($arr -join '、')) + '</span></div>')
    }
    if ($alertHtml.Length -eq 0) { [void]$alertHtml.Append('<p class="empty">今天没有需要特别提醒的事项。</p>') }

    $chips = New-Object System.Text.StringBuilder
    foreach ($k in @('稳定坚持', '正常', '懈怠', '基本没做', '待观察')) {
        $cls = $statusClass[$k]
        [void]$chips.Append('<span class="chip ' + $cls + '">' + $k + ' <b>' + $Stats.Counts[$k] + '</b></span>')
    }

    $yRate = Get-RateText $Stats.Yesterday.RawRate
    $yPatch = Get-RateText $Stats.Yesterday.PatchRate
    $wRate = Get-RateText $Stats.Week.RawRate
    $wPatch = Get-RateText $Stats.Week.PatchRate
    $wWeighted = Get-RateText $Stats.Week.WeightedRate

    $html = @"
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>每日打卡 · $($AsOf.ToString('yyyy-MM-dd'))</title>
<style>
:root{--bg:#f4f6f8;--card:#fff;--ink:#1f2937;--muted:#6b7280;--line:#e5e7eb;--ok:#16a34a;--good:#2563eb;--warn:#d97706;--bad:#dc2626}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif;line-height:1.55}
main{max-width:760px;margin:0 auto;padding:18px 14px 48px}
header{background:linear-gradient(135deg,#1f4e79,#2e75b6);color:#fff;border-radius:16px;padding:22px 20px 18px;box-shadow:0 8px 24px rgba(31,78,121,.18)}
.date{font-size:13px;opacity:.85}
h1{margin:4px 0 6px;font-size:26px;letter-spacing:.5px}
.sub{margin:0;font-size:14px;opacity:.92}
.hero-bar{height:10px;background:rgba(255,255,255,.25);border-radius:6px;margin-top:14px;overflow:hidden}
.hero-bar span{display:block;height:100%;background:#7dd3fc;border-radius:6px}
.cards{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin:14px 0}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px 12px 10px}
.card .k{font-size:12px;color:var(--muted)}
.card .v{font-size:24px;font-weight:700;margin:2px 0}
.card .s{font-size:12px;color:var(--muted)}
section{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:16px;margin-top:12px}
h2{font-size:15px;margin:0 0 12px;color:#374151}
.chips{display:flex;flex-wrap:wrap;gap:8px}
.chip{font-size:13px;padding:4px 10px;border-radius:999px;background:#f3f4f6;color:#374151}
.chip.ok{background:#dcfce7;color:#166534}.chip.good{background:#dbeafe;color:#1e40af}
.chip.warn{background:#fef3c7;color:#92400e}.chip.bad{background:#fee2e2;color:#991b1b}
.chip.muted{background:#f3f4f6;color:#4b5563}
.row{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:10px 0;border-bottom:1px dashed var(--line)}
.row:last-child{border-bottom:0}
.name{font-weight:600;font-size:15px}
.meta{font-size:12px;color:var(--muted);margin-top:2px}
.row-side{display:flex;align-items:center;gap:10px;flex-shrink:0}
.pill{font-size:12px;padding:3px 9px;border-radius:999px;white-space:nowrap}
.pill.ok{background:#dcfce7;color:#166534}.pill.good{background:#dbeafe;color:#1e40af}
.pill.warn{background:#fef3c7;color:#92400e}.pill.bad{background:#fee2e2;color:#991b1b}
.pill.muted{background:#f3f4f6;color:#4b5563}
.rate{font-size:16px;font-weight:700;min-width:58px;text-align:right}
.rate small{display:block;font-size:10px;color:var(--muted);font-weight:400}
.rate.small{font-size:11px;font-weight:400;color:var(--muted);min-width:auto}
.mini-bar{height:4px;background:#eef1f4;border-radius:4px;margin:-2px 0 8px;overflow:hidden}
.mini-bar span{display:block;height:100%;background:linear-gradient(90deg,#60a5fa,#2563eb);border-radius:4px}
.alert{display:flex;gap:10px;align-items:flex-start;padding:8px 10px;border-radius:9px;margin:7px 0;font-size:13px}
.alert.warn{background:#fffbeb;color:#92400e}.alert.bad{background:#fef2f2;color:#991b1b}
.alert b{white-space:nowrap}
.trend-row{display:flex;align-items:center;gap:10px;margin:7px 0}
.trend-date{width:46px;font-size:12px;color:var(--muted)}
.trend-bar{flex:1;height:8px;background:#eef1f4;border-radius:5px;overflow:hidden}
.trend-bar span{display:block;height:100%;background:#93c5fd;border-radius:5px}
.trend-val{width:44px;text-align:right;font-size:12px;color:#374151}
.empty{color:var(--muted);font-size:13px;margin:6px 0}
footer{text-align:center;color:#9ca3af;font-size:12px;margin-top:18px}
@media (max-width:560px){.cards{grid-template-columns:1fr 1fr}.card:last-child{grid-column:span 2}}
</style>
</head>
<body>
<main>
<header>
  <div class="date">$(Get-Html ($AsOf.ToString('yyyy-MM-dd') + ' · ' + $weekday))</div>
  <h1>每日打卡</h1>
  <p class="sub">今天已完成 $($Stats.Today.Done) / $($Stats.Today.Due) 项 · 待办 $($Stats.Today.Pending) 项</p>
  <div class="hero-bar"><span style="width:$donePct%"></span></div>
</header>

<div class="cards">
  <div class="card"><div class="k">昨日原始率</div><div class="v">$yRate</div><div class="s">含补签 $yPatch</div></div>
  <div class="card"><div class="k">近 7 天原始率</div><div class="v">$wRate</div><div class="s">含补签 $wPatch</div></div>
  <div class="card"><div class="k">近 7 天加权率</div><div class="v">$wWeighted</div><div class="s">高3 / 中2 / 低1</div></div>
</div>

<section>
  <h2>状态分布</h2>
  <div class="chips">$($chips.ToString())</div>
</section>

<section>
  <h2>事项 · 近 7 天</h2>
  $($itemsHtml.ToString())
</section>

<section>
  <h2>预警</h2>
  $($alertHtml.ToString())
</section>

<section>
  <h2>近 7 天趋势</h2>
  $($trendHtml.ToString())
</section>

<footer>生成于 $((Get-Date).ToString('yyyy-MM-dd HH:mm')) · 数据来自 checkin.xlsx</footer>
</main>
</body>
</html>
"@
    [IO.File]::WriteAllText($Path, $html, (New-Object System.Text.UTF8Encoding($false)))
}
