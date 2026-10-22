$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$generateOnly = ($args -contains '-GenerateOnly')

function Write-Step {
    param([string]$Text)
    Write-Host $Text -ForegroundColor Cyan
}

$configPath = Join-Path $root 'config.json'
$reminder = '22:00'
if (Test-Path -LiteralPath $configPath) {
    try {
        $cfg = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($cfg.reminderTime) { $reminder = [string]$cfg.reminderTime }
    }
    catch { }
}

$today = (Get-Date).Date
$timeParts = $reminder.Split(':')
$start = $today.AddHours([int]$timeParts[0]).AddMinutes([int]$timeParts[1])
if ($start -lt (Get-Date)) { $start = $start.AddDays(1) }

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$sid = $identity.User.Value
$userName = $identity.Name
$notify = Join-Path $root 'notify.ps1'
$taskName = '每日打卡提醒'
$xmlPath = Join-Path $root 'tasks\DailyCheckin.xml'

$xml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Author>$userName</Author>
    <Description>每天 $reminder 提醒打卡，并刷新完成评估。错过时在登录后补发一次。</Description>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>$($start.ToString('yyyy-MM-ddTHH:mm:ss'))</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByDay>
        <DaysInterval>1</DaysInterval>
      </ScheduleByDay>
    </CalendarTrigger>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <Delay>PT2M</Delay>
      <UserId>$sid</UserId>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>$sid</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT10M</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>powershell.exe</Command>
      <Arguments>-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "$notify"</Arguments>
      <WorkingDirectory>$root</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"@

New-Item -ItemType Directory -Force -Path (Split-Path $xmlPath -Parent) | Out-Null
[IO.File]::WriteAllText($xmlPath, $xml, (New-Object System.Text.UnicodeEncoding($false, $true)))

Write-Step "任务文件已生成：$xmlPath"
if ($generateOnly) {
    Write-Host '仅生成 XML，未注册任务。' -ForegroundColor Yellow
    exit 0
}
Write-Step "正在注册计划任务：$taskName"
$create = & schtasks.exe /Create /TN $taskName /XML $xmlPath /F 2>&1
Write-Host $create

if ($LASTEXITCODE -ne 0) {
    Write-Host ''
    Write-Host '自动注册失败。可以手动导入：' -ForegroundColor Yellow
    Write-Host '  1) 打开「任务计划程序」' -ForegroundColor Yellow
    Write-Host '  2) 右侧点「导入任务」，选择上面的 DailyCheckin.xml' -ForegroundColor Yellow
    Write-Host '  3) 确认后保存即可' -ForegroundColor Yellow
    exit 1
}

$query = & schtasks.exe /Query /TN $taskName /FO LIST 2>&1
$query | ForEach-Object { Write-Host $_ }

Write-Host ''
Write-Host '完成：每天 ' + $reminder + ' 会自动提醒；登录后若当天还没提醒过，会补发一次。' -ForegroundColor Green
