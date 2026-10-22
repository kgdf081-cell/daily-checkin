param(
    [switch]$Force,
    [switch]$DryRun,
    [switch]$NoWindow,
    [switch]$NoState,
    [ValidateSet('balloon', 'window')]
    [string]$Mode,
    [int]$KeepSeconds = 0,
    [string]$SimulateNow
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\run.ps1')
. (Join-Path $PSScriptRoot 'lib\sound.ps1')
Write-Log ("提醒脚本启动：交互={0}，Force={1}，Mode={2}" -f [Environment]::UserInteractive, $Force, $Mode)

function Show-CheckinBalloon {
    param([string]$Title, [string]$Text, [int]$Seconds)
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $global:CheckinOpenPath = $script:DataFile
    $global:CheckinBalloon = New-Object System.Windows.Forms.NotifyIcon
    $global:CheckinBalloon.Icon = [System.Drawing.SystemIcons]::Information
    $global:CheckinBalloon.Text = '每日打卡'
    $global:CheckinBalloon.BalloonTipTitle = $Title
    $global:CheckinBalloon.BalloonTipText = $Text
    $global:CheckinBalloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
    $global:CheckinBalloon.Add_BalloonTipClicked({
        try { Start-Process -FilePath $global:CheckinOpenPath } catch { }
    })
    $global:CheckinBalloon.Visible = $true
    $global:CheckinBalloon.ShowBalloonTip([Math]::Max(10, $Seconds) * 1000)
    if ($script:CheckinPlaySound) { Play-CheckinSound }

    $global:CheckinTimer = New-Object System.Windows.Forms.Timer
    $global:CheckinTimer.Interval = [Math]::Max(10, $Seconds) * 1000
    $global:CheckinTimer.Add_Tick({
        try { $global:CheckinTimer.Stop() } catch { }
        [System.Windows.Forms.Application]::ExitThread()
    })
    $global:CheckinTimer.Start()
    [System.Windows.Forms.Application]::Run()
    $global:CheckinBalloon.Visible = $false
    $global:CheckinBalloon.Dispose()
}

function Show-CheckinWindow {
    param([string]$Title, [string]$Text, [int]$Seconds)
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    $global:CheckinOpenPath = $script:DataFile
    $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        Topmost="True" ShowInTaskbar="False" ShowActivated="False"
        WindowStartupLocation="Manual" Width="440" SizeToContent="Height" ResizeMode="NoResize">
  <Border Background="#F21F2937" CornerRadius="14">
    <Grid>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="7"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <Border Background="#FF60A5FA" CornerRadius="14,0,0,14"/>
      <StackPanel Grid.Column="1" Margin="20,16,20,16">
        <TextBlock x:Name="Head" Foreground="#FF9CA3AF" FontSize="13" Margin="0,0,0,6"/>
        <TextBlock x:Name="Body" Foreground="#FFF9FAFB" FontSize="14" TextWrapping="Wrap" LineHeight="22"/>
        <Grid Margin="0,12,0,0">
          <TextBlock Text="点击打开打卡表" Foreground="#FF93C5FD" FontSize="12"/>
          <TextBlock Text="Esc 关闭" Foreground="#FF6B7280" FontSize="11" HorizontalAlignment="Right"/>
        </Grid>
      </StackPanel>
    </Grid>
  </Border>
</Window>
'@
    $reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
    $win = [Windows.Markup.XamlReader]::Load($reader)
    $win.FindName('Head').Text = $Title
    $win.FindName('Body').Text = $Text
    $wa = [System.Windows.SystemParameters]::WorkArea
    $win.Left = $wa.Right - 464
    $win.Top = $wa.Bottom - 260
    $win.Add_ContentRendered({
        $wa2 = [System.Windows.SystemParameters]::WorkArea
        $win.Left = $wa2.Right - $win.ActualWidth - 24
        $win.Top = $wa2.Bottom - $win.ActualHeight - 24
    })
    $win.Add_MouseLeftButtonUp({
        try { Start-Process -FilePath $global:CheckinOpenPath } catch { }
        $win.Close()
    })
    $win.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Escape) { $win.Close() }
    })
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromSeconds([Math]::Max(8, $Seconds))
    $timer.Add_Tick({ $timer.Stop(); $win.Close() })
    $timer.Start()
    Write-Log '提醒卡片已显示（右下角）'
    [void]$win.ShowDialog()
}

function Show-CheckinToast {
    param([string]$Title, [string]$Text, [int]$Seconds)
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    if (-not ('CheckinToast' -as [type])) {
        $code = @'
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

public class CheckinToast : Form {
    public Label HeadLabel;
    public Label BodyLabel;
    public Label HintLabel;
    public event EventHandler ToastClick;

    public CheckinToast() {
        FormBorderStyle = FormBorderStyle.None;
        StartPosition = FormStartPosition.Manual;
        ShowInTaskbar = false;
        TopMost = true;
        BackColor = Color.FromArgb(31, 41, 55);
        Size = new Size(440, 190);
        DoubleBuffered = true;

        Panel accent = new Panel();
        accent.BackColor = Color.FromArgb(96, 165, 250);
        accent.Location = new Point(0, 0);
        accent.Size = new Size(6, 190);

        HeadLabel = new Label();
        HeadLabel.ForeColor = Color.FromArgb(156, 163, 175);
        HeadLabel.Font = new Font("Microsoft YaHei", 9F);
        HeadLabel.AutoSize = true;
        HeadLabel.Location = new Point(22, 14);
        HeadLabel.BackColor = Color.Transparent;

        BodyLabel = new Label();
        BodyLabel.ForeColor = Color.FromArgb(249, 250, 251);
        BodyLabel.Font = new Font("Microsoft YaHei", 10F);
        BodyLabel.AutoSize = false;
        BodyLabel.Location = new Point(22, 40);
        BodyLabel.Size = new Size(396, 104);
        BodyLabel.BackColor = Color.Transparent;

        HintLabel = new Label();
        HintLabel.ForeColor = Color.FromArgb(147, 197, 253);
        HintLabel.Font = new Font("Microsoft YaHei", 8.5F);
        HintLabel.AutoSize = true;
        HintLabel.Text = "点击打开打卡表 · 稍后自动关闭";
        HintLabel.Location = new Point(22, 156);
        HintLabel.BackColor = Color.Transparent;

        Controls.Add(accent);
        Controls.Add(HeadLabel);
        Controls.Add(BodyLabel);
        Controls.Add(HintLabel);

        EventHandler click = (s, e) => { if (ToastClick != null) ToastClick(this, EventArgs.Empty); };
        this.Click += click;
        foreach (Control c in Controls) c.Click += click;
    }

    protected override bool ShowWithoutActivation { get { return true; } }

    protected override CreateParams CreateParams {
        get {
            CreateParams cp = base.CreateParams;
            cp.ExStyle |= 0x08000000;
            cp.ExStyle |= 0x00000008;
            return cp;
        }
    }

    protected override void OnResize(EventArgs e) {
        base.OnResize(e);
        using (GraphicsPath path = new GraphicsPath()) {
            int r = 12;
            Rectangle rect = new Rectangle(0, 0, Width, Height);
            path.AddArc(rect.X, rect.Y, r, r, 180, 90);
            path.AddArc(rect.Right - r, rect.Y, r, r, 270, 90);
            path.AddArc(rect.Right - r, rect.Bottom - r, r, r, 0, 90);
            path.AddArc(rect.X, rect.Bottom - r, r, r, 90, 90);
            path.CloseFigure();
            Region = new Region(path);
        }
    }
}
'@
        Add-Type -TypeDefinition $code -ReferencedAssemblies System.Windows.Forms, System.Drawing
    }
    $global:CheckinOpenPath = $script:DataFile
    $global:CheckinToastClicked = $false
    $form = New-Object CheckinToast
    $form.HeadLabel.Text = $Title
    $form.BodyLabel.Text = $Text
    $wa = [System.Windows.Forms.Screen]::FromPoint([System.Windows.Forms.Cursor]::Position).WorkingArea
    $form.Left = [Math]::Max($wa.Left + 10, $wa.Right - $form.Width - 24)
    $form.Top = [Math]::Max($wa.Top + 10, $wa.Bottom - $form.Height - 24)
    $form.add_ToastClick({
        if ($global:CheckinToastClicked) { return }
        $global:CheckinToastClicked = $true
        try { Start-Process -FilePath $global:CheckinOpenPath } catch { }
        $form.Close()
    })
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = [Math]::Max(8, $Seconds) * 1000
    $timer.Add_Tick({ $timer.Stop(); $form.Close() })
    $timer.Start()
    Write-Log '提醒卡片已显示（右下角）'
    $form.Show()
    if ($script:CheckinPlaySound) { Play-CheckinSound }
    [System.Windows.Forms.Application]::Run($form)
}

$now = if ($SimulateNow) { [datetime]::Parse($SimulateNow) } else { Get-Date }
$asOf = $now.Date
$config = Get-AppConfig
$state = Get-StateData
$todayKey = $asOf.ToString('yyyy-MM-dd')

$result = $null
try {
    $result = Invoke-CheckinEvaluate -AsOf $asOf -Write $true
}
catch {
    Write-Log "提醒流程异常：$($_.Exception.Message)" 'ERROR'
    exit 1
}

$payload = Get-ReminderPayload $result
$stale = $false
if (-not $result.Ok) {
    $cache = Get-PayloadCache
    if ($cache -and $cache.payload) {
        $payload = $cache.payload
        $stale = $true
        $payload | Add-Member -NotePropertyName text -NotePropertyValue ("表格正在 Excel 中打开，以下是 {0} 的快照。`n请保存并关闭表格后再刷新。`n`n{1}" -f $cache.savedAt, $payload.text) -Force
    }
}
if ($DryRun) {
    Write-Output ("[{0}] {1}`n{2}" -f $payload.title, $payload.text, $(if ($result.Written) { 'written' } else { 'read-only' }))
    exit 0
}

if (-not $Force) {
    if ($state.lastReminderDate -eq $todayKey) {
        Write-Log "今天已提醒过，跳过（$($state.lastReminderAt)）"
        exit 0
    }
    $parts = ([string]$config.reminderTime).Split(':')
    $scheduled = $asOf.AddHours([int]$parts[0]).AddMinutes([int]$parts[1])
    if ($now -lt $scheduled.AddMinutes(-2)) {
        Write-Log "未到提醒时间（$($config.reminderTime)），跳过通知"
        exit 0
    }
}

$isCatchup = $false
$parts = ([string]$config.reminderTime).Split(':')
$scheduled = $asOf.AddHours([int]$parts[0]).AddMinutes([int]$parts[1])
if ($now -gt $scheduled.AddMinutes(10)) { $isCatchup = $true }

$title = '每日打卡 · ' + $config.reminderTime
if ($isCatchup) { $title = '每日打卡 · 补发提醒' }
if ($stale) { $title = '每日打卡 · 请保存表格' }

Write-Log ("发送提醒（{0}）：今日待办 {1}，昨日待补 {2}，即将超期 {3}，写回={4}" -f $title, $payload.todayPending, $payload.yesterdayPending, $payload.expiring, $result.Written)

if (-not $NoWindow) {
    $mode = if ($Mode) { $Mode } else { [string]$config.notifyMode }
    $keep = if ($KeepSeconds -gt 0) { $KeepSeconds } else { [int]$config.keepSeconds }
    $script:CheckinPlaySound = ($config.playSound -ne $false)
    $script:CheckinSoundPlayed = $false
    $script:CheckinVoiceFile = if ($config.voiceFile) { [string]$config.voiceFile } else { 'voice.wav' }
    $script:CheckinVoiceVolume = if ($null -ne $config.voiceVolume) { [int]$config.voiceVolume } else { 70 }
    try {
        if ($mode -eq 'window') {
            Show-CheckinToast -Title $title -Text $payload.text -Seconds $keep
        }
        else {
            Show-CheckinBalloon -Title $title -Text $payload.text -Seconds $keep
        }
    }
    catch {
        Write-Log "通知显示失败（$mode）：$($_.Exception.Message)" 'ERROR'
        if ($script:CheckinPlaySound) { Play-CheckinSound }
        if ($mode -eq 'window') {
            try {
                Show-CheckinBalloon -Title $title -Text $payload.text -Seconds $keep
            }
            catch {
                Write-Log "气泡备用方案也失败：$($_.Exception.Message)" 'ERROR'
            }
        }
    }
}

if (-not $NoState) {
    $state.lastReminderDate = $todayKey
    $state.lastReminderAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Save-StateData $state
}
