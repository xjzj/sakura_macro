
############################################################
# ShowAlarmWindow.ps1
############################################################

#========================================================
# WPF Alarm窓口を作成表示 
#
# パラメータ：
#   $taskName   タスク名称
#   $taskTime   タスク時間
#   $excelColor Excel Cellの背景色
#========================================================

param(
	[string]$ReadyFile,
	[string]$TaskName,
	[datetime]$taskTime,
	[string]$ColorStr
)


Write-Host "ReadyFile = $ReadyFile"
Write-Host "TaskName = $TaskName"
Write-Host "taskTime = $taskTime"
Write-Host "Color = $Color"

while (-not (Test-Path $ReadyFile)) {
	Start-Sleep -Milliseconds 100
}

Write-Host "PROGRAM START"
Write-Host "PID       = $PID"
Write-Host "Thread ID = $([System.Threading.Thread]::CurrentThread.ManagedThreadId)"
Write-Host "Apartment State = $([System.Threading.Thread]::CurrentThread.GetApartmentState())"

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase


#   function ShowAlarmWindow(
#   	$taskName,
#   	$taskTime,
#   	$excelColor
#   )

#----------------------------------------------------
# Window
#----------------------------------------------------

$window = New-Object System.Windows.Window

$window.Title = "タスクアラーム"

$window.Width = 500
$window.Height = 300

$window.WindowStartupLocation =
[System.Windows.WindowStartupLocation]::CenterScreen

$window.Topmost = $true

$window.ResizeMode =
[System.Windows.ResizeMode]::NoResize


#----------------------------------------------------
# 背景色
#----------------------------------------------------

# $Color = Convert-ExcelColor $excelColor

$Color = [System.Windows.Media.ColorConverter]::ConvertFromString($ColorStr)

$window.Background =
New-Object System.Windows.Media.SolidColorBrush($Color)


#----------------------------------------------------
# Grid
#----------------------------------------------------

$grid = New-Object System.Windows.Controls.Grid

$window.Content = $grid


#----------------------------------------------------
# Grid 行
#----------------------------------------------------

$row1 = New-Object System.Windows.Controls.RowDefinition
$row1.Height =
New-Object System.Windows.GridLength(60)

$row2 = New-Object System.Windows.Controls.RowDefinition
$row2.Height =
New-Object System.Windows.GridLength(
	1,
	[System.Windows.GridUnitType]::Star
)

$row3 = New-Object System.Windows.Controls.RowDefinition
$row3.Height =
New-Object System.Windows.GridLength(70)

$grid.RowDefinitions.Add($row1)
$grid.RowDefinitions.Add($row2)
$grid.RowDefinitions.Add($row3)


#----------------------------------------------------
# タイトル
#----------------------------------------------------

$title = New-Object System.Windows.Controls.TextBlock

$title.Text = "タスクアラーム"

$title.FontSize = 28
$title.FontWeight = "Bold"

$title.HorizontalAlignment =
[System.Windows.HorizontalAlignment]::Center

$title.VerticalAlignment =
[System.Windows.VerticalAlignment]::Center

[System.Windows.Controls.Grid]::SetRow($title, 0)

$grid.Children.Add($title)

#----------------------------------------------------
# タスク内容
#----------------------------------------------------

$text = New-Object System.Windows.Controls.TextBlock

$text.Text =
"$TaskName`n`n$($taskTime.ToString('yyyy-MM-dd HH:mm:ss'))"

$text.FontSize = 24

$text.HorizontalAlignment =
[System.Windows.HorizontalAlignment]::Center

$text.VerticalAlignment =
[System.Windows.VerticalAlignment]::Center

$text.TextAlignment =
[System.Windows.TextAlignment]::Center

[System.Windows.Controls.Grid]::SetRow($text, 1)

$grid.Children.Add($text)

# 保存到 Script 作用域
$script:AlarmWindow = $window

#----------------------------------------------------
# 確定ボタン
#----------------------------------------------------

$button = New-Object System.Windows.Controls.Button

$button.Content = "確定"

$button.Width = 120
$button.Height = 40

$button.FontSize = 18

$button.HorizontalAlignment =
[System.Windows.HorizontalAlignment]::Center

$button.VerticalAlignment =
[System.Windows.VerticalAlignment]::Center

[System.Windows.Controls.Grid]::SetRow($button, 2)

$grid.Children.Add($button)

$now = Get-Date
Write-Host   ("Alarm  $($now.ToString('yyyy-MM-dd HH:mm:ss')) => " + "$TaskName  $($taskTime.ToString('yyyy-MM-dd HH:mm:ss'))" + " ....")

#----------------------------------------------------
# 確定をクリック
#----------------------------------------------------
$handler = {
	# $window.Close()
	$script:AlarmWindow.Close()

	# 关闭后取消 Script 作用域中的引用
	$script:AlarmWindow = $null
}

$button.Add_Click($handler)


#----------------------------------------------------
# 窓口を表示
#----------------------------------------------------
try
{
	$window.ShowDialog() | Out-Null
	# $window.Show() | Out-Null

	exit 0
}
catch {
	Write-Host "========================================"
	Write-Host "ShowAlarmWindow ERROR"
	Write-Host "Error:"
	Write-Host $_
	Write-Host ""
	Write-Host "Stack:"
	Write-Host $_.ScriptStackTrace
	Write-Host "Window exception:"
	Write-Host $_.Exception.ToString()
	Write-Host "========================================"

	# 根据异常处理
	if ($_.Exception.HResult -eq [int]0x800706BA) {
		exit 2
	}

	exit 1
}
finally
{
	$button.Remove_Click($handler)

	if ($window)
	{
		$window.Close()
		$button = $null
		$text   = $null
		$title  = $null
		$grid   = $null
		$window = $null
	}
}


