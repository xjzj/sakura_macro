
############################################################
# task.ps1
############################################################

Add-Type -AssemblyName PresentationFramework

#-----------------------------------------------------------
# 設定
#-----------------------------------------------------------

$ExcelFile = "$PSScriptRoot\task.xlsx"

# チェック間隔（秒）
$CheckInterval = 5

#-----------------------------------------------------------
# グローバル変数
#-----------------------------------------------------------

$global:date = (Get-Date 0).Date

$global:Excel = $null
$global:TaskOrder = @()
$global:Tasks = @{}
$global:LastWriteTime = Get-Date 0

#-----------------------------------------------------------
# 初期化 Excel COM
#-----------------------------------------------------------
function InitExcel()
{
	if($global:Excel -ne $null)
	{
		return
	}

	$global:Excel = New-Object -ComObject Excel.Application

	$global:Excel.Visible = $false
	$global:Excel.DisplayAlerts = $false
	$global:Excel.ScreenUpdating = $false
}

#-----------------------------------------------------------
# 解放 Excel COM
#-----------------------------------------------------------
function ReleaseExcel()
{
	if($global:Excel -ne $null)
	{
		$global:Excel.Quit()

		[System.Runtime.InteropServices.Marshal]::ReleaseComObject($global:Excel) | Out-Null

		$global:Excel = $null

		[GC]::Collect()
		[GC]::WaitForPendingFinalizers()
	}
}

function SetFlg($test, $t, $f)
{
	if($test)
	{
		return $t
	}
	else {
		return $f
	}
}

function Get-DisplayWidth($str)
{
	$width = 0

	foreach($c in $str.ToCharArray())
	{
		if([int][char]$c -gt 0xFF)
		{
			$width += 2
		}
		else
		{
			$width += 1
		}
	}

	return $width
}

function Pad-DisplayRight($str, $width)
{
	$currentWidth = Get-DisplayWidth $str

	if($currentWidth -lt $width)
	{
		return $str + (" " * ($width - $currentWidth))
	}

	return $str
}

#-----------------------------------------------------------
# 読込 Excel
#-----------------------------------------------------------
function LoadExcel()
{
	Write-Host ""
	Write-Host "Loading Excel..."

	InitExcel

	$tasks = @()

	$workbook = $null
	$sheet = $null

	try
	{
		$workbook = $global:Excel.Workbooks.Open(
			$ExcelFile,
			0,
			$true          # ReadOnly
		)

		$sheet = $workbook.Worksheets.Item(1)

		$cnt = 1
		$row = 2

		while($true)
		{
			$name = $sheet.Cells.Item($row,1).Text.Trim()

			if([string]::IsNullOrWhiteSpace($name))
			{
				break
			}

			$timeText = $sheet.Cells.Item($row,2).Text.Trim()

			if([string]::IsNullOrWhiteSpace($timeText))
			{
				$row++
				continue
			}

			try
			{
				$dt = [datetime]$timeText
			}
			catch
			{
				Write-Host "Invalid datetime : $timeText"

				$row++
				continue
			}
			$everyday = $sheet.Cells.Item($row,3).Value2

			$status = $sheet.Cells.Item($row,4).Value2

			if($null -ne $status -and
					-not [string]::IsNullOrWhiteSpace([string]$status))
			{
				# Excel 4列に内容がある
				# Excel 状態が最優先

				$alarmed = ([int]$status -eq 1)
			}
			else
			{
				# Excel 4列に内容が空
				# トライ：実行中の状態を継承

				$alarmed = $false

				if($global:Tasks.ContainsKey($name))
				{
					$oldTask = $global:Tasks[$name]

					# 時間が変わらない場合、実行中の状態を継承
					if($oldTask.Time -eq $dt)
					{
						$alarmed = $oldTask.Alarmed
					}
				}
			}

			$color = $sheet.Cells.Item($row,1).Interior.Color

			$tasks += [PSCustomObject]@{

				num = $cnt

				Row = $row

				Name = $name

				Time = $dt

				everyday = $everyday

				Alarmed = $alarmed

				ExcelColor = $color
			}

			$cnt++
			$row++
		}

		# $global:Tasks = $tasks
		$newTasks = @{}
		$changeTasks = @{}
		$changeOrder = @()

		foreach($task in $tasks)
		{
			$key = $task.Name

			$before = ""
			$after = ""
			if($global:Tasks.ContainsKey($key))
			{
				#
				# 既存タスクの変更状态を反映
				#

				$old = $global:Tasks[$key]
				$change=0

				if($old.Time -ne $task.Time)
				{
					$before = $before + $old.Time.ToString('yyyy/MM/dd HH:mm:ss')
					$after = $after + $task.Time.ToString('yyyy/MM/dd HH:mm:ss')
					$change=1
				}

				if($old.Alarmed -ne $task.Alarmed)
				{
					$before = $before + ' ' + ( SetFlg  $old.Alarmed "○" "-"  )
					$after = $after + ' ' +  ( SetFlg  $task.Alarmed "○" "-"  ) 
					$change=1
				}
				if($change -eq 1)
				{

					$changeOrder += [PSCustomObject]@{
						Name = $task.Name
						Before = $before
						After = $after
					}
				}
			}
			else {
				$after = $task.Time.ToString('yyyy/MM/dd HH:mm:ss') + ' ' + ( SetFlg  $task.Alarmed "○" "-"  )
				$changeOrder += [PSCustomObject]@{
					Name = $task.Name
					Before = $before
					After = $after
				}
			}

			$newTasks[$key] = $task
		}

		if($changeOrder.Count -gt 0)
		{
			$now = Get-Date
			Write-Host "タスクを変更或追加 [$($now.ToString('yyyy-MM-dd HH:mm:ss'))]"
			Write-Host "---------------------------------"
			foreach($task in $changeOrder)
			{
				Write-Host ("{0,-20} {1} => {2}" -f $task.Name,$task.Before,$task.After  )
				$changeTasks[$task.Name]=$task
			}
			Write-Host "---------------------------------"
		}

		$global:TaskOrder = $tasks

		$global:Tasks = $newTasks

		$global:LastWriteTime =
		(Get-Item $ExcelFile).LastWriteTime

		Write-Host "$($tasks.Count) task(s) loaded."
	}
	finally
	{
		if($sheet -ne $null)
		{
			[System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
		}

		if($workbook -ne $null)
		{
			$workbook.Close($false)

			[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
		}

		[GC]::Collect()
		[GC]::WaitForPendingFinalizers()

		ReleaseExcel
	}
	return $changeTasks
}

function SaveAlarmState()
{

	InitExcel

	$workbook = $null
	$sheet = $null

	try
	{
		$workbook = $global:Excel.Workbooks.Open(
			$ExcelFile,
			0,
			$false      # 書込可能
		)

		$sheet = $workbook.Worksheets.Item(1)

		foreach($task in $global:Tasks.Values)
		{

			$sheet.Cells.Item($task.Row,4).Value2 =  [int]$task.Alarmed
			$sheet.Cells.Item($task.Row,2).Value = $task.Time

		}

		$workbook.Save()

	}
	finally
	{
		if($sheet)
		{
			[Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)|Out-Null
		}

		if($workbook)
		{
			$workbook.Close($true)

			[Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)|Out-Null
		}

		[GC]::Collect()
		[GC]::WaitForPendingFinalizers()
		ReleaseExcel
	}
}


Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

function Convert-ExcelColor($color)
{
	$r = $color -band 0xFF
	$g = ($color -shr 8) -band 0xFF
	$b = ($color -shr 16) -band 0xFF

	return [System.Windows.Media.Color]::FromRgb(
		[byte]$r,
		[byte]$g,
		[byte]$b
	)
}


#========================================================
# WPF Alarm窓口を作成表示 
#
# パラメータ：
#   $taskName   タスク名称
#   $taskTime   タスク時間
#   $excelColor Excel Cellの背景色
#========================================================
function ShowAlarmWindow(
	$taskName,
	$taskTime,
	$excelColor
)
{
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

	$color = Convert-ExcelColor $excelColor

	$window.Background =
	New-Object System.Windows.Media.SolidColorBrush($color)


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
	"$taskName`n`n$($taskTime.ToString('yyyy-MM-dd HH:mm:ss'))"

	$text.FontSize = 24

	$text.HorizontalAlignment =
	[System.Windows.HorizontalAlignment]::Center

	$text.VerticalAlignment =
	[System.Windows.VerticalAlignment]::Center

	$text.TextAlignment =
	[System.Windows.TextAlignment]::Center

	[System.Windows.Controls.Grid]::SetRow($text, 1)

	$grid.Children.Add($text)


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
	Write-Host   ("Alarm  $($now.ToString('yyyy-MM-dd HH:mm:ss')) => " + "$taskName  $($taskTime.ToString('yyyy/MM/dd HH:mm:ss'))" + " ....")

	#----------------------------------------------------
	# 確定をクリック
	#----------------------------------------------------
	$handler = {
		$window.Close()
	}

	$button.Add_Click($handler)


	#----------------------------------------------------
	# 窓口を表示
	#----------------------------------------------------
	try
	{
		$window.ShowDialog() | Out-Null
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
}

#-----------------------------------------------------------
# Alarmを出力
#-----------------------------------------------------------
function ShowAlarm($task)
{
	$player = New-Object System.Media.SoundPlayer

	try
	{
		# $player.SoundLocation = "D:\task\alarm.wav"
		$player.SoundLocation = "$env:WINDIR\Media\Alarm01.wav"

		# ループでプレイ
		$player.PlayLooping()

		#------------------------------------------------
		# WPF 窓口を呼ぶ
		#------------------------------------------------

		ShowAlarmWindow `
		$task.Name `
		$task.Time `
		$task.ExcelColor

		# ダイアログを出力
		<#
		[System.Windows.MessageBox]::Show(
		"タスク：`n$($task.Name)`n`n時間：$($task.Time)",
		"タスクAlarm"
		) | Out-Null
		#>
	}
	finally
	{
		# ダイアログを閉じた後、音を停止する
		$player.Stop()

		$player.Dispose()
	}
	<#
	[System.Windows.MessageBox]::Show(
	@"
	タスク：

	$($task.Name)

	時間：

	$($task.Time)

	"@,
	"Alarm")
	#>
}


############################################################
# Excelに変化があるかをチェック
############################################################
function CheckExcelChanged()
{
	if(!(Test-Path $ExcelFile))
	{
		return
	}

	$t = (Get-Item $ExcelFile).LastWriteTime

	if($t -ne $global:LastWriteTime)
	{
		Write-Host ""
		Write-Host "************************************************************************************"
		Write-Host "●Excel changed."
		$now = Get-Date
		Write-Host "[$($now.ToString('yyyy-MM-dd HH:mm:ss'))]"
		Write-Host "************************************************************************************"

		$tasks = LoadExcel

		DumpTasks $tasks
		$global:date = (Get-Date 0).Date
	}
}

############################################################
# タスクの時間をチェック
############################################################
function CheckTasks()
{
	$now = Get-Date

	$everyday_flg  = $false
	if($global:date -ne $now.Date)
	{
		$global:date = $now.Date
		$everyday_flg  = $true
	}


	$change_day_tasks = @()
	$Alarmed_day_tasks = @()

	foreach($task in $global:TaskOrder)
	{

		if( $everyday_flg -and $task.everyday)
		{
			if($global:date -ne $task.Time.Date)
			{
				$newDate = $global:date.Add(
					$task.Time.TimeOfDay
				)
				$before = ""
				$after = ""
				$before = $before + $task.Time.ToString('yyyy/MM/dd HH:mm:ss')
				$after = $after + $newDate.ToString('yyyy/MM/dd HH:mm:ss')
				$task.Time = $newDate

				if( $task.Alarmed -and ($task.Time -ge $now) )
				{
					$before = $before + " ○"
					$after = $after + " -"
					$task.Alarmed  = $false
				}

				$change_day_tasks += [PSCustomObject]@{
					Name = $task.Name

					Before = $before

					After = $after

				}

			}
		}


		if($task.Alarmed)
		{
			continue
		}

		if($now -ge $task.Time)
		{
			$task.Alarmed = $true
			$Alarmed_day_tasks += $task
		}
	}
	if($change_day_tasks.Count -gt 0)
	{
		$tasks = @{}
		Write-Host "************************************************************************************"
		Write-Host "●日付変更"
		Write-Host "[$($now.ToString('yyyy-MM-dd HH:mm:ss'))]"
		Write-Host "************************************************************************************"
		Write-Host "---------------------------------"
		foreach($task in $change_day_tasks)
		{
			Write-Host ("{0,-20} {1} => {2}" -f $task.Name,$task.Before,$task.After  )
			$tasks[$task.Name]=$task
		}
		Write-Host "---------------------------------"
		DumpTasks $tasks
	}
	if($Alarmed_day_tasks.Count -gt 0)
	{
		$tasks = @{}
		Write-Host "************************************************************************************"
		Write-Host "●Alarm"
		Write-Host "[$($now.ToString('yyyy-MM-dd HH:mm:ss'))]"
		Write-Host "************************************************************************************"
		foreach($task in $Alarmed_day_tasks)
		{
			ShowAlarm $task
			$tasks[$task.Name]=$task
		}
		DumpTasks $tasks
	}
}

############################################################
# 現在タスクを表示
############################################################
function DumpTasks($tasks)
{
	Write-Host ""
	$now = Get-Date
	Write-Host "Current Tasks [$($now.ToString('yyyy-MM-dd HH:mm:ss'))]"

	$num = Pad-DisplayRight  "順番" 5
	$op = Pad-DisplayRight  "処理" 7
	$name = Pad-DisplayRight  "タスク" 20
	$time = Pad-DisplayRight  "Time" 25
	$everyday = Pad-DisplayRight  "Everyday" 10
	$alarmed = Pad-DisplayRight  "Alarmed" 10


	$output = $num + $op + $name + $time + $everyday+ $alarmed

	Write-Host $output
	Write-Host "------------------------------------------------------------------------"

	foreach($task in $global:TaskOrder)
	{
		$flag = ""
		if($tasks.ContainsKey($task.Name))
		{
			$flag = "★"
		}
		$num = Pad-DisplayRight  $task.num.ToString() 5
		$op = Pad-DisplayRight  $flag 7
		$name = Pad-DisplayRight  $task.Name 20
		$time = Pad-DisplayRight  $task.Time.ToString("yyyy/MM/dd HH:mm:ss") 25
		$everyday = Pad-DisplayRight ( SetFlg  $task.everyday "◎" "-"  ) 10
		$alarmed = Pad-DisplayRight  ( SetFlg  $task.Alarmed "○" "-"  ) 10
		$output = $num + $op + $name + $time + $everyday+ $alarmed

		Write-Host $output
	}

	Write-Host "------------------------------------------------------------------------"
}

############################################################
# 初期化
############################################################
function Initialize()
{
	if(!(Test-Path $ExcelFile))
	{
		throw "Excel File Not Found.`n$ExcelFile"
	}
	$now = Get-Date
	Write-Host "************************************************************************************"
	Write-Host "●Initialize"
	Write-Host "[$($now.ToString('yyyy-MM-dd HH:mm:ss'))]"
	Write-Host "************************************************************************************"

	# InitExcel

	$tasks = LoadExcel

	DumpTasks $tasks
}


############################################################
# メインループ
############################################################
function Main()
{
	try
	{
		Write-Host ""
		Write-Host "Task monitor started."
		Write-Host "Press Ctrl+C to exit."
		Write-Host ""

		Initialize

		while($true)
		{
			try
			{
				# 检查 Excel 是否变化
				CheckExcelChanged

				# 检查是否有任务到时间
				CheckTasks
			}
			catch
			{
				Write-Host ""
				Write-Host "Check error:"
				Write-Host $_.Exception.Message
			}


			Start-Sleep -Seconds $CheckInterval
		}
	}
	catch
	{
		Write-Host ""
		Write-Host "Program error:"
		Write-Host $_.Exception.Message
	}
	finally
	{
		SaveAlarmState

		Write-Host ""
		Write-Host "Releasing Excel..."

		# ReleaseExcel

		Write-Host "Exit."
	}
}


############################################################
# メイン入口
############################################################

Main
