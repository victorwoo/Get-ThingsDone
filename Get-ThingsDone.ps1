#========================================================================
# Created with: SAPIEN Technologies, Inc., PowerShell Studio 2012 v3.1.12
# Created on:   2013-3-7 16:21
# Created by:   Victor.Woo
# Organization: www.vichamp.com
# Filename:  Get-ThingsDone.ps1   
#========================================================================

function Check-Enviroment
{
  $gtdPath = "HKCU:\Software\Vichamp\GTD"
  if ((Get-ItemProperty $gtdPath -ErrorAction SilentlyContinue).AutoStart -eq "False")
  {
    return
  }

  $runPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
  $run = Get-ItemProperty $runPath
  if ($run.GTD -eq $null)
  {
    $title = '自动执行请求'
    $message = '在当前 Windows 用户登录时自动运行此脚本，可以自动帮助您整理、规划当日的工作内容。如果您选择了“是”，但将来不希望自动启动，请执行 uninstall.cmd。是否在当前用户登录时自动执行脚本？'
    $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Windows 用户登录时自动运行此脚本。"
    $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No","Windows 用户登录时不运行此脚本，并且不再提示。"
    $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes,$no)
    $result = $Host.UI.PromptForChoice($title,$message,$options,0)
    switch ($result)
    {
      0 {
        Set-ItemProperty -Path $runPath -Name GTD -Value $gtdCmd
      }
      1 {
        md $gtdPath -Force
        Set-ItemProperty -Path $gtdPath -Name AutoStart -Value "False"
      }
    }
  }
}

function TryCreate-Directory ([Parameter(Mandatory = $True)] [string]$dirName)
{
  $private:dir = Join-Path $baseDir $dirName
  if (-not (Test-Path $dir))
  {
    Write-Output "$dir 不存在，正在创建。"
    mkdir $dir | Out-Null
  }
}

function TryCreate-Directories ()
{
  Write-Output "正在检查目录完整性"
  $dirNames |
  % {
    TryCreate-Directory $_
  }
}

function Remove-Directories ()
{
  $dirNames |
  % {
    $private:dir = Join-Path $baseDir $_
    if (Test-Path $dir)
    {
      Write-Warning "正在移除$dir"
      rm $dir -Recurse
    }
  }
}

function MoveTo-WithRenamming (
  [Parameter(Mandatory = $True)] [System.IO.FileSystemInfo]$item,
  [Parameter(Mandatory = $True)] [string]$targetDir)
{
  function Get-NextFilePath ([string]$dir,[System.IO.FileInfo]$fileInfo)
  {
    $Private:targetFilePath = Join-Path $dir $fileInfo.Name
    if (Test-Path $Private:targetFilePath)
    {
      $Private:index = 1
      do {
        $Private:targetFilePath = Join-Path $dir "$($fileInfo.BaseName) ($index)$($fileInfo.Extension)"
        $Private:index++
      }
      while (Test-Path $Private:targetFilePath)
    }
    return [System.IO.FileInfo]$Private:targetFilePath
  }

  function Get-NextDirectoryPath ([string]$dir,[System.IO.DirectoryInfo]$directoryInfo)
  {
    $Private:targetDirectoryPath = Join-Path $dir $directoryInfo.Name
    if (Test-Path $Private:targetDirectoryPath)
    {
      $Private:index = 1
      do {
        $Private:targetDirectoryPath = Join-Path $dir "$($directoryInfo.Name) ($index)"
        $Private:index++
      }
      while (Test-Path $Private:targetDirectoryPath)
    }
    return [System.IO.DirectoryInfo]$Private:targetDirectoryPath
  }

  Write-Output "正在移动 $item 至 $targetDir 目录"
  if ($item -is [System.IO.FileInfo])
  {
    [System.IO.FileInfo]$item = [System.IO.FileInfo]$item
    $Private:targetFilePath = Join-Path $targetDir $item.Name
    if (Test-Path $Private:targetFilePath)
    {
      $targetFileInfo = [System.IO.FileInfo]$Private:targetFilePath
      $Private:targetFilePath = Get-NextFilePath $targetDir $item
      if ($item.LastWriteTime -eq $targetFileInfo.LastWriteTime -and $item.Length -eq $targetFileInfo.Length)
      {
        Write-Warning "源文件 $item.FullName 与目标文件 $targetFileInfo.FullName 相同，删除源文件"
        Remove-Item $item.FullName
      }
      else
      {
        Write-Warning "目标文件已存在，自动改名为$($Private:targetFilePath.Name)"
        Move-Item $item.FullName $Private:targetFilePath | Out-Null
      }
    }
  } elseif ($item -is [System.IO.DirectoryInfo])
  {
    [System.IO.DirectoryInfo]$item = [System.IO.DirectoryInfo]$item
    $Private:targetDirectoryPath = Join-Path $targetDir $item.Name
    if (Test-Path $Private:targetDirectoryPath)
    {
      $Private:targetDirectoryPath = Get-NextDirectoryPath $targetDir $item
      Write-Warning "目标文件夹已存在，自动改名为$($Private:targetDirectoryPath.Name)"
    }
    Move-Item $item.FullName $Private:targetDirectoryPath | Out-Null
  }
}

function Process-IsolatedItems
{
  Write-Output "正在将游离内容移至 [STUFF] 目录"
  Get-ChildItem $baseDir -Exclude ($dirNames + $reservedDirs + $reservedFiles) |
  % {
    MoveTo-WithRenamming $_ $stuffDir
  }
}

function Process-TomorrowDir
{
  Write-Output "正在处理 [TOMORROW] 目录"
  Get-ChildItem $tomorrowDir |
  % {
    MoveTo-WithRenamming $_ $todayDir
  }
}

function Process-CalendarDir
{
  Write-Output "正在处理 [CALENDAR] 目录"
  Get-ChildItem $calendarDir -File |
  % {
    MoveTo-WithRenamming $_ $stuffDir
  }

  Get-ChildItem $calendarDir -Directory |
  % {
    $regex = [regex]'(?m)^(?<year>19|20\d{2})[-_.](?<month>\d{1,2})[-_.](?<day>\d{1,2})$'
    $match = $regex.Match($_.Name)
    if ($match.Success)
    {
      $Private:year = $regex.Match($_.Name).Groups['year'].Value;
      $Private:month = $regex.Match($_.Name).Groups['month'].Value;
      $Private:day = $regex.Match($_.Name).Groups['day'].Value;
      $Private:date = New-Object System.DateTime $Private:year,$Private:month,$Private:day
      $now = (Get-Date)
      $today = $now.Subtract($now.TimeOfDay)
      if ($date -lt $today)
      {
        Write-Output "移动过期任务 $($_.Name) 到 [STUFF] 目录"
        MoveTo-WithRenamming $_ $stuffDir
      }
      elseif ($date -eq $today)
      {
        Write-Output "移动今日任务 $($_.Name) 到 [TODAY] 目录"
        MoveTo-WithRenamming $_ $todayDir
      }
      elseif ($date -eq $today.AddDays(1))
      {
        Write-Output "移动明日任务 $($_.Name) 到 [TOMORROW] 目录"
        MoveTo-WithRenamming $_ $tomorrowDir
      }
    }
    else
    {
      Write-Output "[CALENDAR] 目录下，$($_.Name) 名字不符合规范，将移动至 [STUFF] 目录"
      MoveTo-WithRenamming $_ $stuffDir
    }
  }
}

function Process-ArchiveDir
{
  Write-Output "正在检查 [ARCHIVE] 目录"

  # 创建本月目录
  $nowString = "{0:yyyy.MM}" -f (Get-Date)
  $thisMonthDir = Join-Path $archiveDir $nowString
  if (-not (Test-Path $thisMonthDir))
  {
    Write-Output "正在创建本月目录"
    md $thisMonthDir
  }

  # 移除除本月之外的空目录
  Get-ChildItem $archiveDir -Exclude $nowString -Recurse |
  Where { $_.PSIsContainer -and @( Get-ChildItem -LiteralPath $_.FullName -Recurse | Where { !$_.PSIsContainer }).Length -eq 0 } |
  % {
    Write-Output "正在删除空目录$($_.FullName)"
    Remove-Item -Recurse
  }

  # 移动所有文件到 本月存档 目录
  Get-ChildItem $archiveDir -File |
  % {
    $createTime = $nowString = "{0:yyyy.MM}" -f $_.CreationTime
    Write-Output "移动 [ARCHIVE] 目录下，$($_.Name) 游离文件至 $createTime 存档目录"
    MoveTo-WithRenamming $_ $thisMonthDir
  }

  # 检查目录命名是否符合规范。
  Get-ChildItem $archiveDir -Directory |
  % {
    $regex = [regex]'(?m)^(?<year>19|20\d{2})[-_.](?<month>\d{1,2})$'
    $match = $regex.Match($_.Name)
    if ($match.Success)
    {
      $year = $regex.Match($_.Name).Groups['year'].Value;
      $month = $regex.Match($_.Name).Groups['month'].Value;
      $date = New-Object System.DateTime $year,$month,1
      if ($date -gt (Get-Date))
      {
        Write-Output "[ARCHIVE] 目录下，$($_.Name) 名字不符合规范（存档日期超出当前时间），将移动至 [STUFF] 目录"
        MoveTo-WithRenamming $_ $stuffDir
      }
      else
      {
        $formattedDate = "{0:yyyy.MM}" -f $date
        if ($_.Name -ne $formattedDate)
        {
          $targetDirectory = [System.IO.DirectoryInfo](Join-Path $_.Parent.FullName $formattedDate)
          Write-Warning "将 [ARCHIVE] 下的目录名 $($_.Name) 处理为规范格式 $($targetDirectory.Name)"
          Move-Item $_.FullName $targetDirectory.FullName
        }
      }
    } else
    {
      $createTime = $nowString = "{0:yyyy.MM}" -f $_.CreationTime
      Write-Output "移动 [ARCHIVE] 目录下，$($_.Name) 游离文件夹至 $createTime 存档目录"
      MoveTo-WithRenamming $_ $thisMonthDir
    }
  }
}

function Explore-Dirs
{
  if ((Get-ChildItem $stuffDir) -ne $null)
  {
    explorer $stuffDir
  }

  if ((Get-ChildItem $todayDir) -ne $null)
  {
    explorer $todayDir
  }
}

$STUFF = "1.STUFF"
$TODAY = "2.TODAY"
$TOMORROW = "3.TOMORROW"
$UPCOMING = "4.UPCOMING"
$CALENDAR = "5.CALENDAR"
$SOMEDAY = "6.SOMEDAY"
$ARCHIVE = "7.ARCHIVE"

$dirNames = $STUFF,$TODAY,$TOMORROW,$UPCOMING,$CALENDAR,$SOMEDAY,$ARCHIVE
$reservedDirs = ".git","_gsdata_"
$reservedFiles = ".gitignore","Get-ThingsDone.ps1","README*.md","gtd_logo.png","LICENSE.md","GTD.cmd","uninstall.cmd"

$baseDir = Split-Path $MyInvocation.MyCommand.Path
$stuffDir = Join-Path $baseDir $STUFF
$todayDir = Join-Path $baseDir $TODAY
$tomorrowDir = Join-Path $baseDir $TOMORROW
$calendarDir = Join-Path $baseDir $CALENDAR
$archiveDir = Join-Path $baseDir $ARCHIVE
$gtdCmd = Join-Path $baseDir "GTD.cmd"

Get-Date | Write-Output

Check-Enviroment
TryCreate-Directories

Process-IsolatedItems
Process-TomorrowDir
Process-CalendarDir
Process-ArchiveDir

Explore-Dirs

######################### 开发临时用（在 ISE 中选中并按 F8 执行） #########################
{
  return
  # 创建游离内容。
  $null | Set-Content (Join-Path $baseDir "to.del.file.txt")
  md (Join-Path $baseDir "to.del.dir") | Out-Null
}

{
  return
  # 对代码排版。
  Import-Module D:\Dropbox\script\DTW.PS.PrettyPrinterV1\DTW.PS.PrettyPrinterV1.psd1
  Edit-DTWCleanScript D:\Dropbox\vichamp\GTD\Get-ThingsDone.ps1
}

{
  return
  # 移除所有目录
  Remove-Directories
}
