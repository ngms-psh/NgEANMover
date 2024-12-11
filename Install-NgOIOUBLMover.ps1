
<#PSScriptInfo

.VERSION 1.0

.GUID b00572a7-8e47-4c57-9be2-b0ccad3fa98f

.AUTHOR Phillip Schjeldal Hansen | NgMS Consult ApS

.COMPANYNAME NgMS Consult ApS

.COPYRIGHT (c) 2024 - Phillip Schjeldal Hansen | NgMS Consult ApS. All rights reserved.

.TAGS

.LICENSEURI

.PROJECTURI

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 081c47a1-20d0-47ab-9d30-2dbac7107499

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


.PRIVATEDATA

#> 



<# 

.DESCRIPTION 
 Install script NgOIOBULMover from nuget, creates shortcuts and optional scheaduled task 

#> 
#requires -PSEdition Desktop
[CmdletBinding()]
Param (
    [switch]$DisableScheduledTask,
    [switch]$DisableStartMenuShortcut,
    [switch]$DisableDesktopShortcut,
    [string]$InstallLocation = $env:USERPROFILE,
    [string]$FolderName = "NgOIOUBLMover"
)

function Write-NgLogMessage {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 1)]
        $Message,

        [Parameter(Mandatory = $true, Position = 0)]
        [validateSet('Error', 'Warning', 'Information')]
        [string]$Level
    )
    $ParameterList = (Get-Command -Name $MyInvocation.MyCommand).Parameters
    $MaxLength = ($ParameterList["Level"].Attributes.ValidValues | Sort-Object { $_.Length } -Descending | Select-Object -First 1).Length

    # Pad the message to the maximum length
    $LevelPadded = $Level.PadRight($MaxLength)

    $LogFile = "$LogFolder\$LogFilePrefix$(get-date -Format 'dd-MM-yyyy').log"
    If (!(Test-Path $LogFolder)) {New-Item -Path $LogFolder -Type Directory -Force | Out-Null}
    foreach ($M in $Message) {
        $Date = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
        $FullM = "$Date | $LevelPadded - $M"
        Add-Content -Path $LogFile -Value $FullM -Force
    }
}

function Add-NgStartMenuShortcut {
    $FolderName = "NgMS"
    $FolderPath = "$($env:APPDATA)\Microsoft\Windows\Start Menu\Programs\$FolderName"
    if(!(Test-Path -Path $FolderPath)){
        New-Item -Path $FolderPath -ItemType Directory | Out-Null
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut("$FolderPath\Ng EAN mover.lnk")
        $shortcut.TargetPath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
        $shortcut.Arguments = "-ExecutionPolicy ByPass -WindowStyle Minimized -File `"$PSCommandPath`""
        $shortcut.IconLocation = "%SystemRoot%\System32\SHELL32.dll,45"
        $shortcut.Save()
    }
}

function Add-NgScheduledTask {
    $TaskName = "Ng EAN Mover"
    $TaskDescription = "Move OIOUBL/EAN files from the downloads folder to $SourceFolder"
    $TaskAction = New-ScheduledTaskAction -Execute "C:\Windows\System32\WindowsPowerShell\v1.0\Powershell.exe" -Argument "-ExecutionPolicy ByPass -WindowStyle Minimized -File `"$PSCommandPath`""
    $TaskTrigger = New-ScheduledTaskTrigger -AtLogOn
    $TaskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
    Register-ScheduledTask -TaskName $TaskName -Action $TaskAction -Trigger $TaskTrigger -Description $TaskDescription -Settings $TaskSettings
}

# Set the log folder and log file prefix
[string]$LogFolder = Join-Path -Path $env:temp -ChildPath $FolderName # Log files will be stored in the temp folder in a folder named NgOIOUBLMover
[string]$LogFilePrefix = "Install_" # Date will be appended to the prefix ex. Install_10-12-2024.log

$InstallPath = Join-Path -Path $InstallLocation -ChildPath $FolderName
$RequiredFiles = @("Install-NgOIOUBLMover.ps1", "NgOIOUBLMover.ps1")

if (!(Test-Path -Path $InstallPath -PathType Container)) {
    New-Item -Path $InstallPath -ItemType Directory | Out-Null
}

$MissingFiles = $RequiredFiles | Where-Object { -not (Test-Path -Path (Join-Path -Path $InstallPath -ChildPath $_)) }


