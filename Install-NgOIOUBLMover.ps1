
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
    [Parameter(Mandatory = $true,HelpMessage="URL to the Azure File Share or the drive letter of the mapped drive")]
    [string]$AzureFileShare,
    [switch]$DisableScheduledTask,
    [switch]$DisableStartMenuShortcut,
    [switch]$DisableDesktopShortcut,
    [int]$TaskInterval = 30,
    [string]$InstallLocation = $env:USERPROFILE,
    [string]$FolderName = "NgOIOUBLMover",
    [switch]$Force
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

function New-NgShortcut {
    param (
        [parameter(Mandatory)][string]$ShortcutName,
        [parameter(Mandatory)][string]$ShortLocation,
        [string]$TargetPath = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
        [parameter(Mandatory)][string]$ActionCommand,
        [string]$IconLocation = "%SystemRoot%\System32\SHELL32.dll,45",
        [string]$WorkingDirectory = $InstallPath,
        [switch]$Force
    )
    $Arguments = "-ExecutionPolicy ByPass -WindowStyle Minimized -File `"$ActionCommand`""
    try {
        $ShortcutPath = Join-Path -Path $ShortLocation -ChildPath "$ShortcutName.lnk"
        if (!(Test-Path -Path $ShortcutPath) -or $Force) {
            $shell = New-Object -ComObject WScript.Shell
            $shortcut = $shell.CreateShortcut($ShortcutPath)
            $shortcut.TargetPath = $TargetPath
            $shortcut.Arguments = $Arguments
            $shortcut.IconLocation = $IconLocation
            $shortcut.WorkingDirectory = $WorkingDirectory
            $shortcut.Save()
            write-NgLogMessage -Message "Created shortcut: '$ShortcutPath'" -Level Information
        }
    }
    catch {
        write-NgLogMessage -Message "Unable to create shortcut: '$ShortcutPath'" -Level Error
        Write-Error $_
    }
}

function Add-NgStartMenuShortcut {
    param(
        [string]$ShortcutName = "EAN Mover",
        [string]$StartMenuFolderName = "EAN Mover",
        [parameter(Mandatory)][string]$ActionCommand
    )

    $StartMenuFolderPath = "$($env:APPDATA)\Microsoft\Windows\Start Menu\Programs\$StartMenuFolderName"
    if (!(Test-Path $StartMenuFolderPath)) {
        New-Item -Path $StartMenuFolderPath -ItemType Directory | Out-Null
        write-NgLogMessage -Message "Created start menu folder: '$StartMenuFolderPath'" -Level Information
    }

    New-NgShortcut -ShortcutName $ShortcutName -ShortLocation $StartMenuFolderPath -ActionCommand $ActionCommand
}

function Add-NgDesktopShortcut {
    param (
        [parameter(Mandatory)][string]$ActionCommand,
        [string]$ShortcutName = "EAN Mover"
    )
    $DesktopPath = [Environment]::GetFolderPath("Desktop")
    New-NgShortcut -ShortcutName $ShortcutName -ShortLocation $DesktopPath -ActionCommand $ActionCommand
}

function Add-NgScheduledTask {
    param(
        [parameter(Mandatory)][string]$TaskName,
        [parameter(Mandatory)][string]$ActionCommand,
        [string]$TaskDescription,
        [int]$TaskId = 124563,
        [int]$TaskInterval = 30,
        [parameter(Mandatory)][System.Object]$TaskAction,
        [parameter(Mandatory)][System.Object]$TaskTrigger,
        [parameter(Mandatory)][System.Object]$TaskSettings,
        [string]$ScriptParameters,
        [switch]$Disabled,
        [int]$TimeOut = 15
    )
    if ($ScriptParameters){$ActionCommand = $ActionCommand + " " + $ScriptParameters}
    try {
        (Get-Date -Format "HH:mm").AddMinutes(2)
        $TaskAction = New-ScheduledTaskAction -Execute "C:\Windows\System32\WindowsPowerShell\v1.0\Powershell.exe" -Argument "-ExecutionPolicy ByPass -WindowStyle Minimized -File `"$ActionCommand`"" -WorkingDirectory $InstallPath -Id $TaskId
        $TaskTrigger = New-ScheduledTaskTrigger -Once -at ((Get-Date).AddMinutes(2)) -RepetitionInterval (New-TimeSpan -Minutes $TaskInterval)
        if($Disabled){
            $TaskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -Hidden -DontStopOnIdleEnd -ExecutionTimeLimit (New-TimeSpan -Minutes $TimeOut) -Disable
        }
        else{
            $TaskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -Hidden -DontStopOnIdleEnd -ExecutionTimeLimit (New-TimeSpan -Minutes $TimeOut)
        }
        Register-ScheduledTask -TaskName $TaskName -Action $TaskAction -Trigger $TaskTrigger -Description $TaskDescription -Settings $TaskSettings -Force
        write-NgLogMessage -Message "Created scheduled task: '$TaskName'" -Level Information
    }
    catch {
        write-NgLogMessage -Message "Unable to create scheduled task: '$TaskName'" -Level Error
        Write-Error $_
    }

}

function Install-NgFiles {
    param(
        [string]$InstallPath,
        [parameter(Mandatory)]
        [System.Object]$RequiredFiles
    )

    $HTTP_RequestRepo = [System.Net.WebRequest]::Create($GitHubRepoUrl.Uri)

    

    if (!(Test-Path -Path $InstallPath -PathType Container)) {
        New-Item -Path $InstallPath -ItemType Directory | Out-Null
    }

    $MissingFiles = $RequiredFiles | Where-Object { -not (Test-Path -Path (Join-Path -Path $InstallPath -ChildPath $_)) }

    if ((!$MissingFiles) -and (!$Force)) {
        Write-NgLogMessage -Message "All requiredfiles for NgOIOUBLMover is already installed in '$InstallPath'" -Level Warning
    }
    else{
        $HTTP_ResponseRepo = $HTTP_RequestRepo.GetResponse()
        if ([int]$HTTP_ResponseRepo.StatusCode -ne 200){
            write-NgLogMessage -Message "Unable to download files from $GitHubRepoUrl" -Level Error
            exit 1
        }
        foreach ($MissingFile in $MissingFiles) {
            try {
                Invoke-WebRequest -Uri "$($GitHubRawUrl.Uri)/$MissingFile" -OutFile (Join-Path -Path $InstallPath -ChildPath $MissingFile)
                write-NgLogMessage -Message "Downloaded $MissingFile to $InstallPath" -Level Information
            }
            catch {
                write-NgLogMessage -Message "Unable to download $MissingFile to $InstallPath" -Level Error
                Write-Error "Install-NgFiles: Unable to download $MissingFile to $InstallPath"
                return $_.Exception.Message
            }
        }
    }
}

$TaskName = "NgOIOUBLMover"
$TaskDescription = "Move OIOUBL/EAN files from the downloads folder to $SourceFolder"

$NgScript = "NgOIOUBLMover.ps1"
$NgInstaller = $MyInvocation.MyCommand.Name
$RequiredFiles = @($NgInstaller, $NgScript)



$InstallPath = Join-Path -Path $InstallLocation -ChildPath $FolderName
$NgScriptPath = Join-Path -Path $InstallPath -ChildPath $NgScript
$ActionCommand = "$NgScriptPath -AzureFileShare `"$AzureFileShare`""

# Set the log folder and log file prefix
[string]$LogFolder = Join-Path -Path $env:temp -ChildPath $FolderName # Log files will be stored in the temp folder in a folder named NgOIOUBLMover
[string]$LogFilePrefix = "Install_" # Date will be appended to the prefix ex. Install_10-12-2024.log



[System.UriBuilder]$GitHubRepoUrl = "https://github.com/ngms-psh/NgEANMover"
$GitHubRawUrl = [uri]::new("https://raw.githubusercontent.com/ngms-psh/NgEANMover/main")

Install-NgFiles -InstallPath $InstallPath -RequiredFiles $RequiredFiles -ErrorAction Stop



if (-not $DisableStartMenuShortcut) {
    Add-NgStartMenuShortcut -ActionCommand $ActionCommand
}

if (-not $DisableDesktopShortcut) {
    Add-NgDesktopShortcut -ActionCommand $ActionCommand
}

Add-NgScheduledTask -TaskName $TaskName -ActionCommand $ActionCommand -TaskDescription $TaskDescription -TaskInterval $TaskInterval -TaskId 124563 -Disabled $DisableScheduledTask
