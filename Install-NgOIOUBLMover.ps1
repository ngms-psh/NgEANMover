
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
    [int]$RunInterval = 30,
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

    $LogFile = "$LogFolder\$LogFilePrefix$(get-date -Format 'dd-MM-yyyy_HHmmss').log"
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
        [parameter(Mandatory)][string]$ScriptPath,
        [parameter(Mandatory)][string]$ScriptParameters,
        [string]$IconLocation = "%SystemRoot%\System32\SHELL32.dll,45",
        [string]$WorkingDirectory = $InstallPath,
        [switch]$Force
    )
    $Arguments = "-ExecutionPolicy ByPass -WindowStyle Minimized -File `"$ScriptPath`"$ScriptParameters"
    try {
        $ShortcutPath = Join-Path -Path $ShortLocation -ChildPath "$ShortcutName.lnk"
        if (!(Test-Path -Path $ShortcutPath) -or $Force) {
            if(Test-Path -Path $ShortcutPath){
                Remove-Item -Path $ShortcutPath -Force
            }
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
        [parameter(Mandatory)][string]$ScriptPath,
        [parameter(Mandatory)][string]$ScriptParameters
    )

    $StartMenuFolderPath = "$($env:APPDATA)\Microsoft\Windows\Start Menu\Programs\$StartMenuFolderName"
    if (!(Test-Path $StartMenuFolderPath)) {
        New-Item -Path $StartMenuFolderPath -ItemType Directory | Out-Null
        write-NgLogMessage -Message "Created start menu folder: '$StartMenuFolderPath'" -Level Information
    }

    New-NgShortcut -ShortcutName $ShortcutName -ShortLocation $StartMenuFolderPath -ScriptPath $ScriptPath -ScriptParameters $ScriptParameters -Force
}

function Add-NgDesktopShortcut {
    param (
        [parameter(Mandatory)][string]$ScriptPath,
        [parameter(Mandatory)][string]$ScriptParameters,
        [string]$ShortcutName = "EAN Mover"
    )
    $DesktopPath = [Environment]::GetFolderPath("Desktop")
    New-NgShortcut -ShortcutName $ShortcutName -ShortLocation $DesktopPath -ScriptPath $ScriptPath -ScriptParameters $ScriptParameters -Force
}

function Add-NgScheduledTask {
    param(
        [parameter(Mandatory)][string]$TaskName,
        [parameter(Mandatory)][string]$ScriptParameters,
        [parameter(Mandatory)][string]$ScriptPath,
        [string]$TaskDescription,
        [int]$TaskId = 124563,
        [int]$TaskInterval = 30,
        [bool]$Disabled,
        [int]$TimeOut = 15
    )
    
    try {
        $TaskAction = New-ScheduledTaskAction -Execute "C:\Windows\System32\WindowsPowerShell\v1.0\Powershell.exe" -Argument "-ExecutionPolicy ByPass -WindowStyle Minimized -File `"$ScriptPath`"$ScriptParameters" -WorkingDirectory $InstallPath -Id $TaskId
        $TaskTrigger = New-ScheduledTaskTrigger -Once -at ((Get-Date).AddMinutes(2)) -RepetitionInterval (New-TimeSpan -Minutes $TaskInterval)
        if($Disabled -eq $true){
            $TaskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -Hidden -DontStopOnIdleEnd -ExecutionTimeLimit (New-TimeSpan -Minutes $TimeOut) -Disable
        }
        else{
            $TaskSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -Hidden -DontStopOnIdleEnd -ExecutionTimeLimit (New-TimeSpan -Minutes $TimeOut)
        }
        Register-ScheduledTask -TaskName $TaskName -Action $TaskAction -Trigger $TaskTrigger -Description $TaskDescription -Settings $TaskSettings -Force | Out-Null
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
        [System.Object]$RequiredFiles,
        [parameter(Mandatory)]
        [System.UriBuilder]$GitHubRawUrl,
        [parameter(Mandatory)]
        [System.UriBuilder]$GitHubRepoUrl
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
        try {
            $HTTP_ResponseRepo = $HTTP_RequestRepo.GetResponse()
        }
        catch {
            write-NgLogMessage -Message "Unable to connect to $GitHubRawUrl" -Level Error
            write-Error "Install-NgFiles: Unable to connect to $GitHubRawUrl $_"
            throw $_
            exit 1
        }
        
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
                write-NgLogMessage -Message "Unable to download $MissingFile to $InstallPath $_" -Level Error
                Write-Error "Install-NgFiles: Unable to download $MissingFile to $InstallPath $_"
                return $_
            }
        }
    }
}

$TaskName = "NgOIOUBLMover"
$TaskDescription = "Move OIOUBL/EAN files from the downloads folder to $AzureFileShare"

$GitHubRepoUrl = "https://github.com/ngms-psh/NgEANMover"
$GitHubRawUrl = "https://raw.githubusercontent.com/ngms-psh/NgEANMover/main"

$NgScript = "NgOIOUBLMover.ps1"
$NgScriptParameters = " -AzureFileShare `"$AzureFileShare`""
$NgInstaller = $MyInvocation.MyCommand.Name
$RequiredFiles = @($NgInstaller, $NgScript)

$LogPath = $env:temp
[string]$LogFilePrefix = "Install_" # Date will be appended to the prefix ex. Install_10-12-2024.log

try {
    # Set the log folder
    [string]$LogFolder = Join-Path -Path $LogPath -ChildPath $FolderName # Log files will be stored in the temp folder in a folder named NgOIOUBLMover
    write-NgLogMessage -Message "Starting installation of NgOIOUBLMover" -Level Information
    write-NgLogMessage -Message "Azure File Share: $AzureFileShare" -Level Information
    write-NgLogMessage -Message "Install location: $InstallLocation" -Level Information
    write-NgLogMessage -Message "Folder name: $FolderName" -Level Information
    write-NgLogMessage -Message "Force: $Force" -Level Information
    write-NgLogMessage -Message "Run interval: $RunInterval" -Level Information

    $InstallPath = Join-Path -Path $InstallLocation -ChildPath $FolderName
    $NgScriptPath = Join-Path -Path $InstallPath -ChildPath $NgScript

    write-NgLogMessage -Message "Install path: $InstallPath" -Level Information
    write-NgLogMessage -Message "NgScript path: $NgScriptPath" -Level Information


    write-NgLogMessage -Message "Log folder: $LogFolder" -Level Information

    try {
        Install-NgFiles -InstallPath $InstallPath -RequiredFiles $RequiredFiles -GitHubRawUrl $GitHubRawUrl -GitHubRepoUrl $GitHubRepoUrl

    }
    catch {
        Write-Error $_
        exit 1
    }




    if (-not $DisableStartMenuShortcut) {
        Add-NgStartMenuShortcut -ScriptPath $NgScriptPath -ScriptParameters $NgScriptParameters
    }

    if (-not $DisableDesktopShortcut) {
        Add-NgDesktopShortcut -ScriptPath $NgScriptPath -ScriptParameters $NgScriptParameters
    }

    Add-NgScheduledTask -TaskName $TaskName -ScriptPath $NgScriptPath -ScriptParameters $NgScriptParameters -TaskDescription $TaskDescription -TaskInterval $RunInterval -TaskId 124563 -Disabled $DisableScheduledTask -TimeOut 15
    write-NgLogMessage -Message "Installation of NgOIOUBLMover completed" -Level Information
    [System.Windows.Forms.MessageBox]::Show($THIS, "Installation of NgOIOUBLMover Completed",'OIOUBL Mover','OK','Information')
}
catch {
    write-NgLogMessage -Message "Unable to install NgOIOUBLMover $_" -Level Error
    [System.Windows.Forms.MessageBox]::Show($THIS, "Installation of NgOIOUBLMover Failed",'OIOUBL Mover','OK','error')
    Write-Error $_
    exit 1
}

