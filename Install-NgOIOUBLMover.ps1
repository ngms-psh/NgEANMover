
<#PSScriptInfo

.VERSION 1.1

.GUID b00572a7-8e47-4c57-9be2-b0ccad3fa98f

.AUTHOR Phillip Schjeldal Hansen | NgMS Consult ApS

.COMPANYNAME NgMS Consult ApS

.COPYRIGHT (c) 2024 - Phillip Schjeldal Hansen | NgMS Consult ApS. All rights reserved.

.TAGS

.LICENSEURI

.PROJECTURI

.ICONURI

.EXTERNALMODULEDEPENDENCIES ps2exe

.REQUIREDSCRIPTS 081c47a1-20d0-47ab-9d30-2dbac7107499

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


.PRIVATEDATA

#> 



<# 

.DESCRIPTION 
 Install script NgOIOBULMover from Github, creates shortcuts and optional scheaduled task 

#> 
#requires -PSEdition Desktop
[CmdletBinding()]
Param (
    [validateSet('exe', 'ps1')]
    [string]$Type = "exe",
    [string]$AzureFileShare,
    [switch]$DisableScheduledTask,
    [switch]$DisableStartMenuShortcut,
    [switch]$DisableDesktopShortcut,
    [int]$RunInterval = 30,
    [string]$InstallLocation = $env:USERPROFILE,
    [string]$FolderName = "NgOIOUBLMover",
    [switch]$Force
)

Add-Type -AssemblyName PresentationCore,PresentationFramework,System.Windows.Forms

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

    $LogFile = "$LogFolder\$LogFilePrefix$StartTime.log"
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
        [parameter(Mandatory)][string]$ScriptPath,
        [parameter(Mandatory)][string]$ScriptParameters,
        [string]$IconLocation = (Join-Path -Path $InstallPath -ChildPath $IconFile),
        [string]$WorkingDirectory = $InstallPath,
        [switch]$Force
    )
    if ($Type -eq "exe") {
        $TargetPath = $NgScriptPath -replace '.ps1','.exe'
        $Argument = $ScriptParameters
    }
    else{
        $TargetPath = "C:\Windows\System32\WindowsPowerShell\v1.0\Powershell.exe"
        $Argument = "-ExecutionPolicy ByPass -WindowStyle Minimized -File `"$ScriptPath`"$ScriptParameters"
    }

    try {
        $ShortcutPath = Join-Path -Path $ShortLocation -ChildPath "$ShortcutName.lnk"
        if (!(Test-Path -Path $ShortcutPath) -or $Force) {
            if(Test-Path -Path $ShortcutPath){
                Remove-Item -Path $ShortcutPath -Force
            }
            $shell = New-Object -ComObject WScript.Shell
            $shortcut = $shell.CreateShortcut($ShortcutPath)
            $shortcut.TargetPath = $TargetPath
            $shortcut.Arguments = $Argument
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

    if ($Type -eq "exe") {
        $Execute = $NgScriptPath -replace '.ps1','.exe'
        $Argument = $ScriptParameters

    }
    else{
        $Execute = "C:\Windows\System32\WindowsPowerShell\v1.0\Powershell.exe"
        $Argument = "-ExecutionPolicy ByPass -WindowStyle Minimized -File `"$ScriptPath`"$ScriptParameters"
    }

    write-NgLogMessage -Message "Execute: $Execute" -Level Information
    write-NgLogMessage -Message "Argument: $Argument" -Level Information

    try {
        $TaskAction = New-ScheduledTaskAction -Execute $Execute -Argument $Argument -WorkingDirectory $InstallPath -Id $TaskId
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
        [System.UriBuilder]$GitHubRepoUrl,
        [parameter(Mandatory)]
        [string]$Compile,
        [string]$Icon
    )
    $HTTP_RequestRepo = [System.Net.WebRequest]::Create($GitHubRepoUrl.Uri)

    

    if (!(Test-Path -Path $InstallPath -PathType Container)) {
        New-Item -Path $InstallPath -ItemType Directory | Out-Null
    }

    if ($Force){
        $MissingFiles = $RequiredFiles
    }
    else{
        $MissingFiles = $RequiredFiles | Where-Object { -not (Test-Path -Path (Join-Path -Path $InstallPath -ChildPath $_)) }
    }

    if ((!$MissingFiles) -and (!$Force)) {
        Write-NgLogMessage -Message "All requiredfiles for NgOIOUBLMover is already installed in '$InstallPath'" -Level Warning
    }
    else{
        try {
            write-host "Checking connection to '$($GitHubRepoUrl.Uri)'" -ForegroundColor Yellow
            $HTTP_ResponseRepo = $HTTP_RequestRepo.GetResponse()
            write-host "Success: Connection was successful" -ForegroundColor Green
        }
        catch {
            write-host "Error: Connection failed`nUnable to connect to '$($GitHubRepoUrl.Uri)'`nTerminating installation" -ForegroundColor Red
            write-NgLogMessage -Message "Unable to connect to '$($GitHubRepoUrl.Uri)'" -Level Error
            write-Error "Install-NgFiles: Unable to connect to '$($GitHubRepoUrl.Uri)' $_"
            throw $_
            exit 1
        }
        
        if ([int]$HTTP_ResponseRepo.StatusCode -ne 200){
            write-host "Error: Connection failed`nUnable to connect to '$($GitHubRepoUrl.Uri)'`nTerminating installation" -ForegroundColor Red
            write-NgLogMessage -Message "Unable to download files from $GitHubRepoUrl" -Level Error
            exit 1
        }
        if ($Force){
            try {
                if (Test-Path $InstallPath){Remove-Item $InstallPath -Recurse -Force}
                
                Write-host "Downloading source files from '$($GitHubRawUrl.Uri)'" -ForegroundColor yellow

                Invoke-WebRequest -Uri $GitHubRawUrl.Uri -OutFile (Join-Path -Path $InstallLocation -ChildPath "NgEANMover.Zip")
                Expand-Archive (Join-Path -Path $InstallLocation -ChildPath "NgEANMover.Zip") $InstallLocation
                Move-Item (Join-Path -Path $InstallLocation -ChildPath "NgEANMover-main") $InstallPath -Force
                Remove-Item (Join-Path -Path $InstallLocation -ChildPath "NgEANMover.Zip")
                
                write-host "Success: Downloaded and extracted source files to '$($InstallPath)'" -ForegroundColor Green
                write-NgLogMessage -Message "Downloaded $($GitHubRawUrl.Uri) to $(Join-Path -Path $InstallPath -ChildPath 'NgEANMover.Zip')" -Level Information
            }
            catch {
                write-host "Error: Failed to download source files`nTerminating installation" -ForegroundColor Red
                write-NgLogMessage -Message "Unable to download $($GitHubRawUrl.Uri) to $InstallPath $_" -Level Error
                Write-Error "Install-NgFiles: Unable to download $($GitHubRawUrl.Uri) to $InstallPath $_"
                return $_
            }
        }

        else {
            foreach ($MissingFile in $MissingFiles) {
                try {
                    Write-host "Downloading source files from '$($GitHubRawUrl.Uri)'" -ForegroundColor yellow
                    Invoke-WebRequest -Uri "$($GitHubRawUrl.Uri)/$MissingFile" -OutFile (Join-Path -Path $InstallPath -ChildPath $MissingFile)
                    write-host "Success: Downloaded source files to '$($InstallPath)'" -ForegroundColor Green
                    write-NgLogMessage -Message "Downloaded $MissingFile to $InstallPath" -Level Information
                }
                catch {
                    write-host "Error: Failed to download source files`nTerminating installation" -ForegroundColor Red
                    write-NgLogMessage -Message "Unable to download $MissingFile to $InstallPath $_" -Level Error
                    Write-Error "Install-NgFiles: Unable to download $MissingFile to $InstallPath $_"
                    return $_
                }
            }
        }
        
        if ($Type -eq "exe") {
            try {
                write-host "Installing ps2exe module" -ForegroundColor Yellow
                Set-PSRepository "PSGallery" -InstallationPolicy Trusted
                Install-Module ps2exe -Scope CurrentUser -Force -AllowClobber
                Import-Module ps2exe -Force
                write-host "Success: installed ps2exe" -ForegroundColor Green
                write-NgLogMessage -Message "Installed ps2exe module" -Level Information
            }
            catch {
                write-host "Error: Unable to install ps2exe module" -ForegroundColor Red
                write-NgLogMessage -Message "Unable to install ps2exe module $_" -Level Error
                Write-Error "Install-NgFiles: Unable to install ps2exe module $_"
                return $_
            }
    
            try {
                write-host "ps2exe will now complice source files to EXE" -ForegroundColor Yellow
                Invoke-ps2exe -inputFile $Compile -outputFile "$($Compile -replace '.ps1','.exe')" -title "EAN Mover" -company "NgMS Consult ApS" -version "2.0" -product "NgOIOUBLMover" -noConsole -copyright "Copyright (c) 2024 - Phillip Schjeldal Hansen | NgMS Consult ApS. All rights reserved." -longPaths -iconFile (Join-Path -Path $InstallPath -ChildPath $Icon) -configFile -verbose
                Write-Host "Success: Compiled NgOIOUBLMover.exe" -ForegroundColor Green
                write-NgLogMessage -Message "Created executable NgOIOUBLMover.exe" -Level Information
            }
            catch {
                write-host "Error: Failed to compile NgOIOUBLMover.exe`nTerminating installation" -ForegroundColor Red
                write-NgLogMessage -Message "Unable to create executable $_" -Level Error
                Write-Error "Install-NgFiles: Unable to create executable $_"
                throw $_
            }
        }
    }
}
Write-Host "##############################################" 
Write-Host "Starting installation of EAN Mover" -ForegroundColor Green
Write-Host "COPYRIGHT (c) 2024 - Phillip Schjeldal Hansen | NgMS Consult ApS. All rights reserved" -ForegroundColor Green
Write-Host "##############################################" 
Write-Host ""


$TaskName = "NgOIOUBLMover"
$TaskDescription = "Move OIOUBL/EAN files from the downloads folder to $AzureFileShare"

$GitHubRepoUrl = "https://github.com/ngms-psh/NgEANMover"
$GitHubRawUrl = "https://github.com/ngms-psh/NgEANMover/archive/refs/heads/main.zip"

$NgScript = "NgOIOUBLMover.ps1"
$NgScriptParameters = " -AzureFileShare `"$AzureFileShare`" -Archive"
$NgInstaller = $MyInvocation.MyCommand.Name
$IconFile = "Icon.ico"
$RequiredFiles = @($NgInstaller, $NgScript, $IconFile)

$LogPath = $env:temp
[string]$LogFilePrefix = "Install_" # Date will be appended to the prefix ex. Install_10-12-2024.log

try {

    Write-Host "Checking if EAN Mover is currently running, if already installed" -ForegroundColor Yellow
    #check if NgOIOUBLMover is already running
    if ((Get-Process -Name "NgOIOUBLMover" -ErrorAction SilentlyContinue).count -gt 0) {
        Write-Host "EAN Mover already running`nPlease wait for it to complete and start the installation again`nTerminating installation" -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show($THIS, "EAN Mover already running`nPlease wait for it to complete and start the installation again",'EAN Mover','OK','error','Button1','ServiceNotification')
        exit "Process already running"
    }
    write-host "Success: EAN Mover is not running`nInstallation will continue" -ForegroundColor Green

    # Set the log folder
    [string]$LogFolder = Join-Path -Path $LogPath -ChildPath $FolderName # Log files will be stored in the temp folder in a folder named NgOIOUBLMover
    $StartTime = Get-Date -Format "dd-MM-yyyy_HHmmss"
    write-NgLogMessage -Message "Starting installation of NgOIOUBLMover" -Level Information
    write-NgLogMessage -Message "Azure File Share: $AzureFileShare" -Level Information
    write-NgLogMessage -Message "Install location: $InstallLocation\$FolderName" -Level Information
    write-NgLogMessage -Message "Folder name: $FolderName" -Level Information
    write-NgLogMessage -Message "Force: $Force" -Level Information
    write-NgLogMessage -Message "Run interval: $RunInterval" -Level Information

    $InstallPath = Join-Path -Path $InstallLocation -ChildPath $FolderName
    $NgScriptPath = Join-Path -Path $InstallPath -ChildPath $NgScript

    

    write-NgLogMessage -Message "Install path: $InstallPath" -Level Information
    write-NgLogMessage -Message "NgScript path: $NgScriptPath" -Level Information


    write-NgLogMessage -Message "Log folder: $LogFolder" -Level Information


    Install-NgFiles -InstallPath $InstallPath -RequiredFiles $RequiredFiles -GitHubRawUrl $GitHubRawUrl -GitHubRepoUrl $GitHubRepoUrl -Compile $NgScriptPath -Icon $IconFile -ErrorAction Stop


    if (-not $DisableStartMenuShortcut) {
        Add-NgStartMenuShortcut -ScriptPath $NgScriptPath -ScriptParameters $NgScriptParameters
        write-host "Success: Created Start menu shortcut" -ForegroundColor Green
    }

    if (-not $DisableDesktopShortcut) {
        Add-NgDesktopShortcut -ScriptPath $NgScriptPath -ScriptParameters $NgScriptParameters
        write-host "Success: Created desktop shortcut" -ForegroundColor Green
    }

    Add-NgScheduledTask -TaskName $TaskName -ScriptPath $NgScriptPath -ScriptParameters $NgScriptParameters -TaskDescription $TaskDescription -TaskInterval $RunInterval -TaskId 124563 -Disabled $DisableScheduledTask -TimeOut 15
    write-host "Success: Created Scheduled Task" -ForegroundColor Green
    write-NgLogMessage -Message "Installation of NgOIOUBLMover completed" -Level Information
    Write-Host ""
    Write-Host "##############################################" 
    Write-Host "Success: Finished installation of EAN Mover" -ForegroundColor Green
    Write-Host "##############################################"


    [System.Windows.Forms.MessageBox]::Show($THIS, "Installation of NgOIOUBLMover Completed`nInitial run starting in 2 minutes",'OIOUBL Mover','OK','Information','Button1','ServiceNotification')
}
catch {
    write-NgLogMessage -Message "Unable to install NgOIOUBLMover $_" -Level Error
    #[System.Windows.Forms.MessageBox]::Show($THIS, "Installation of NgOIOUBLMover Failed",'OIOUBL Mover','OK','error')
    [System.Windows.Forms.MessageBox]::Show($THIS, "Installation of NgOIOUBLMover Failed",'OIOUBL Mover','OK','error','Button1','ServiceNotification')

    Write-Error $_
    exit 1
}

