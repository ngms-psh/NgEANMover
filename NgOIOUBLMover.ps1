<#PSScriptInfo

.VERSION 1.0

.GUID 081c47a1-20d0-47ab-9d30-2dbac7107499

.AUTHOR Phillip Schjeldal Hansen | NgMS Consult ApS

.COMPANYNAME NgMS Consult ApS

.COPYRIGHT (c) 2024 - Phillip Schjeldal Hansen | NgMS Consult ApS. All rights reserved.

.TAGS

.LICENSEURI

.PROJECTURI

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


.PRIVATEDATA

#>

<#
.SYNOPSIS
    Move OIOUBL/EAN files from the downloads folder to mapped Azure File Share.

.DESCRIPTION
    Moves XML files with OIOBUL schema from downloads to mapped azure file share. Has the option to change the source folder, the destination folder, and the archive folder. The script can be run manually or as a scheduled task.

.INPUTS
    Description of objects that can be piped to the script.

.OUTPUTS
    Log file stored in folder c:\users\<username>\appdata\local\temp\NgOIOUBLMover

.EXAMPLE
    .\NgOIOUBLMover.ps1 -AzureStorageAccount "\\<storageaccount>.file.core.windows.net\<fileshare>" -Archive -recurse

.NOTES
    Creation Date:  11-12-2024

#>
#requires -PSEdition Desktop
[CmdletBinding()]
param (
    # Input parameters. if not provided, the default values are used
    # SourceFolder Default value: Getting the Downloads folder path from the registry
    [Parameter(HelpMessage="The folder to move the files from. Default is the Downloads folder")]
    [string]$SourceFolder,
    [Parameter(Mandatory = $true,HelpMessage="URL to the Azure File Share or the drive letter of the mapped drive")]
    [string]$AzureFileShare,
    [Parameter(HelpMessage="Use switch to disable the popup messages for failed and duplicate files")]
    [switch]$DisablePopup,
    [Parameter(HelpMessage="Use switch to enable recursive search in the source folder")]
    [switch]$Recurse,
    [Parameter(HelpMessage="Use switch to keep a copy of the original file. Default path is :<AzureFileShare>\Arkiv`nUse ArchivePath to specify a different path")]
    [switch]$Archive,
    #Default value is set further down in the script
    [Parameter(HelpMessage="Optional: Change the default archive path. Default path is: <AzureFileShare>\Arkiv")]
    [string]$ArchivePath

)
#---------------------------------------------------------[Initialisations]--------------------------------------------------------

if (!$SourceFolder) {$SourceFolder = Get-ItemPropertyValue -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name "{374DE290-123F-4565-9164-39C4925E467B}"}

# Set the log folder and log file prefix
[string]$LogFolder = "$($env:temp)\NgOIOUBLMover" # Log files will be stored in the temp folder in a folder named NgOIOUBLMover
[string]$LogFilePrefix = "Move_" # Date will be appended to the prefix ex. Move_10-12-2024.log
[int]$RetainLogs = 30 # Number of days to retain the log files


#-----------------------------------------------------------[Functions]------------------------------------------------------------
function Get-SproomDrive {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline)][string]$inp
    )
    if ($inp -match "^[d-z]$") {
        $out = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Name -eq $inp }
    }
    elseif ($inp -match ".*\.file\.core\.windows\.net\\.*"){
        $out = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.DisplayRoot -match ($inp.Replace('\','\\')) }
    }
    elseif ($inp -match ".*\.file\.core\.windows\.net\/.*") {
        $out = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.DisplayRoot -match ($inp.Replace('/','\\')) }
    }
    else {
        Write-NgLogMessage -Level Error -Message "Get-SproomDrive: '$inp' not correct format, AzureStorageAccount must be a drive letter or '<storageaccountname>.file.core.windows.net\<sharename>'"
        Write-Error "AzureStorageAccount must be a drive letter or '<storageaccountname>.file.core.windows.net\<sharename>'"
        return
    }
    if (!$out) {
        Write-NgLogMessage -Level Error -Message "Get-SproomDrive: Drive '$inp' not found"
        Write-Error "Drive $inp not found"
        return
    }
    if (!(Test-Path "$($out.Name):\")) {
        Write-NgLogMessage -Level Error -Message "Get-SproomDrive: No connection to drive '$inp'"
        Write-Error "Drive $inp not connected"
        return
    }
    return $out
}

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
        switch -wildcard ($M) {
            "Duplicate file*" { Add-Content -Path "$LogFolder\$($LogFilePrefix)duplicates_$StartTime.log" -Value $M }
            "Failed*" { Add-Content -Path "$LogFolder\$($LogFilePrefix)failed_$StartTime.log" -Value $M }
        }
    }
}

function Clear-NgOldLog {
    $OldLogFiles = Get-ChildItem -Path $LogFolder -Filter "$LogFilePrefix*.log" | Where-Object {[datetime]::ParseExact((($_.Name | Select-String -Pattern "\d+-\d+-\d{4}").Matches.Value), "dd-MM-yyyy",$null) -lt (Get-Date).AddDays(-$RetainLogs)}
    $LogMessage = @()
    $OldLogFiles | ForEach-Object {
        $LogMessage += "Removing old log file '$($_.Name)'"
    }
    $OldLogFiles | Remove-Item -Force
    Write-NgLogMessage -Level Warning -Message $LogMessage
}

function Open-NgLogFile {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [validateSet('Full', 'Failed', 'Duplicates')]
        [string]$Type
    )
    switch ($type) {
        "Full" { $LogFile = "$LogFolder\$LogFilePrefix$(get-date -Format 'dd-MM-yyyy').log" }
        "Failed" { $LogFile = "$LogFolder\$LogFilePrefix" + "failed_$StartTime.log" }
        "Duplicates" { $LogFile = "$LogFolder\$LogFilePrefix" + "duplicates_$StartTime.log" }
    }

    try {
        Invoke-Item $LogFile
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($THIS, "Failed to open log file '$LogFile'",'OIOUBL Mover','OK','warning')
    }
}

function Show-NgNotification {
    [cmdletbinding()]
    [OutputType([Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime])]
    Param (
        [string]$ToastTitle,
        [parameter(ValueFromPipeline)]
        [string]$ToastText
    )

    process {
        [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null

        $Template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)

        $RawXml = [xml] $Template.GetXml()
        ($RawXml.toast.visual.binding.text | Where-Object {$_.id -eq "1"}).AppendChild($RawXml.CreateTextNode($ToastTitle)) > $null
        ($RawXml.toast.visual.binding.text | Where-Object {$_.id -eq "2"}).AppendChild($RawXml.CreateTextNode($ToastText)) > $null

        $SerializedXml = New-Object Windows.Data.Xml.Dom.XmlDocument
        $SerializedXml.LoadXml($RawXml.OuterXml)

        $Toast = [Windows.UI.Notifications.ToastNotification]::new($SerializedXml)
        $Toast.Tag = "Ng EAN Mover"
        $Toast.Group = "Ng EAN Mover"
        #$Toast.ExpirationTime = [DateTimeOffset]::Now.AddMinutes(1)

        $Notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("Ng EAN Mover")
        $Notifier.Show($Toast);
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

function Move-OIOUBLFile {
    param (
    # Parameters for the function
        [Parameter(Mandatory = $true, Position = 0)]
        [System.IO.FileInfo]$SourceFile,
        [Parameter(Mandatory = $true, Position = 1)]
        [ValidatePattern("^[a-zA-Z]:\\.*$|^\\\\[^\\]+\\[^\\].*$")]
        [string]$Destination
    )

    if ((Test-Path "$Destination$($SourceFile.Name)") -or (Test-Path $Destination -PathType Leaf)) {
        Write-NgLogMessage -Level Error -Message "Duplicate file '$($SourceFile.FullName)' already in destination '$Destination', skipping the file"
        $Results.SkippedOIOUBLFiles += $SourceFile.FullName
        $Results.TotalFiles++
        return
    }

    # Move file if archive switch is not used
    try {
        # Move file if archive switch is not used
        if (!$Archive){
            Move-Item -Path $SourceFile.FullName -Destination $Destination
            Write-NgLogMessage -Level Information -Message "Moved file '$($SourceFile.FullName)' to '$Destination'"
            $Results.MovedFiles += $SourceFile.FullName
            $Results.TotalFiles++
            return
        }

        # Move file and create archive folder if archive switch is used
        else {
            # If ArchivePath is not set, use the default path
            if (!$ArchivePath) {$ArchivePath = "$($Drive):\Arkiv"}

            # Create the Archive folder if it does not exist
            if (!(Test-Path "$ArchivePath" -PathType Container)) {
                New-Item -Path "$ArchivePath" -ItemType Directory | Out-Null
                write-NgLogMessage -Level Information -Message "Created archive folder '$ArchivePath'"
            }

            # if the file already exists in the archive folder, log the error and skip the file
            if (test-path -Path "$ArchivePath\$($SourceFile.Name)"){
                Write-NgLogMessage -Level Error -Message "Duplicate archive file '$($SourceFile.FullName)'-  archive folder: '$ArchivePath', skipping archive copy, still moving file to destination"
            }
            # If the file does not exist in the archive folder, copy it
            else {
                Copy-Item -Path $SourceFile.FullName -Destination "$ArchivePath\$($SourceFile.Name)"
                Write-NgLogMessage -Level Information -Message "Copied file '$($SourceFile.FullName)' to archive '$ArchivePath\$($SourceFile.Name)'"
            }

            # Move the file to the destination folder, regardless of the archive copy
            Move-Item -Path $SourceFile.FullName -Destination $Destination
            Write-NgLogMessage -Level Information -Message "Moved file '$($SourceFile.FullName)' to destination '$Destination'"
            $Results.MovedFiles += $SourceFile.FullName
            $Results.TotalFiles++
            return
        }
    }
    catch {
        Write-NgLogMessage -Level Error -Message "Failed to move file '$($SourceFile.FullName)' to '$Destination'"
        Write-NgLogMessage -Level Error -Message "$_"
        $Results.FailedFiles += $SourceFile.FullName
        $Results.TotalFiles++
        Write-Error "Failed to move file '$($SourceFile.FullName)' to '$Destination'"
        return $_.Exception.Message
    }
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------
Add-NgStartMenuShortcut
# Start the script
try {
    Write-NgLogMessage Information "-------------------------------------------------------------------[Starting script]-------------------------------------------------------------------"
    Write-NgLogMessage Information "Source folder: '$SourceFolder'"
    Write-NgLogMessage Information "Azure File Share: '$AzureFileShare'"
    Write-NgLogMessage Information "Archive: '$([bool]$Archive)'"
    write-NgLogMessage Information "ArchivePath: '$ArchivePath'"
    Write-NgLogMessage Information "DisablePopup: '$([bool]$DisablePopup)'"
    Write-NgLogMessage Information "Recurse: '$([bool]$Recurse)'"
    Write-NgLogMessage Information "----------------------------------------------------------------------------------------------------------------------------------------------------"

    $StartTime = Get-Date -Format "dd-MM-yyyy_HHmmss"

    # Initialize the results hashtable
    $Results = @{
        SkippedOIOUBLFiles = @()
        NonOIOUBL = 0
        MovedFiles = @()
        FailedFiles = @()
        TotalFiles = 0
        Status = [string]::Empty
    }

    # Get the Azure File Share drive letter
    try {
        $Drive =  $AzureFileShare | Get-SproomDrive -ErrorAction Stop
    }
    # If the drive is not found, show a notification and exit the script
    catch {
        Show-NgNotification -ToastTitle "Results - Failed" -ToastText "Cannot find the Azure File Share"
        exit $_.Exception.Message
    }


    #-----------------------------[Get files from source]-----------------------------

    #Get all XML files in the source folder and subfolders if Recurse switch is used
    if($Recurse){$ImportFiles = Get-ChildItem -Path $SourceFolder -Filter "*.xml" -Force -Recurse}
    else {$ImportFiles = Get-ChildItem -Path $SourceFolder -Filter "*.xml" -Force}
    #---------------------------------------------------------------------------------


    # Check if any XML files are found in the source folder
    if ($ImportFiles.Count -eq 0) {
        Write-NgLogMessage -Level Information -Message "No XML files found in source folder '$SourceFolder'"
        $Results.TotalFiles = 0
    }

    # If XML files are found, process them
    else {
        foreach ($ImportFile in $ImportFiles) {
            [xml]$ImportFileContent = Get-Content -Path $ImportFile.FullName
            # Check if the file is an OIOUBL file
            if (!($ImportFileContent.DocumentElement.CustomizationID -match "OIOUBL")) {
                # If not, skip the file
                Write-NgLogMessage -Level Warning -Message "Skipping file '$($ImportFile.FullName)' as it is not an OIOUBL file"
                $Results.NonOIOUBL++
                $Results.TotalFiles++
                continue
            }
            # if the file is an OIOUBL file, move it to the destination folder
            elseif ($ImportFileContent.DocumentElement.CustomizationID -match "OIOUBL") {
                Move-OIOUBLFile -SourceFile $ImportFile -Destination "$($Drive.name):\"
            }
            # If the first OIOUBL validation fails, log the error and skip the file
            else {
                Write-NgLogMessage -Level Error -Message "Failed to determine XML Content type of file '$($ImportFile.FullName)'"
                Write-Output "Failed to determine the file type of '$($ImportFile.FullName)'"
                $Results.FailedFiles += $ImportFile.FullName
                $Results.TotalFiles++
            }
        }
    }
    if (!$DisablePopup) {
        # Display MessageBox with failed Files and promt to open log file
        if ($Results.FailedFiles.Count -ne 0) {
            $ShowError = [System.Windows.Forms.MessageBox]::Show($THIS, "Failed to move '$($Results.FailedFiles.Count)' EAN files `nShow log details?",'OIOUBL Mover','YesNo','error')
            if ($ShowError -eq 'Yes') {
                Open-NgLogFile Failed
            }
        }


        # Display MessageBox with duplicate Files and promt to open log file
        if ($Results.SkippedOIOUBLFiles.Count -ne 0) {
            $ShowSkipped = [System.Windows.Forms.MessageBox]::Show($THIS, "Skipped '$($Results.SkippedOIOUBLFiles.Count)' OIOUBL file(s), already exists in destination folder`n`nCheck for duplicates or contact sproom to verify they have connection to SFTP`n`nShow log details?",'OIOUBL Mover','YesNo','warning')
            if ($ShowSkipped -eq 'Yes') {
                Open-NgLogFile Duplicates
            }
        }
    }
    switch ($Results) {
        {$_.FailedFiles.Count -eq 0} { $Results.Status = "Completed" }
        {$_.FailedFiles.Count -gt 0} { $Results.Status = "Failed" }
        {$_.SkippedOIOUBLFiles.Count -gt 0} { $Results.Status = "Duplicates" }

    }
    Show-NgNotification -ToastTitle "Results - $($Results.Status)" -ToastText "Success: $($Results.MovedFiles.Count)`nOther XML: $($Results.NonOIOUBL)`nDuplicates: $($Results.SkippedOIOUBLFiles.Count)`nFailed: $($Results.FailedFiles.Count)`nTotal: $($Results.TotalFiles)"
    Clear-NgOldLog
    Write-NgLogMessage Information "Results: '$($Results.TotalFiles)' files processed, '$($Results.MovedFiles.Count)' moved, '$($Results.FailedFiles.Count)' failed, '$($Results.SkippedOIOUBLFiles.Count)' duplicate, '$($Results.NonOIOUBL)' non-OIOUBL files"
    Write-NgLogMessage Information "---------------------------------------------------------[Script Completed - $($Results.Status)]---------------------------------------------------------"
}
catch {
    Show-NgNotification -ToastTitle "Results - Failed" -ToastText "Success: $($Results.MovedFiles.Count)`nOther XML: $($Results.NonOIOUBL)`nDuplicates: $($Results.SkippedOIOUBLFiles.Count)`nFailed: $($Results.FailedFiles.Count)`nTotal: $($Results.TotalFiles)"
    Write-NgLogMessage Information "Results: '$($Results.TotalFiles)' files processed, '$($Results.MovedFiles.Count)' moved, '$($Results.FailedFiles.Count)' failed, '$($Results.SkippedOIOUBLFiles.Count)' duplicate, '$($Results.NonOIOUBL)' non-OIOUBL files"
    Write-NgLogMessage Information "--------------------------------------------------------------[Script Completed - failed]--------------------------------------------------------------"
}