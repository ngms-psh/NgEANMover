<#PSScriptInfo

.VERSION 2.0

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
Using namespace System.IO.Compression.FileSystem
Using module .\Modules\NgOIOUBL
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
    [string]$ArchivePath,
    [Parameter(HelpMessage="Optional: Max size (MB) of Zip files to scan for OIOUBL. Default path is: 1 MB")]
    [int]$ZipMaxSizeMB

)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------



# Set the log folder and log file prefix
[string]$LogFolder = "$($env:temp)\NgOIOUBLMover" # Log files will be stored in the temp folder in a folder named NgOIOUBLMover
[string]$LogFilePrefix = "Move_" # Date will be appended to the prefix ex. Move_10-12-2024.log
[int]$RetainLogs = 30 # Number of days to retain the log files
[int]$ZipMaxSizeMB = 5 # Max size(MB) of ZIP files to scan

# Initialize the results hashtable
$Results = @{
    SkippedOIOUBLFiles = [int]0
    NonOIOUBL = [int]0
    MovedFiles = [int]0
    FailedFiles = [int]0
    TotalFiles = [int]0
    TotalOIOUBL = [int]0
    Status = [string]::Empty
    Errors = [int]0
    FailedSources = [int]0
}

Add-Type -AssemblyName PresentationCore,PresentationFramework,System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression.FileSystem

#Import-Module .\Modules\NgOIOUBL


#-------------------------------------------------------------[Classes]--------------------------------------------------------------

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
    if (!(Test-Path $out.root)) {
        Write-NgLogMessage -Level Error -Message "Get-SproomDrive: No connection to drive '$inp'"
        Write-Error "Drive $inp not connected"
        return
    }
    return $out.Root
}

function Write-NgLogMessage {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true)]
        $Message,
        [Parameter(Mandatory = $true, Position = 0)]
        [validateSet('Error', 'Warning', 'Information')]
        [string]$Level
    )

    #$ParameterList = (Get-Command -Name $MyInvocation.MyCommand).Parameters
    #$MaxLength = ($ParameterList["Level"].Attributes.ValidValues | Sort-Object { $_.Length } -Descending | Select-Object -First 1).Length

    begin {
        $ParameterList = @('Error', 'Warning', 'Information')
        $MaxLength = ($ParameterList | Sort-Object { $_.Length } -Descending | Select-Object -First 1).Length
        $LogFile = Join-Path $LogFolder ("$LogFilePrefix$(get-date -Format 'dd-MM-yyyy').log")

        If (!(Test-Path $LogFolder)) {New-Item -Path $LogFolder -Type Directory -Force | Out-Null}
    }
    process {
        # Pad the message to the maximum length
        $LevelPadded = $Level.PadRight($MaxLength)

        if ($Level -eq "Error"){$Results.Errors++}
        
        foreach ($M in $Message) {
            $Date = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
            $FullM = "$Date | $LevelPadded - $M"
            Add-Content -Path $LogFile -Value $FullM -Force
            switch -wildcard ($M) {
                "Skipped*" { Add-Content -Path  (Join-Path $LogFolder ("$($LogFilePrefix)duplicates_$StartTime.log")) -Value $M }
                "Failed*" { Add-Content -Path (Join-Path $LogFolder ("$($LogFilePrefix)failed_$StartTime.log")) -Value $M }
                "Success*" { Add-Content -Path (Join-Path $LogFolder ("$($LogFilePrefix)success_$StartTime.log")) -Value $M }
            }
        }
    }
    
}

function Clear-NgOldLog {
    Param (
        [Parameter(Mandatory)]
        [int]$OlderThanDays
    )
    try {
        $OldLogFiles = Get-ChildItem -Path $LogFolder -Filter "$LogFilePrefix*.log" | Where-Object {[datetime]::ParseExact((($_.Name | Select-String -Pattern "\d+-\d+-\d{4}").Matches.Value), "dd-MM-yyyy",$null) -lt (Get-Date).AddDays(-$RetainLogs)}
        $LogMessage = @()
        $OldLogFiles | ForEach-Object {
            $LogMessage += "Removing old log file '$($_.Name)'"
        }
        $OldLogFiles | Remove-Item -Force
        Write-NgLogMessage -Level Warning -Message $LogMessage
    }
    catch {
        Write-NgLogMessage -Level Error -Message "Clear-NgOldLog: Failed to Clear old log files"
        Write-NgLogMessage -Level Error -Message "$_"
    }
    
}

function Open-NgLogFile {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [validateSet('Full', 'Failed', 'Duplicates','Success')]
        [string]$Type
    )
    switch ($type) {
        "Full" { $LogFile = "$LogFolder\$LogFilePrefix$(get-date -Format 'dd-MM-yyyy').log" }
        "Failed" { $LogFile = "$LogFolder\$LogFilePrefix" + "failed_$StartTime.log" }
        "Duplicates" { $LogFile = "$LogFolder\$LogFilePrefix" + "duplicates_$StartTime.log" }
        "Success" { $LogFile = "$LogFolder\$LogFilePrefix" + "success_$StartTime.log" }
    }

    try {
        Invoke-Item $LogFile
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($THIS, "Failed to open log file '$LogFile'",'EAN Mover','OK','warning','Button1','ServiceNotification')
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

function Move-OIOUBLFiles {
    [CmdletBinding()]
    param (
    # Parameters for the function
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true)]
        [NgOIOUBL]$Source,
        [Parameter(Mandatory = $true, Position = 1)]
        [ValidatePattern("^[a-zA-Z]:\\.*$|^\\\\[^\\]+\\[^\\].*$")]
        [string]$Destination,
        [bool]$Archive = $Archive,
        [string]$ArchivePath = $ArchivePath

    )
    
    begin {
        #Initialize array for function results
        $SourceRes = @()
    }


    process {
        try {
            #Initialize variables for process results
            $Failed = @()
            $Skipped = @()
            $Success = @()

            Write-Verbose "Move-OIOUBLFiles: Now Processing $($Source.Source)"

            #Get unique entries in source and test against destination
            $UniqueEntries = $Source.GetUniqueEntries($Destination)

            #if there are duplicates, skip them and log file name
            if ($UniqueEntries.count -lt $Source.OIOUBLCount) {
                $Source.GetDuplicatesEntries($Destination) | ForEach-Object {
                    Write-NgLogMessage -Level Warning -Message "Skipped: Duplicate file '$($_.FullPath)' already in destination '$Destination', skipping the file"
                    $_.status = "Skipped"
                    $Skipped += $_
                    $Results.SkippedOIOUBLFiles++
                    $Results.TotalFiles++
                }
            }

            #if no unique entries, log and skip processing further
            if($UniqueEntries.count -eq 0){
                Write-NgLogMessage -Level Information -Message "No unique OIOUBL XML Files found in '$($Source.Source)'"

            }
            #Else continue to process unique entries
            else {
                # Copy unique entries to destination, overwrite false. It will test for duplicates for every file again
                $res = $Source.BulkCopyEntries($UniqueEntries, $Destination, $false)

                # Write each successfully copied file to log
                $Success += $res.Success
                $res.Success | ForEach-Object {
                    Write-NgLogMessage -Level Information -Message "Success: Copied '$($_.FullPath)' to '$Destination'"
                    $Results.MovedFiles++
                    $Results.TotalFiles++
                }

                # If all unique entries was copied successfully, and no further duplicates was found write to log
                if ($res.Success.count -eq $UniqueEntries.count){
                    Write-NgLogMessage -Level Information -Message "Success: Copied all unique OIOUBL XML Files '$($res.Success.count)' of '$($UniqueEntries.count)' from '$($Source.Source)'"
                }

                # if errors or duplicates was found, log this
                else {
                    if ($res.Skipped) {
                        $Skipped += $res.Skipped
                        $res.Skipped | ForEach-Object {
                            Write-NgLogMessage -Level Warning -Message "Skipped: Duplicate file '$($_.FullPath)' already in destination '$Destination', skipping the file"
                            $Results.SkippedOIOUBLFiles++
                            $Results.TotalFiles++
                        }
                    }
                    elseif ($res.Failed) {
                        $Failed += $res.Failed
                        $res.Failed | ForEach-Object { 
                            Write-NgLogMessage -Level Error -Message "Failed: To Copy '$($_.FullPath)' to destination '$Destination', skipping the file"
                            $Results.FailedFiles++
                            $Results.TotalFiles++
                        }
                    }
                }



                # If archive is enabled copy entries to archive folder and log the results. Overwrite archive file if found
                if ($Archive){
                    try {
                        $ArhRes = $Source.BulkCopyEntries($res.Success, $ArchivePath, $true)
                        $ArhRes.Success | ForEach-Object {
                            Write-NgLogMessage -Level Information -Message "Success: Archived '$($_.FullPath)' to '$ArchivePath'"
                        }
                        $ArhRes.Failed | ForEach-Object {
                            Write-NgLogMessage -Level Information -Message "Failed: Archived '$($_.FullPath)' to '$ArchivePath'"
                            $Results.FailedFiles++
                        }
                    }
                    catch {
                        Write-NgLogMessage -Level Error -Message "Move-OIOUBLFiles: Unknown Error: Failed to Archive files from '$($Source.Source)' to '$ArchivePath'"
                        Write-NgLogMessage -Level Error -Message "$_"
                        $results.FailedSources++
                    }
                }
            }

            # Create NgOIOUBL custom class object with results for the operation. Used in clean up
            $tempRes = [NgOIOUBL]::new()
            $tempRes.Source = $Source.Source
            $tempRes.Type = $Source.Type
            $tempRes.OIOUBLCount = $Source.OIOUBLCount
            $tempRes.OtherCount = $Source.OtherCount
            $tempRes.entries = [PSCustomObject]@{
                Failed = $Failed 
                Skipped = $Skipped
                Success = $Success 
            }

            # Determine status for the operation and write to temp NgOIOUBL object
            switch ($tempRes.entries) {
                # All files in source proccess without errors and source does not contain other files. If type 'folder' other = only XML files, for type 'zip' other = any other file in ZIP 
                {(($tempRes.entries.Success.count + $tempRes.entries.Skipped.count) -eq $tempRes.OIOUBLCount) -and ($tempRes.OtherCount -eq 0)}  { $tempRes.status = "Proccessed" }

                # All files in source proccess without errors, but source contains other files
                {(($tempRes.entries.Success.count + $tempRes.entries.Skipped.count) -eq $tempRes.OIOUBLCount) -and ($tempRes.OtherCount -ne 0)}  { $tempRes.status = "ProccessedWithOthers" }
                
                # If no entries was moved, all failed
                {$tempRes.entries.Failed.count  -eq $tempRes.OIOUBLCount}  { $tempRes.status = "Failed" }

                # If some failed but other was successfull
                {($tempRes.entries.Failed.count) -and ($tempRes.entries.Success.count -ne 0)} { $tempRes.status = "ProccessedWithErrors" }
                
                #Else
                Default {$tempRes.status = "Unknown"}
            }

            
            write-verbose "OIOUBLCount: $($tempRes.OIOUBLCount | Out-String)"
            write-verbose "Success: $($tempRes.entries.Success.count | Out-String)"
            #write-verbose "$($tempRes.entries.Success | Out-String)"

            write-verbose "Skipped: $($tempRes.entries.Skipped.count | Out-String)"
            #write-verbose "$($tempRes.entries.Skipped | Out-String)"
            write-verbose "Failed: $($tempRes.entries.Failed.count| Out-String)"
            #write-verbose "$($tempRes.entries.Failed| Out-String)"
            write-verbose "OtherCount: $($tempRes.OtherCount | Out-String)"
            write-verbose "Status: $($tempRes.Status | Out-String)"
            Write-Verbose "Move-OIOUBLFiles: Finished Processing $($Source.Source)"
            # Write results saved in temp NgOIOUBL object to results array for the function. Then continue to next source obecjt in pipeline
            $SourceRes += $tempRes

        }

        catch {
            Write-NgLogMessage -Level Error -Message "Move-OIOUBLFiles: Unknown Error: Failed to move files from '$($Source.Source)' to '$Destination'"
            Write-NgLogMessage -Level Error -Message "$_"
            $Source.Status = "Failed"
            $SourceRes += $Source
            $results.FailedSources++
        }
    }
    end {
        try {
                    # Start clean up proccess. Delete success entries and duplicates. Will not delete failed entries
            if ($SourceRes) {
                Write-NgLogMessage Information "-------------------------------------------------------------------[Clean up]-------------------------------------------------------------------"
                foreach ($R in $SourceRes){
                    Write-NgLogMessage -Level Information -Message "Cleaning up source '$($R.Source)'"
                    switch ($R.Status) {
                        "Proccessed" { 
                            $DelRes = $R.delete() 
                            if ($DelRes.failed.count -eq 0) {
                                Write-NgLogMessage -Level Information -Message "Success: Deleted source file '$($R.Source)'"
                            }
                            else {
                                $DelRes.failed | ForEach-Object {
                                    Write-NgLogMessage -Level Warning -Message "Failed: to delete source file '$($R.Source)'"
                                    $Results.FailedFiles++
                                }
                            }
                        }
                        "ProccessedWithOthers" { 
                            $Deletes = $R.Entries.Skipped + $R.Entries.Success
                            $DelRes = $R.BulkDeleteEntries($Deletes)
                            if ($DelRes.failed.count -eq 0) {
                                $DelRes.Success | ForEach-Object {
                                    Write-NgLogMessage -Level Information -Message "Success: Deleted entry file '$($_.FullPath)'"
                                }
                                
                            }
                            else {
                                $DelRes.failed | ForEach-Object {
                                    Write-NgLogMessage -Level Warning -Message "Failed: to delete entry file '$($_.FullPath)'"
                                    $Results.FailedFiles++
                                }
                            }
                        }
                        "ProccessedWithErrors" { 
                            $Deletes = $R.Entries.Skipped + $R.Entries.Success
                            $DelRes = $R.BulkDeleteEntries($Deletes) 
                            if ($DelRes.failed.count -eq 0) {
                                $DelRes.Success | ForEach-Object {
                                    Write-NgLogMessage -Level Information -Message "Success: Deleted entry file '$($_.FullPath)'"
                                }
                            }
                            else {
                                $DelRes.failed | ForEach-Object {
                                    Write-NgLogMessage -Level Warning -Message "Failed: to delete entry file '$($_.FullPath)'"
                                    $Results.FailedFiles++
                                }
                            }
                        }
                        "Unknown" {
                            Write-NgLogMessage -Level Error -Message "Failed: Could not determine status of move operation for source '$($R.Source)'"
                            $Results.FailedFiles++
                        }
                    }
                }   
            }
        }
        catch {
            Write-NgLogMessage -Level Error -Message "Failed: Move-OIOUBLFiles failed to delete moved and skipped files"
            Write-NgLogMessage -Level Error -Message "----------------------------------------------------------------[Clean up - Failed]----------------------------------------------------------------"
        }

    }
}

function Get-CompressedOIOUBL {
    param (
        [Parameter(Mandatory = $true)]
        $SourceFolder,
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [int]$MaxSizeMB = $ZipMaxSizeMB,
        [bool]$Recurse = $Recurse
    )

    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        # Convert the MaxSizeMB to bytes
        $MaxSize = $MaxSizeMB * 1048576    

        # Get all zip files in the source folder and filter for size
        if ($Recurse){ $ZipFiles = Get-ChildItem -Path $SourceFolder -Filter "*.zip" -Recurse }
        else { $ZipFiles = Get-ChildItem -Path $SourceFolder -Filter "*.zip" }
        
        $CorrectSizeZipFiles = $ZipFiles | Where-Object { $_.Length -lt $MaxSize }

        # Log the zip files that are too large
        $ZipFiles | Where-Object { $_ -notin $CorrectSizeZipFiles } | ForEach-Object {
            Write-NgLogMessage -Level Warning -Message "Skip checking zip file for XML: '$($_.FullName)' is larger than $MaxSizeMB MB"
        } 

        $Out = $CorrectSizeZipFiles | ForEach-Object { [NgOIOUBL]::new($_) }

        foreach ($OutZip in $Out){
            if (($OutZip.Status -eq "Failed") -or ($null -eq $Out)) {
                Write-NgLogMessage -Level Error -Message "Failed: to Scan '$($OutZip.Source)' for OIOUBL Files"
                $Results.FailedFiles++
            }
            else {
                Write-NgLogMessage -Level Information -Message "OIOUBL XML files '$($OutZip.OIOUBLCount)' found in zip file '$($OutZip.Source)'"
                $Results.TotalOIOUBL += $OutZip.OIOUBLCount

                if ($OutZip.OtherCount -ne 0) {
                    Write-NgLogMessage -Level Information -Message "Other files '$($OutZip.OtherCount)' found in '$($OutZip.Source)'"
                    $Results.NonOIOUBL += $OutZip.OtherCount 
                    $Results.TotalFiles += $OutZip.OtherCount 
                }
                if ($OutZip.Entries.Failed.count -ne 0) {
                    $OutZip.Entries.Failed | ForEach-Object {
                        Write-NgLogMessage -Level Error -Message "Failed: to check if file is OIOUBL '$($_.FullName)' in Source '$($_.Parrent)'"
                        Write-NgLogMessage -Level Error -Message "$($_ | out-string)"
                        $Results.FailedFiles++
                    }
                }
            }
        }

        return $Out
    }
    catch {
        Write-NgLogMessage -Level Error -Message "Get-CompressedOIOUBL: Failed to process zip files from source folder '$SourceFolder'"
        Write-NgLogMessage -Level Error -Message "$_"
        throw "Failed to get zip files from source folder '$SourceFolder' $_"
    }
    
}
function Get-OIOUBL {
    param (
        [Parameter(Mandatory = $true)]
        $SourceFolder,
        [bool]$Recurse = $Recurse
    )

    try {
       
        $Out = [NgOIOUBL]::new($SourceFolder, $Recurse)

        if (($Out.Status -eq "Failed") -or ($null -eq $Out)) {
            Write-NgLogMessage -Level Error -Message "Failed: to Scan '$($Out.Source)' for OIOUBL Files"
            $Results.FailedFiles++
        }
        else {
            Write-NgLogMessage -Level Information -Message "OIOUBL XML files '$($Out.OIOUBLCountCount)' found in Source '$($Out.Source)'"
            $Results.TotalOIOUBL += $Out.OIOUBLCount

            if ($Out.OtherCount -ne 0) {
                Write-NgLogMessage -Level Information -Message "Other XML files '$($Out.OtherCount)' found in '$($Out.Source)'"
                $Results.NonOIOUBL += $Out.OtherCount 
                $Results.TotalFiles += $Out.OtherCount 
            }
            if ($Out.Entries.Failed.count -ne 0) {
                $OutZip.Entries.Failed | ForEach-Object {
                    Write-NgLogMessage -Level Error -Message "Failed: to check if file is OIOUBL '$($_.FullName)' in Source '$($_.Parrent)'"
                    $Results.FailedFiles++
                }
            }
        }
        
        return $Out
    }
    catch {
        Write-NgLogMessage -Level Error -Message "Get-OIOUBL: Failed to process zip files from source folder '$SourceFolder'"
        Write-NgLogMessage -Level Error -Message "$_"
        throw "Failed to get zip files from source folder '$SourceFolder' $_"
    }
    
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------
# Start the script
try {
    Write-NgLogMessage Information "-------------------------------------------------------------------[Starting script]-------------------------------------------------------------------"
    #check if the proccess is already running
    if ((Get-Process -Name "NgOIOUBLMover" -ErrorAction SilentlyContinue).count -gt 1) {
        Write-NgLogMessage -Level Error -Message "Process already running, terminating script"
        Write-NgLogMessage Information "--------------------------------------------------------------[Script - failed]--------------------------------------------------------------"
        Show-NgNotification -ToastTitle "Results - Failed" -ToastText "Process already running"
        $ShowError = [System.Windows.Forms.MessageBox]::Show($THIS, "EAN Mover already running`nPlease wait for it to complete before running EAN Mover again",'EAN Mover','OK','error','Button1','ServiceNotification')
        exit "Process already running"
    }
    

    #Get start time, used for log file names 
    $StartTime = Get-Date -Format "dd-MM-yyyy_HHmmss"

# Check if the source folder exists
    if ($SourceFolder){
        if (!(Test-Path $SourceFolder)) {
            Write-NgLogMessage -Level Error -Message "Source folder '$SourceFolder' does not exist, terminating script"
            Show-NgNotification -ToastTitle "Results - Failed" -ToastText "Source folder '$SourceFolder' does not exist"
            $ShowError = [System.Windows.Forms.MessageBox]::Show($THIS, "Failed to move EAN files, Source folder '$SourceFolder' does not exist`nShow log details?",'EAN Mover','YesNo','error','Button1','ServiceNotification')
            if ($ShowError -eq 'Yes') {
                Open-NgLogFile Failed
            }
            exit "Source folder '$SourceFolder' does not exist"
        }
    }
    # else use the downloads folder
    else {
        try {
            $SourceFolder = Get-ItemPropertyValue -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name "{374DE290-123F-4565-9164-39C4925E467B}"
        }
        catch {
            try {
                $SourceFolder = (New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path
                if (!(Test-Path $SourceFolder)) {
                    if (Test-Path "$env:USERPROFILE\Downloads") {
                        $SourceFolder = "$env:USERPROFILE\Downloads"
                    }
                    else {
                        Write-NgLogMessage -Level Error -Message "Cannot find the Downloads folder, terminating script"
                        Show-NgNotification -ToastTitle "Results - Failed" -ToastText "Cannot find the Downloads folder"
                        $ShowError = [System.Windows.Forms.MessageBox]::Show($THIS, "Failed to move EAN files, Cannot find the Downloads folder`nShow log details?",'EAN Mover','YesNo','error','Button1','ServiceNotification')
                        if ($ShowError -eq 'Yes') {
                            Open-NgLogFile Failed
                        }
                        exit $_.Exception.Message
                    }
                }
            }
            catch {
                if (Test-Path "$env:USERPROFILE\Downloads") {
                    $SourceFolder = "$env:USERPROFILE\Downloads"
                }
                else {
                    Write-NgLogMessage -Level Error -Message "Cannot find the Downloads folder, terminating script"
                    Show-NgNotification -ToastTitle "Results - Failed" -ToastText "Cannot find the Downloads folder"
                    $ShowError = [System.Windows.Forms.MessageBox]::Show($THIS, "Failed to move EAN files, Cannot find the Downloads folder`nShow log details?",'EAN Mover','YesNo','error','Button1','ServiceNotification')
                    if ($ShowError -eq 'Yes') {
                        Open-NgLogFile Failed
                    }
                    exit $_.Exception.Message
                }
            }
        }
        
    }

    # Get the Azure File Share drive letter
    try {
        $Drive =  $AzureFileShare | Get-SproomDrive -ErrorAction Stop
    }
    # If the drive is not found, show a notification and exit the script
    catch {
        Write-NgLogMessage -Level Error -Message "Cannot find Azure File Share, terminating script"
        Show-NgNotification -ToastTitle "Results - Failed" -ToastText "Cannot find the Azure File Share"
        exit $_.Exception.Message
    }

    #set ArchivePath if Archive switch is used
    if ($Archive) {
        if (!$ArchivePath) {$ArchivePath = Join-Path $Drive "Arkiv"}
        if (!(Test-Path "$ArchivePath" -PathType Container)) {
            New-Item -Path "$ArchivePath" -ItemType Directory | Out-Null
            write-NgLogMessage -Level Information -Message "Created archive folder '$ArchivePath'"
        }
    }

    # Write the start message to the log file
    Write-NgLogMessage Information "Source folder: '$SourceFolder'"
    Write-NgLogMessage Information "Azure File Share: '$AzureFileShare'"
    Write-NgLogMessage Information "Destination: '$Drive'"
    Write-NgLogMessage Information "ZIP files Max Size(MB): '$ZipMaxSizeMB MB'"
    Write-NgLogMessage Information "Archive: '$([bool]$Archive)'"
    write-NgLogMessage Information "ArchivePath: '$ArchivePath'"
    Write-NgLogMessage Information "DisablePopup: '$([bool]$DisablePopup)'"
    Write-NgLogMessage Information "Recurse: '$([bool]$Recurse)'"
    Write-NgLogMessage Information "RetainLogs: '$RetainLogs' days"
    Write-NgLogMessage Information "----------------------------------------------------------------------------------------------------------------------------------------------------"


    

    #-----------------------------[Get files from source]-----------------------------

    #Get all OIOUBL XML files in the source folder and ZIP files. Recurse switch is optional
    try {
        $OIOUBLSources = @()
        $OIOUBLSources += Get-OIOUBL -SourceFolder $SourceFolder
        $OIOUBLSources += Get-CompressedOIOUBL -SourceFolder $SourceFolder -MaxSizeMB $ZipMaxSizeMB
    }
    catch {
        Write-NgLogMessage -Level Error -Message "Failed to scan for OIOUBL files, terminating script"
        Write-NgLogMessage -Level Error -Message "$_"
        Show-NgNotification -ToastTitle "Results - Failed" -ToastText "Failed to scan for OIOUBL files"
        $ShowError = [System.Windows.Forms.MessageBox]::Show($THIS, "Failed to scan for OIOUBL files`nShow log details?",'EAN Mover','YesNo','error','Button1','ServiceNotification')
        if ($ShowError -eq 'Yes') {
            Open-NgLogFile Failed
        }
        exit $_.Exception.Message
    }
 
    # Check if any XML files are found in Sources 
    if (($OIOUBLSources.OIOUBLCount | Measure-Object -Sum).Sum -eq 0){
        Write-NgLogMessage -Level Information -Message "No OIOUBL XML files to process"
    }
    else{

        # Writes the number of sources and files to the log
        $ProcessSources = $OIOUBLSources | Where-Object {$_.OIOUBLCount -ne 0}
        Write-NgLogMessage -Level Information -Message "Found '$($ProcessSources.count)' sources to process"
        Write-NgLogMessage -Level Information -Message "Found '$(($OIOUBLSources.OIOUBLCount | Measure-Object -Sum).Sum)' OIOUBL XML files to process"


        #-----------------------------[Move files to destination]-----------------------------
        ##Main Process###
        $ProcessSources | Move-OIOUBLFiles -Destination $Drive -Archive $Archive 

    }

    if (!$DisablePopup) {
        # Display MessageBox with failed Files and promt to open log file
        if ($Results.FailedFiles -ne 0) {
            $ShowError = [System.Windows.Forms.MessageBox]::Show($THIS, "Failed to move '$($Results.FailedFiles)' EAN files `nShow log details?",'EAN Mover','YesNo','error','Button1','ServiceNotification')
            if ($ShowError -eq 'Yes') {
                Open-NgLogFile Failed
            }
        }


        # Display MessageBox with duplicate Files and promt to open log file
        if ($Results.SkippedOIOUBLFiles -ne 0) {
            $ShowSkipped = [System.Windows.Forms.MessageBox]::Show($THIS, "Skipped '$($Results.SkippedOIOUBLFiles)' OIOUBL file(s), already exists in destination folder`n`nCheck for duplicates or contact sproom to verify they have connection to SFTP`n`nShow log details?",'EAN Mover','YesNo','warning','Button1','ServiceNotification')
            if ($ShowSkipped -eq 'Yes') {
                Open-NgLogFile Duplicates
            }
        }
    }
    switch ($Results) {
        {$_.FailedFiles -eq 0} { $Results.Status = "Completed" }
        {$_.SkippedOIOUBLFiles -gt 0} { $Results.Status = "Duplicates" }
        {$_.FailedFiles -ne 0} { $Results.Status = "Failed" }

    }
    Show-NgNotification -ToastTitle "Results - $($Results.Status)" -ToastText "Success: $($Results.MovedFiles)`nOther files: $($Results.NonOIOUBL)`nDuplicates: $($Results.SkippedOIOUBLFiles)`nFailed: $($Results.FailedFiles)`nTotal: $($Results.TotalFiles)"
    Clear-NgOldLog -OlderThanDays $RetainLogs
    Write-NgLogMessage Information "Results: '$($Results.TotalFiles)' files processed, '$($Results.MovedFiles)' moved, '$($Results.FailedFiles)' failed, '$($Results.SkippedOIOUBLFiles)' duplicate, '$($Results.NonOIOUBL)' non-OIOUBL files"
    Write-NgLogMessage Information "$($Results | Out-String)"
    Write-NgLogMessage Information "---------------------------------------------------------[Script - $($Results.Status)]---------------------------------------------------------"
}
catch {
    Show-NgNotification -ToastTitle "Results - Failed" -ToastText "Success: $($Results.MovedFiles)`nOther files: $($Results.NonOIOUBL)`nDuplicates: $($Results.SkippedOIOUBLFiles)`nFailed: $($Results.FailedFiles)`nTotal: $($Results.TotalFiles)"
    Write-NgLogMessage Information "Results: '$($Results.TotalFiles)' files processed, '$($Results.MovedFiles)' moved, '$($Results.FailedFiles)' failed, '$($Results.SkippedOIOUBLFiles)' duplicate, '$($Results.NonOIOUBL)' non-OIOUBL files"
    Write-NgLogMessage -Level Error -Message "Main - Unknown Error: $_"
    Write-NgLogMessage Information "--------------------------------------------------------------[Script - failed]--------------------------------------------------------------"
}