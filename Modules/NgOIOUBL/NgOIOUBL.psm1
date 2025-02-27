enum NgSourceType {
    Zip
    Folder
    Unknown
}

enum NgSourceStatus {
    Deleted
    Read
    Copied
    Moved
    Skipped
    Proccessed
    ProccessedWithErrors
    ProccessedWithOthers
    Failed
    Unknown
}

class NgFileSystemInfo {
    [string] $Name
    [string] $FullName
    [bool] $IsCompressed
    [string] $Extension
    [string] $Parrent
    [string] $ParrentFullName
    [string] $FullPath
    [NgSourceStatus] $Status

    NgFileSystemInfo([System.IO.Compression.ZipArchiveEntry]$Entry, [System.IO.FileInfo]$Parrent) {
        $this.Name = $Entry.Name
        $this.FullName = $Entry.FullName
        $this.IsCompressed = $true
        $this.Extension = [System.IO.Path]::GetExtension($Entry.FullName)
        $this.Parrent = $Parrent.Name
        $this.ParrentFullName = $Parrent.FullName
        $this.FullPath = Join-Path $Parrent.FullName $Entry.FullName
        $this.Status = "Read"
    }

    NgFileSystemInfo([System.IO.FileInfo] $FileInfo) {
        $this.Name = $FileInfo.Name
        $this.FullName = $FileInfo.FullName
        $this.IsCompressed = $false
        $this.Extension = $FileInfo.Extension
        $this.Parrent = (Get-Item -path $FileInfo.DirectoryName).Name
        $this.ParrentFullName = $FileInfo.DirectoryName
        $this.FullPath = $FileInfo.FullName
        $this.Status = "Read"
    }

    [bool] Exists ([string] $Path) {
        return (Test-Path -Path (Join-Path $this.Name $Path))
    }

    [void] Delete () {
        if ($this.IsCompressed) {
            $Zip = [System.IO.Compression.ZipFile]::Open($this.ParrentFullName, 'Update')
            $Entry = $Zip.GetEntry($this.FullName)
            $Entry.Delete()
            $Zip.Dispose()
            $this.Status = "Deleted"
        }
        else {
            [System.IO.File]::Delete($this.FullName)
            $this.Status = "Deleted"
        }
    }

    [void] Copy ([string] $Destination, [bool] $Overwrite) {
        if ($this.IsCompressed) {
            $Zip = [System.IO.Compression.ZipFile]::OpenRead($this.ParrentFullName)
            $Entry = $Zip.GetEntry($this.FullName)
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($Entry, (Join-Path $Destination $Entry.Name), $Overwrite)
            $Zip.Dispose()
            $this.Status = "Copied"
        }
        else {
            [System.IO.File]::Copy($this.FullName, (Join-Path $Destination $this.Name), $Overwrite)
            $this.Status = "Copied"
        }
    }

    [void] Move ([string] $Destination, [bool] $Overwrite) {
        if ($this.IsCompressed) {
            $Zip = [System.IO.Compression.ZipFile]::Open($this.ParrentFullName, 'Update')
            $Entry = $Zip.GetEntry($this.FullName)
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($Entry, (Join-Path $Destination $Entry.Name), $Overwrite)
            $Entry.Delete()
            $Zip.Dispose()
            $this.Status = "Moved"
        }
        else {
            [System.IO.File]::Move($this.FullName, (Join-Path $Destination $this.Name), $Overwrite)
            $this.Status = "Moved"
        }
    }

}

class NgOIOUBL {
    [string] $Source
    [NgSourceType] $Type
    [int] $OIOUBLCount
    [Int] $OtherCount
    [NgSourceStatus] $Status
    [PSCustomObject]$Entries = @()

    NgOIOUBL() {}

    NgOIOUBL([string] $source, [bool]$Recurse) {
        $this.Source = $source
        $this.Type = [NgOIOUBL]::GetSourceType($source)

        $this.Entries = [NgOIOUBL]::GetOIOUBLXmlFiles($source, $Recurse)
        $this.OIOUBLCount = $this.Entries.OIOUBL.Count
        $this.OtherCount = $this.Entries.Other.Count

        if (!($this.Entries) -or $this.Entries -eq "Failed"){$this.Status = "Failed"}
        else {$this.Status = "Read"}
    }

    NgOIOUBL([System.IO.FileSystemInfo] $source) {
        $this.Source = $source.FullName
        $this.Type = [NgOIOUBL]::GetSourceType($source.FullName)
        
        $this.Entries = [NgOIOUBL]::GetOIOUBLFromZIP($source)
        $this.OIOUBLCount = $this.Entries.OIOUBL.Count
        $this.OtherCount = $this.Entries.Other.Count
        
        if (!($this.Entries) -or $this.Entries -eq "Failed"){$this.Status = "Failed"}
        else {$this.Status = "Read"}
    }

    [NgOIOUBL] GetOtherZipFiles () {
        if ($this.OIOUBLCount -eq 0) {
            return $this
        }
        return $null
        
    }

    [System.Collections.ArrayList] GetUniqueEntries([string] $Path) {
        $res = @()
        $Existing = (Get-ChildItem -Path $Path -File).name
        $this.Entries.OIOUBL | ForEach-Object {
            if ($_.name -notin $Existing) {
                $res += $_ 
            }
        }
        return $res

    }

    [System.Collections.ArrayList] GetDuplicatesEntries([string] $Path) {
        $res = @()
        $Existing = (Get-ChildItem -Path $Path).name
        $this.Entries.OIOUBL | ForEach-Object {
            if ($_.name -in $Existing) {
                $res += $_
            }
        }
        return $res
    }

    [NgOIOUBL] GetUniqueSource ([string] $Path) {
        if ($this.OIOUBLCount -eq 0) {
            return $null
        }

        $res = [NgOIOUBL]::new()

        $UniqueEntries = $this.GetUniqueEntries($Path)

        $UniqueEntries = $this.GetUniqueEntries($Path)
        if ($UniqueEntries.count -eq 0) {
            return $null
        }
        elseif ($UniqueEntries.count -lt $this.OIOUBLCount) {
            $clone = $this
            $clone.OIOUBLCount = $UniqueEntries.count
            $clone.Entries = $UniqueEntries
            $res = $clone
        }
        elseif ($UniqueEntries.count -eq $this.OIOUBLCount) {
            $res = $this
        }
        return $res
    }

    [void] ExtractUniqueToDirectory ([string] $Path) {
        if ($this.Type -ne "Zip"){
            return
        }
        $UniqueEntries = $this.GetUniqueEntries($Path)
        $Zip = [System.IO.Compression.ZipFile]::OpenRead($this.Source)


        $Zip.Entries | ForEach-Object {
            if ($_.FullName -notin $UniqueEntries.FullName) {
                return
            }
            else {
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($_, (Join-Path $Path $_.Name))
            }
        }
    
        $Zip.Dispose()

    }

    [void] ExtractAllToDirectory ([string] $Path) {
        if ($this.Type -ne "Zip"){
            return
        }
        [System.IO.Compression.ZipFile]::ExtractToDirectory($this.Source, $Path)   
    }

    [PSCustomObject] BulkCopyEntries ([System.Collections.ArrayList]$Entries, [string] $Destination, [bool] $Overwrite) {
        $res = [PSCustomObject]@{
            Success = @()
            Skipped = @()
            Failed = @()
        }
        if ($this.Type -eq "Zip") {
            $Zip = [System.IO.Compression.ZipFile]::OpenRead($this.Source)
            foreach ($Entry in $Entries) {
                try {
                    if ($Overwrite){
                        $ZipEntry = $Zip.GetEntry($Entry.FullName)
                        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($ZipEntry, (Join-Path $Destination $ZipEntry.Name), $Overwrite)
                        $Entry.Status = "Copied"
                        $res.Success += $Entry
                    }
                    else {
                        if (!($Entry.Exists($Destination))) {
                            $ZipEntry = $Zip.GetEntry($Entry.FullName)
                            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($ZipEntry, (Join-Path $Destination $ZipEntry.Name), $Overwrite)
                            $Entry.Status = "Copied"
                            $res.Success += $Entry
                        }
                        else {
                            $Entry.Status = "Skipped"
                            $res.Skipped += $Entry
                        }
                    }
                }
                catch {
                    $Entry.Status = "Failed"
                    $res.Failed += $Entry
                }
            }
            $Zip.Dispose()


        }
        elseif ($this.Type -eq "Folder") {
            foreach ($Entry in $Entries) {
                try {
                    if ($Overwrite){
                        $Entry.Copy($Destination, $Overwrite)
                        $Entry.Status = "Copied"
                        $res.Success += $Entry
                    }
                    else {
                        if (!($Entry.Exists($Destination))) {
                            $Entry.Copy($Destination, $Overwrite)
                            $Entry.Status = "Copied"
                            $res.Success += $Entry
                        }
                        else {
                            $Entry.Status = "Skipped"
                            $res.Skipped += $Entry
                        }
                    }
                }
                catch {
                    $Entry.Status = "Failed"
                    $res.Failed += $Entry
                }
            }
        }
        return $res
    }

    [PSCustomObject] BulkDeleteEntries ([System.Collections.ArrayList]$Entries) {
        $res = [PSCustomObject]@{
            Success = @()
            Skipped = @()
            Failed = @()
        }
        if ($this.Type -eq "Zip") {
            $Zip = [System.IO.Compression.ZipFile]::Open($this.Source, 'Update')
            foreach ($Entry in $Entries) {
                try {
                    $ZipEntry = $Zip.GetEntry($Entry.FullName)
                    $ZipEntry.Delete()
                    $Entry.Status = "Deleted"
                    $res.Success += $Entry
                }
                catch {
                    $Entry.Status = "Failed"
                    $res.Failed += $Entry
                }
            }
            $Zip.Dispose()
        }
        elseif ($this.Type -eq "Folder") {
            foreach ($Entry in $Entries) {
                try {
                    $Entry.Delete()
                    $Entry.Status = "Deleted"
                    $res.Success += $Entry
                }
                catch {
                    $Entry.Status = "Failed"
                    $res.Failed += $Entry
                }
            }
        }
        return $res
    }

    [void] Delete (){
        if ($this.Type -eq "Zip") {
            [System.IO.File]::Delete($this.Source)
            $this.Status = "Deleted"
        }
        elseif ($this.Type -eq "Folder") {
            $this.Entries.OIOUBL | ForEach-Object {
                [System.IO.File]::Delete($_.FullName)
            }
            $this.Status = "Deleted"
        }
    }

    static [PSCustomObject] GetOIOUBLXmlFiles([string]$source, [bool]$Recurse) {
        try {
            $OIOUBLEntries = @()
            $OtherEntries = @()
            $Failed = @()

            if ($Recurse) {$Items = Get-ChildItem -Path $source -Filter "*.xml" -Recurse}
            else {$Items = Get-ChildItem -Path $source -Filter "*.xml"}

            ForEach ($Item in $items) {
                try {
                    if ([NgOIOUBL]::IsOIOUBLXml($Item.FullName)) {
                        $OIOUBLEntries += [NgFileSystemInfo]::new($Item)
                    }
                    else {
                        $OtherEntries += [NgFileSystemInfo]::new($Item)
                    }
                }
                catch {
                    $FailedEntry = [NgFileSystemInfo]::new($Item, $source)
                    $FailedEntry.Status = "Failed"
                    $Failed += $FailedEntry
                } 
            }
            return [PSCustomObject]@{
                OIOUBL = $OIOUBLEntries
                Other = $OtherEntries
                Failed = $Failed
            }
        }
        catch {
            
            return ([NgSourceStatus]"Failed")
        }

    }

    static [PSCustomObject] GetOIOUBLFromZIP($source) {
        try {
            $Zip = [System.IO.Compression.ZipFile]::OpenRead($source.FullName)
            $OIOUBLEntries = @()
            $OtherEntries = @()
            $Failed = @()
            $MetEntries = $Zip.Entries | where-object { $_.Name -like "*.xml"}

            

            ForEach ($MetEntry in $MetEntries) {
                try {
                    $stream = $MetEntry.Open()
                    $reader = New-Object IO.StreamReader($stream)
                    [xml]$text = $reader.ReadToEnd()
    
                    $reader.Close()
                    $stream.Close()
    
                    if  ([NgOIOUBL]::IsOIOUBLXml($text)) {
                        $OIOUBLEntries += [NgFileSystemInfo]::new($MetEntry, $source)
                    }
                    else {
                        $OtherEntries += [NgFileSystemInfo]::new($MetEntry, $source)
                    }
                }
                catch {
                    $FailedEntry = [NgFileSystemInfo]::new($MetEntry, $source)
                    $FailedEntry.Status = "Failed"
                    $Failed += $FailedEntry
                }
            }
            

            $Zip.Entries | where-object { $_.Name -notlike "*.xml"} | ForEach-Object {
                $OtherEntries += [NgFileSystemInfo]::new($_, $source)
            }

            $Zip.Dispose()
            return [PSCustomObject]@{
                OIOUBL = $OIOUBLEntries
                Other = $OtherEntries
                Failed = $Failed
            }
        }
        catch {
            return ([NgSourceStatus]"Failed")
        }
    }

    static [string] GetSourceType([string]$source) {
        $TypeName = Get-Item -Path $source

        if ($TypeName.PSIsContainer) {
            return "Folder"
        } elseif ($TypeName.Extension -eq ".zip") {
            return "Zip"
        } else {
            return "Unknown"
        }
    }

    static [bool] IsOIOUBLXml([string] $FullName) {
        try {
            [xml]$xmlDoc = Get-Content -Path $FullName
            return ($xmlDoc.DocumentElement.CustomizationID -match "OIOUBL")
        } catch {
            throw $_
        }
    }

    static [bool] IsOIOUBLXml([xml] $xmlDoc) {
        try {
            return ($xmlDoc.DocumentElement.CustomizationID -match "OIOUBL")
        } catch {
            throw $_
        }
    }
}