<#
.SYNOPSIS
    PST Collector Agent - Scans, collects, and removes PST files from local drives or network paths.

.DESCRIPTION
    This script is deployed to workstations by CollectorMaster.ps1 or can run standalone.
    It scans specified locations for .PST files, copies them to a central collection path
    using RoboCopy with VSS snapshots (for locked files), and can remove source files
    after successful collection.

    The script maintains state in an XML configuration file, allowing it to resume
    interrupted operations and track progress across multiple runs.

.PARAMETER Mode
    Operation mode: FIND (discover PSTs), COLLECT (copy to collection path), or REMOVE (delete source files).
    Modes run in sequence - COLLECT will auto-run FIND first if needed, REMOVE will auto-run both.

.PARAMETER JobName
    Unique identifier for this collection job. Must be consistent across all runs.
    Used for naming log files, XML configs, and collection folders.

.PARAMETER Locations
    Array of paths to scan. Use "AllLocal" to scan all local fixed drives,
    or specify UNC paths like "\\server\share".

.PARAMETER CollectPath
    UNC path where collected PST files will be stored.

.PARAMETER ConfigPath
    Local path for configuration and log files. Defaults to C:\PSTCollector.
    Note: This is forced to C:\PSTCollector to support VSS hardlinks.

.PARAMETER ForceRestart
    Wipe existing configuration and start fresh. Use with caution.

.PARAMETER NoSkipCommon
    Scan all folders including Program Files, Windows, etc. that are normally skipped.

.PARAMETER ipg
    RoboCopy interpacket gap in milliseconds. Default is 1ms to throttle network usage.

.EXAMPLE
    .\CollectorAgent.ps1 -Mode FIND -JobName "Migration2024" -Locations "AllLocal" -CollectPath "\\server\PSTs"

.EXAMPLE
    .\CollectorAgent.ps1 -Mode COLLECT -JobName "Migration2024" -Locations "\\fileserver\home$" -CollectPath "\\server\PSTs"

.NOTES
    Author: Appleoddity
    Requires: Windows 7+, PowerShell 2.0+, Administrator privileges (for VSS)
    
    Exit Codes:
        0 - Success
        1 - Error (RoboCopy already running or save failure)
#>
param(
    [Parameter(Mandatory = $True, Position = 1)]
    [string]$mode,

    [Parameter(Mandatory = $True, Position = 2)]
    [string]$jobname,
    
    [Parameter(Mandatory = $True, Position = 3)]
    [string[]]$locations,

    [Parameter(Mandatory = $True, Position = 4)]
    [string]$collectpath,

    [Parameter(Position = 5)]
    [string]$configpath = "$env:SystemDrive\PSTCollector",  #This is obsoleted. Due to the hardlinks used in the Agent, we are forcing the Agent to use C:\PSTCollector.
  
    [Parameter(Position = 6)]
    [switch]$forcerestart = $false,
    
    [Parameter(Position = 7)]
    [switch]$noskipcommon = $false,

    [Parameter(Position = 8)]
    [string]$ipg = "1"
)

# Set encoding for proper Unicode support in file paths and console output
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

#region Functions

<#
.SYNOPSIS
    Initializes or loads the XML configuration document.
.DESCRIPTION
    If an existing config exists and ForceRestart is false, loads it and updates runtime parameters.
    Otherwise, creates a new XML structure with default values.
.OUTPUTS
    System.Xml.XmlDocument - The configuration XML document
#>
function InitializeXML {
    $XML = $null
    if (!$forcerestart) {
        #Config file exists?
        if (Test-Path "$configpath\$jobname.xml") {
            #Load it using XmlDocument.Load() for proper encoding handling
            $XML = New-Object System.Xml.XmlDocument
            $XML.Load("$configpath\$jobname.xml")
                
            try {
                #Update our settings on this run
                $XML.SelectSingleNode("//Configuration/Parameters/Parameter[@JobName]").SetAttribute("JobName", $JobName) #Should never be different. But whatever.
                $XML.Configuration.SetAttribute("LastRunTime", (Get-Date).tostring("G"))
                $XML.SelectSingleNode("//Configuration/Parameters/Parameter[@Mode]").SetAttribute("Mode", $mode)
                $XML.SelectSingleNode("//Configuration/Parameters/Parameter[@ConfigPath]").SetAttribute("ConfigPath", $ConfigPath)
                $XML.SelectSingleNode("//Configuration/Parameters/Parameter[@CollectPath]").SetAttribute("CollectPath", $CollectPath)
                $XML.SelectSingleNode("//Configuration/Parameters/Parameter[@ForceRestart]").SetAttribute("ForceRestart", $ForceRestart)
                return $XML #We're all done here
            } catch {
                #If we get here then we failed to load something from the config properly. Dropping through will recreate the XML and start over.
                TeeLog -message "Failure loading existing configuration XML at '$ConfigPath' - ignoring." -Logfile $log
            }
        }
    }

    #Initialize XML Nodes
    TeeLog -message "Building initial config" -Logfile $log
    $XML = New-Object System.Xml.XmlDocument
    $XMLRoot = $XML.AppendChild($XML.CreateElement("Configuration"))
    $XMLParameters = $XMLRoot.AppendChild($XML.CreateElement("Parameters"))
    $XMLErrors = $XMLRoot.AppendChild($XML.CreateElement("Errors"))
    $XMLLocations = $XMLRoot.AppendChild($XML.CreateElement("Locations"))
        
    #Store run parameters
    $XMLRoot.SetAttribute("ComputerName", "$Env:ComputerName")
    $XMLRoot.SetAttribute("Status", "Incomplete")
    $XMLRoot.SetAttribute("Description", "This is the PST CollectorAgent configuration file.")
    $XMLRoot.SetAttribute("LastRunTime", (Get-Date).tostring("G"))
    $XMLParameter = $XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("Mode", $Mode)
    $XMLParameter = $XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("ConfigPath", $ConfigPath)
    $XMLParameter = $XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("CollectPath", $CollectPath)
    $XMLParameter = $XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("JobName", $JobName)
    $XMLParameter = $XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("ForceRestart", $ForceRestart)
    $XMLParameter = $XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("IPG", $ipg)

    return $XML
}

<#
.SYNOPSIS
    Recursively scans a folder path for PST files.
.DESCRIPTION
    Traverses directory tree, skipping reparse points (junctions/symlinks) and
    optionally common system folders. Returns FileInfo objects for all .pst files found.
.PARAMETER path
    The folder path to scan recursively.
.OUTPUTS
    System.IO.FileInfo[] - Array of PST file objects found
#>
function RecurseFolder($path) {

    if ($path -eq "") { return } 
    $files = @()
    $filesTemp = @()
    $err = $null
    
    if (!$noskipcommon) {
        if ($path -eq ${env:ProgramFiles(x86)}) { return }
        if ($path -eq $Env:ProgramFiles) { return }
        if ($path -eq $Env:WinDir) { return }
        #if ($path -eq $Env:ProgramData) { return } #Seems to be questionable if valid PST files would EVER be here
        if ($path.EndsWith("System Volume Information", "CurrentCultureIgnoreCase")) { return }
        if ($path.EndsWith("`$Recycle.bin", "CurrentCultureIgnoreCase")) { return }
    }

    $directory = @(Get-ChildItem -LiteralPath $path -Force -ErrorAction SilentlyContinue -ErrorVariable err | Select-Object FullName, Attributes | Where-Object { $_.Attributes -like "*directory*" -and $_.Attributes -notlike "*reparsepoint*" })
    if ($err) {
        TeeLog -Message "RecurseFolder: ($path): $err" -Logfile $log -Errors $XMLErrors
        return
    }

    foreach ($folder in $directory) { $files += @(RecurseFolder($folder.FullName)) } #Recursive folder lookup

    $filesTemp = @(Get-ChildItem -LiteralPath $path -Filter "*.pst" -Force -ErrorAction SilentlyContinue | Where-Object { $_.Attributes -notlike "*directory*" -and $_.Attributes -notlike "*reparsepoint*" })
    
    foreach ($file in $filesTemp) {
        if (!$File.Fullname) { continue } #Fixing PS 2.0
        Teelog -message "File Found: $($file.FullName)" -Logfile $log
    }
    
    $files += $filesTemp
    $files
}

<#
.SYNOPSIS
    Writes a message to the log file, console, and optionally XML errors node.
.DESCRIPTION
    Timestamps messages in UTC format and writes to multiple outputs simultaneously.
    Used for consistent logging throughout the script.
.PARAMETER message
    The message text to log.
.PARAMETER logfile
    Path to the log file.
.PARAMETER errors
    Optional XML element to append error nodes to for tracking in config.
#>
function TeeLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [string]$message,
        [string]$logfile,
        [Parameter()]
        [System.Xml.XMLElement]$errors
    )
    
    if (!$message) { return }    
    try {
        if ($errors) {
            $XMLError = $errors.AppendChild($XML.CreateElement("Error"))
            $XMLError.SetAttribute("Message", $message)
            $XMLError.SetAttribute("Time", (Get-Date).tostring("G"))
        }
        $message = (Get-Date).ToUniversalTime().tostring("G") + ": " + $Env:ComputerName + ": " + $message 
        $message | Out-File -FilePath $logfile -Append
        Write-Host $message
    } catch { }
}

<#
.SYNOPSIS
    Creates or retrieves a VSS (Volume Shadow Copy) snapshot for a volume.
.DESCRIPTION
    Enables copying of locked/in-use PST files by accessing them through a VSS snapshot.
    Creates a directory junction (hardlink) from C:\PSTCollector\{SnapshotID} to the
    snapshot device path. Reuses existing snapshots when possible.
.PARAMETER Volume
    Drive letter with colon (e.g., "C:") to snapshot.
.PARAMETER Location
    XML element for the current location, used to store/retrieve snapshot metadata.
.OUTPUTS
    Boolean - True if snapshot is ready for use, False on failure
#>
function SnapShotVolume ([string]$Volume, $Location) {
    $Volume = $Volume.TrimEnd("\").ToUpper()
    $s1 = $null
    $s2 = $null
    #Get "vss" service startup type
    $VssStartMode = (Get-WmiObject -Query "Select StartMode From Win32_Service Where Name='vss'").StartMode
    if ($VssStartMode -eq "Disabled") {
        TeeLog -Message "SnapShotVolume: Changing Service 'VSS' startup type from 'Disabled' to 'Manual.'" -Logfile $log 
        Set-Service vss -StartUpType Manual
    }

    #Get "vss" Service status and start it if not running
    #While ((Get-Service vss).status -ne "Running")
    #{
    #    TeeLog -Message "SnapShotVolume: Waiting for Service 'VSS' to start." -Logfile $log
    #    if ($count -ge 5)
    #    {
    #        TeeLog -Message "SnapShotVolume: Failure starting Service 'VSS'" -Logfile $log -Errors $XMLErrors
    #        $Global:ErrorFlag = $true    
    #        exit
    #    }
    #    [int]$Count++ | Out-Null
    #    Start-Service vss
    #    Sleep -seconds 2
    #}

    #Test existing snapshot
    $XMLSnapshot = $Location.SelectSingleNode("Snapshot[@Volume='$volume']")
    if ($XMLsnapshot) { #Try to get reference to existing snapshot
        try {
            $s2 = Get-WmiObject Win32_ShadowCopy | Where-Object { $_.ID -eq $XMLSnapshot.ID }
            if (!$s2) { throw "No Object Returned" }
            Get-ChildItem "$ConfigPath\$($s2.ID)" -ErrorAction Stop | Out-Null
        } catch {
            TeeLog -Message "SnapShotVolume: Error using snapshot ID '$($XMLSnapshot.ID)' for volume '$volume': $($_.Exception.Message)" -Logfile $log
            try { $Location.RemoveChild($Location.SelectSingleNode("Snapshot[@Volume='$Volume']")) | Out-Null } catch {}
            if ($XMLSnapshot.ID) { & 'cmd.exe' /c rmdir "$ConfigPath\$($XMLSnapshot.ID)" | Out-Null }
            $s2 = $null
        }
    }
    
    if (!$s2) { #If no handle exists, then make a new one
        TeeLog -Message "SnapShotVolume: No VSS Snapshot found for volume '$volume'" -Logfile $log
        
        try {
            TeeLog -Message "SnapShotVolume: Creating VSS snapshot for volume '$volume'" -Logfile $log
            $s1 = (Get-WmiObject -List Win32_ShadowCopy).Create("$Volume\", "ClientAccessible")
            if ($s1.ReturnValue -gt 0) { throw "Error calling Win32_ShadowCopy.Create - ReturnValue = $($s1.ReturnValue)" }
            $s2 = Get-WmiObject Win32_ShadowCopy | Where-Object { $_.ID -eq $s1.ShadowID }
            $XMLSnapshot = $Location.AppendChild($XML.CreateElement("Snapshot"))
            $XMLSnapshot.SetAttribute("Volume", $volume)
            $XMLSnapshot.SetAttribute("ID", $s2.ID)
            TeeLog -Message "SnapShotVolume: Snapshot for volume '$volume' created. ID = '$($s2.ID)'" -Logfile $log
        } catch {
            TeeLog -Message "SnapShotVolume: Error creating VSS snapshot for volume '$volume': $($_.Exception.message)" -Logfile $log -Errors $XMLErrors
            return $false
        }
    }
    
    #Create hardlink
    try { Get-ChildItem "$ConfigPath\$($s2.ID)" -ErrorAction Stop | Out-Null }
    catch { & "cmd.exe" /c mklink /d "$Configpath\$($s2.ID)" "$($s2.DeviceObject)\" | Out-Null }

    #Test hardlink
    try { Get-ChildItem "$ConfigPath\$($s2.ID)" -ErrorAction Stop | Out-Null }
    catch {
        TeeLog -Message "SnapShotVolume: Failed to create valid hardlink to snapshot with ID '$($s2.ID)' for Volume '$volume'" -Logfile $log -Errors $XMLErrors
        try { $Location.RemoveChild($Location.SelectSingleNode("Snapshot[@ID='$($s2.ID)']")) | Out-Null } catch {}
        if ($XMLSnapshot.ID) { & 'cmd.exe' /c rmdir "$ConfigPath\$($s2.ID)" | Out-Null }
        return $false
    }

    return $true
}

<#
.SYNOPSIS
    Transfers a single file using RoboCopy with verification.
.DESCRIPTION
    Copies a file from source to destination using RoboCopy with:
    - /Z: Restartable mode for large files
    - /DCOPY:T: Copy directory timestamps
    - /COPY:DAT: Copy Data, Attributes, Timestamps
    - /IPG: Interpacket gap for bandwidth throttling
    - /UNILOG: Unicode log output for international filenames
    
    After copy, verifies file size and timestamp match.
.PARAMETER spath
    Source directory path.
.PARAMETER fname
    Filename to copy.
.PARAMETER dpath
    Destination directory path (created if needed by RoboCopy).
.PARAMETER ipg
    Interpacket gap in milliseconds.
.PARAMETER log
    Path to log file.
.PARAMETER errors
    XML element for error tracking.
.OUTPUTS
    Boolean - True if copy and verification succeeded
#>
function TrxFile($spath, $fname, $dpath, $ipg, $log, $errors) {
    # Use /UNILOG for proper Unicode path/filename support in RoboCopy output
    $tempLog = [System.IO.Path]::GetTempFileName()
    
    # Run robocopy directly with /UNILOG to capture Unicode output to file
    # This avoids console encoding issues and works with PS remoting
    TeeLog -Message "Starting RoboCopy with arguments: `"$($spath)`" `"$($dpath)`" `"$($fname)`" /Z /DCOPY:T /COPY:DAT /IPG:$ipg /R:5 /W:30 /NP /UNILOG:`"$($tempLog)`"" -Logfile $log
    & "RoboCopy.exe" "$spath" "$dpath" "$fname" /Z /DCOPY:T /COPY:DAT /IPG:$ipg /R:5 /W:30 /NP /UNILOG:"$tempLog" | Out-Null
    
    $results = Get-Content -LiteralPath $tempLog -Encoding Unicode -Raw -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $tempLog -ErrorAction SilentlyContinue
    
    TeeLog -Message "RoboCopy Results: $results" -Logfile $log
    TeeLog -Message "Verifying file: '$dpath\$fname'" -Logfile $log
    try {
        $sfile = Get-Item -LiteralPath "$spath\$fname" -ErrorAction Stop
        $dfile = Get-Item -LiteralPath "$dpath\$fname" -ErrorAction Stop
        if (Compare-Object $sfile $dfile -Property Length, LastWriteTime) {
            throw "Source and Destination files do not match."
        }
        TeeLog -Message "Verify success: Source and Destination files are a match." -Logfile $log
    } catch {
        TeeLog -Message "Verify failed: '$dpath\$fname' - $($_.Exception.Message)" -Logfile $log -Errors $errors
        return $false
    }

    return $true
}

<#
.SYNOPSIS
    Scans all configured locations for PST files.
.DESCRIPTION
    Iterates through each location, recursively scanning for .pst files.
    Records each found file in the XML config with metadata (path, owner, size, timestamp).
    Skips Outlook virtual stores (SharePoint Lists, Internet Calendar Subscriptions).
.OUTPUTS
    Boolean - True if all locations scanned successfully without errors
#>
function DoFind {
    $Complete = $true #Flag to track job status
    
    #Clear all previous locations and errors
    try { $XMLRoot.RemoveChild($XMLRoot.SelectSingleNode("Locations")) | Out-Null } catch {}
    $XMLLocations = $XMLRoot.AppendChild($XML.CreateElement("Locations"))
    try { $XMLRoot.RemoveChild($XMLRoot.SelectSingleNode("Errors")) | Out-Null } catch {}
    $XMLErrors = $XMLRoot.AppendChild($XML.CreateElement("Errors"))

    #Search each location for .PST files
    foreach ($location in $locations) {
        $Files = @()
        $location = $location.TrimEnd("\").ToLower()
        if ($location -eq "alllocal") { $location = "AllLocal" }

        $LocationComplete = $true #Flag to track location status

        $XMLLocation = $XMLLocations.AppendChild($XML.CreateElement("Location"))
        $XMLLocation.SetAttribute("Path", $location)
        $XMLLocation.SetAttribute("Status", "Incomplete")

        $XMLFiles = $XMLLocation.AppendChild($XML.CreateElement("Files"))

        if ($location -eq "AllLocal") { #Scan all local drives
            TeeLog -message "Scanning All Local Drives" -Logfile $log
            try {
                $Drives = @(get-wmiobject win32_volume -errorAction Stop | ? { $_.DriveType -eq 3 -or $_.DriveType -eq 2 } | Where-Object { $null -ne $_.DriveLetter } | % { Get-PSDrive $_.DriveLetter[0] -ErrorAction Stop })
                Teelog -message "DriveCount = $($Drives.count)" -Logfile $log

                foreach ($drive in $drives) {
                    Teelog "Scanning Drive: $($Drive.Root)" -Logfile $log
                    $files += RecurseFolder($($Drive.Root))
                }
            } catch {
                TeeLog -Message "Error while recursing drive '$($Drive.Root)' : $($_.Exception.Message)" -LogFile $Log -Errors $XMLErrors
                $LocationComplete = $false
                $Complete = $false
            }
        } else { #Scan path
            TeeLog -Message "Scanning Location: '$location'" -Logfile $log
            try {
                if (!(Test-Path $location)) { throw "Path Not Found" }
                $files += RecurseFolder($location)
            } catch {
                TeeLog -Message "Error while recursing location '$location' : $($_.Exception.Message)" -LogFile $Log -Errors $XMLErrors
                $LocationComplete = $false
                $Complete = $false
            }
        }

        #Write each file to XML
        foreach ($file in $files) {
            if (!$file.FullName) { continue } #Constantly fixing Powershell 2.0 issues
            
            #Skip special .PSTs (Outlook virtual stores)
            if ($file.FullName -like "*sharepoint lists*") { continue }
            if ($file.FullName -like "*internet calendar subscriptions*") { continue }
            
            try {
                $XMLFile = $XMLFiles.AppendChild($XML.CreateElement("File"))
                $XMLFile.SetAttribute("Path", $file.FullName)
                $XMLFile.SetAttribute("Owner", $(get-acl $file.FullName).Owner)
                $XMLFile.SetAttribute("LastModified", $file.LastWriteTime)
                $XMLFile.SetAttribute("FileSize", $file.Length)
                $XMLFile.SetAttribute("LastProcessed", (Get-Date).tostring("G")) 
                $XMLFile.SetAttribute("Status", "Found")
            } catch {
                TeeLog -Message "Error while processing files, file='$($file.FullName)' : $($_.Exception.Message)" -LogFile $Log -Errors $XMLErrors
                $LocationComplete = $false
                $Complete = $false
            }
        }
        if ($LocationComplete)
        { $XMLLocation.SetAttribute("Status", "Found") }
        else
        { $XMLLocation.SetAttribute("Status", "FindError") }
    }

    return $Complete
}

<#
.SYNOPSIS
    Copies discovered PST files to the collection path.
.DESCRIPTION
    For each PST file with status "Found", copies it to the CollectPath using RoboCopy.
    For local files, creates a VSS snapshot first to handle locked files.
    For UNC paths, copies directly from the source.
    Verifies each copy by comparing file size and timestamp.
    Updates file status to "Collected" or "CollectError" in XML config.
.OUTPUTS
    Boolean - True if all files collected successfully
#>
function DoCollect {
    #Go through each location and transfer files to collectpath as defined in XML

    $Complete = $true #Flag to track job status

    foreach ($location in $locations) {
        $location = $location.trimend("\").ToLower()
        if ($location -eq "alllocal") { $location = "AllLocal" }
        TeeLog -Message "Collecting location '$location'" -Logfile $log
        $XMLLocation = $XMLLocations.SelectSingleNode("Location[@Path='$location']")
        if ($XMLLocation.Status -eq "Found" -or $XMLLocation.Status -eq "CollectError") {
            #Collect location
            $LocationComplete = $true #Flag to track status of location

            $XMLFiles = $XMLLocation.SelectSingleNode("Files")
            foreach ($file in $XMLFiles.SelectNodes("*")) {
                #Skip if processed already
                if ($file.status -eq "Collected" -or $file.status -eq "Removed" -or $file.status -eq "RemoveError" -or $file.status -eq "Void") {
                    TeeLog -Message "File '$($File.Path)' has status '$($file.Status)' - Skipping" -Logfile $log
                } else {
                    $spath = $null
                    $dpath = $null

                    #Collect file
                    TeeLog -Message "Collecting file: '$($File.path)'" -Logfile $log
                    $file.SetAttribute("LastProcessed", (Get-Date).tostring("G")) 
                    $fname = Split-Path -Path $file.path -Leaf

                    if ($file.path.StartsWith("\\")) {
                        #Build UNC path
                        $spath = Split-Path -Path $file.path -Parent
                        $dpath = "$CollectPath\$jobname\" + $(Split-Path -Path $file.path -Parent).ToLower().Replace("$location\", "").TrimEnd("\")
                    } else {
                        #Build local path and Snapshot volume
                        if (SnapShotVolume -Volume $(Split-Path -Path $file.path -Qualifier) -Location $XMLLocation) {
                            $XMLLocation = $XMLLocations.SelectSingleNode("Location[@Path='$location']")
                            $XMLSnapshot = $XMLLocation.SelectSingleNode("Snapshot[@Volume='$(Split-Path -Path $file.path -Qualifier)']")
                            $spath = "$ConfigPath\$($XMLSnapshot.ID)" + $(Split-Path -Path $(Split-Path -Path $file.path -NoQualifier) -Parent)
                            $dpath = "$CollectPath\$jobname\" + $(Split-Path -Path $file.path -Parent).Replace(":", "").TrimEnd("\")
                        } else {
                            #Shadowcopy failure
                            $XMLLocation = $XMLLocations.SelectSingleNode("Location[@Path='$location']")
                            TeeLog -Message "$($file.path) - ShadowCopy failure - Collecting without VSS." -Logfile $log
                            $spath = Split-Path -Path $file.path -Parent
                            $dpath = "$CollectPath\$jobname\" + $(Split-Path -Path $file.path -Parent).Replace(":", "").TrimEnd("\")
                        }    
                    }

                    if ($spath -and $dpath) {
                        #Save data
                        $XML.Save("$ConfigPath\$jobname.xml")
                    
                        TeeLog -Message "Transferring file: '$($File.path)'" -Logfile $log
                        if (TrxFile -spath $spath -dpath $dpath -fname $fname -ipg $ipg -log $log -errors $XMLErrors) {
                            TeeLog -Message "Transfer successful: $($File.path)" -Logfile $log
                            $file.SetAttribute("Status", "Collected")
                        } else {
                            TeeLog -Message "Transfer failed: $($File.path)" -Logfile $log
                            $file.SetAttribute("Status", "CollectError")
                            $Complete = $false
                            $LocationComplete = $false
                        }
                    } else {
                        TeeLog -Message "Unable to transfer file: $($File.path) - Unknown error - Does the drive letter still exist?" -Logfile $log
                        $file.SetAttribute("Status", "CollectError")
                        $Complete = $false
                        $LocationComplete = $false
                    }

                }
            }

            #Update location status and save data
            if ($LocationComplete) { $XMLLocation.SetAttribute("Status", "Collected") }
            else { $XMLLocation.SetAttribute("Status", "CollectError") }
            $XML.Save("$ConfigPath\$jobname.xml")
        } else {
            if ($XMLLocation.Status -eq "Collected")
            { TeeLog -Message "The location '$location' has status '$($XMLLocation.Status)' - Skipping." -Logfile $log }
            else {
                TeeLog -Message "The location '$location' has status '$($XMLLocation.Status)' which is other than FOUND, COLLECTED or COLLECTERROR - Skipping." -Logfile $log -Errors $XMLErrors
                $Complete = $false
            }
        }
    }
    
    return $Complete
}

<#
.SYNOPSIS
    Removes source PST files after successful collection.
.DESCRIPTION
    For each PST file with status "Collected", deletes the original source file.
    Skips whitelisted files (SharePoint Lists, Internet Calendar Subscriptions).
    Updates file status to "Removed" or "RemoveError" in XML config.
    
    WARNING: Only run after verifying PST files are successfully imported to Exchange Online
    and users have detached PSTs from Outlook using unloadPST.vbs.
.OUTPUTS
    Boolean - True if all files removed successfully
#>
function DoRemove {
    #Go through each location and remove files as defined in XML

    $Complete = $true #Flag to track job status

    foreach ($location in $locations) {
        $location = $location.trimend("\").ToLower()
        if ($location -eq "alllocal") { $location = "AllLocal" }
        TeeLog -Message "Removing location '$location'" -Logfile $log
        $XMLLocation = $XMLLocations.SelectSingleNode("Location[@Path='$location']")
        if ($XMLLocation.Status -eq "Collected" -or $XMLLocation.Status -eq "RemoveError") {
            #Remove location
            $LocationComplete = $true #Flag to track status of location

            $XMLFiles = $XMLLocation.SelectSingleNode("Files")
            foreach ($file in $XMLFiles.SelectNodes("*")) {
                #Skip if processed already
                if ($file.status -eq "Removed" -or $file.status -eq "Void") {
                    TeeLog -Message "File '$($File.Path)' has status '$($file.Status)' - Skipping" -Logfile $log
                } else {
                    #Remove file
                    TeeLog -Message "Removing file: '$($File.path)'" -Logfile $log
                    $file.SetAttribute("LastProcessed", (Get-Date).tostring("G")) 

                    try {
                        $fname = $(Split-Path -Path $file.path -Leaf).ToLower()
                        #Skip whitelist files - for backwards compatibility
                        if ($fname -like "*sharepoint lists*" -or $fname -like "*internet calendar subscriptions*") {
                            TeeLog -Message "File Whitelisted - Setting status and skipping: $($File.path)" -Logfile $log
                        } else {
                            Remove-Item -LiteralPath $file.path -Force -ErrorAction Stop
                            TeeLog -Message "Remove successful: $($File.path)" -Logfile $log
                        }
                        $file.SetAttribute("Status", "Removed")
                    } catch {
                        if ($_.Exception.Message -like "*cannot find path*") {
                            TeeLog -Message "Remove successful: File no longer exists - $($File.path)" -Logfile $log
                            $file.SetAttribute("Status", "Removed")
                        } else {
                            TeeLog -Message "Remove Failed: $($File.path) - $($_.Exception.Message)" -Logfile $log -Errors $XMLErrors
                            $file.SetAttribute("Status", "RemoveError")
                            $Complete = $false
                            $LocationComplete = $false
                        }
                    }
                }
            }

            #Update location status and save data
            if ($LocationComplete) { $XMLLocation.SetAttribute("Status", "Removed") }
            else { $XMLLocation.SetAttribute("Status", "RemoveError") }
            $XML.Save("$ConfigPath\$jobname.xml")
        } else {
            if ($XMLLocation.Status -eq "Removed")
            { TeeLog -Message "The location '$location' has status '$($XMLLocation.Status)' - Skipping." -Logfile $log }
            else {
                TeeLog -Message "The location '$location' has status '$($XMLLocation.Status)' which is other than COLLECTED, REMOVED or REMOVEERROR - Skipping." -Logfile $log -Errors $XMLErrors
                $Complete = $false
            }
        }
    }
    
    return $Complete
    
}

#endregion Functions

#region Main Script
# ==================================================================================
# MAIN SCRIPT EXECUTION
# ==================================================================================

# Force config path to C:\PSTCollector to support VSS hardlinks (directory junctions)
$configpath = "C:\PSTCollector"
$collectpath = $collectpath.trimend("\")
$log = $configpath + "\$jobname.log"

Teelog -message "Running As: $Env:Username" -Logfile $log

# Check for existing RoboCopy process to prevent duplicate runs
# (Can happen when system resumes from standby/hibernation mid-collection)
if (Get-Process | Where-Object { $_.Name -eq "robocopy" }) {
    Teelog -message "Found process 'Robocopy' already running. This script may already be active - Quitting." -Logfile $log
    exit 1
}

# Create config path if not exist
New-Item -ItemType Directory -Path $configpath -ErrorAction SilentlyContinue | Out-Null

# Initialize or load XML configuration
$XML = InitializeXML
$XMLRoot = $XML.SelectSingleNode("//Configuration")
$XMLErrors = $XML.SelectSingleNode("//Configuration/Errors")
$XMLParameters = $XML.SelectSingleNode("//Configuration/Parameters")
$XMLLocations = $XML.SelectSingleNode("//Configuration/Locations")

# Execute requested mode
switch ($mode) {
    FIND {
        #FIND Mode

        if ($XMLRoot.Status -eq "Found" -or $XMLRoot.Status -eq "Collected" -or $XMLRoot.Status -eq "Removed" -or $XMLRoot.Status -eq "RemoveError" -or $XMLRoot.Status -eq "CollectError") {
            TeeLog -message "This job has status '$($XMLRoot.Status)' - Quitting." -Logfile $log
        } else {
            TeeLog -message "Starting FIND job." -Logfile $log
            if (DoFind) {
                TeeLog -message "The FIND operation completed successfully." -Logfile $log
                $XMLRoot.SetAttribute("Status", "Found")
            } else {
                TeeLog -message "The FIND operation failed." -Logfile $log
                $XMLRoot.SetAttribute("Status", "FindError")
            }
        }
    }

    COLLECT {
        if ($XMLRoot.Status -eq "FindError" -or $XMLRoot.Status -eq "Incomplete") {
            TeeLog -message "This job has status '$($XMLRoot.Status)' - Performing FIND first." -Logfile $log
            if (DoFind) {
                TeeLog -message "The FIND operation completed successfully." -Logfile $log
                $XMLRoot.SetAttribute("Status", "Found")
            } else {
                TeeLog -message "The FIND operation failed." -Logfile $log
                $XMLRoot.SetAttribute("Status", "FindError")
            }
            $XML.Save("$ConfigPath\$jobname.xml")
            $XMLLocations = $XML.SelectSingleNode("//Configuration/Locations")
        }

        if ($XMLRoot.Status -eq "Found" -or $XMLRoot.Status -eq "CollectError") {
            TeeLog -message "Starting COLLECT job." -Logfile $log

            if (DoCollect) {
                TeeLog -message "COLLECT job finished successfully." -Logfile $log
                $XMLRoot.SetAttribute("Status", "Collected")
            } else {
                TeeLog -message "COLLECT job failed." -Logfile $log
                $XMLRoot.SetAttribute("Status", "CollectError")
            }
        } else {
            if ($XMLRoot.Status -eq "Collected" -or $XMLRoot.Status -eq "Removed" -or $XMLRoot.Status -eq "RemoveError" -or $XMLRoot.Status -eq "FindError")
            { TeeLog -message "This job has status '$($XMLRoot.Status)' - Quitting." -Logfile $log }
            else {
                TeeLog -message "This job has an unknown status '$($XMLRoot.Status)' - Quitting." -Logfile $log
            }
        }
    }
    REMOVE {
        if ($XMLRoot.Status -eq "FindError" -or $XMLRoot.Status -eq "Incomplete") {
            TeeLog -message "This job has status '$($XMLRoot.Status)' - Performing FIND first." -Logfile $log
            if (DoFind) {
                TeeLog -message "The FIND operation completed successfully." -Logfile $log
                $XMLRoot.SetAttribute("Status", "Found")
            } else {
                TeeLog -message "The FIND operation failed." -Logfile $log
                $XMLRoot.SetAttribute("Status", "FindError")
            }
            $XML.Save("$ConfigPath\$jobname.xml")
            $XMLLocations = $XML.SelectSingleNode("//Configuration/Locations")
        }

        if ($XMLRoot.Status -eq "CollectError" -or $XMLRoot.Status -eq "Found") {
            TeeLog -message "This job has status '$($XMLRoot.Status)' - Performing COLLECT first." -Logfile $log
            if (DoCollect) {
                TeeLog -message "The COLLECT operation completed successfully." -Logfile $log
                $XMLRoot.SetAttribute("Status", "Collected")
            } else {
                TeeLog -message "The COLLECT operation failed." -Logfile $log
                $XMLRoot.SetAttribute("Status", "CollectError")
            }
            $XML.Save("$ConfigPath\$jobname.xml")
            $XMLLocations = $XML.SelectSingleNode("//Configuration/Locations")
        }

        if ($XMLRoot.Status -eq "Collected" -or $XMLRoot.Status -eq "RemoveError") {
            TeeLog -message "Starting REMOVE job." -Logfile $log

            if (DoRemove) {
                TeeLog -message "REMOVE job finished successfully." -Logfile $log
                $XMLRoot.SetAttribute("Status", "Removed")
            } else {
                TeeLog -message "REMOVE job failed." -Logfile $log
                $XMLRoot.SetAttribute("Status", "RemoveError")
            }
        } else {
            if ($XMLRoot.Status -eq "Removed")
            { TeeLog -message "This job has status '$($XMLRoot.Status)' - Quitting." -Logfile $log }
            else {
                TeeLog -message "This job has an unknown status '$($XMLRoot.Status)' - Quitting." -Logfile $log
            }
        }
    }
    default {
        TeeLog -message "This job has an unknown MODE '$mode' - Quitting." -Logfile $log
    }
}


try {
    if ($null -ne $XML) {
        #Write data to disk
        $XML.Save("$ConfigPath\$jobname.xml")
 
        #Create collect path if not exist
        New-Item -ItemType Directory -Path "$CollectPath" -ErrorAction SilentlyContinue | Out-Null
        
        #Write data to collection
        $XML.Save("$CollectPath\$jobname.xml")
        Copy-Item "$Configpath\$jobname.log" "$collectpath" -ErrorAction Stop -Force | Out-Null
    }
} catch {
    TeeLog -Message "Error while saving XMLs during finalize : $($_.Exception.Message)" -LogFile $Log
    exit 1
}

TeeLog -Message "CollectorAgent - Complete" -LogFile $Log
exit 0