# This is the PST Collector agent script. It should run as as a startup script, or be run by the CollectorMaster
# on each system to scan for PST files.
# It will scan all provided locations for .PST files.
#
# Usage: CollectorAgent.ps1 -Mode <mode> -JobName <jobname> -Locations <locations> -CollectPath <path> [-ForceRestart] [-NoSkipCommon] [-ipg <interpacket gap throttle>]
#
Param(
    [Parameter(Mandatory=$True, Position=1)]
        [string]$mode,

    [Parameter(Mandatory=$True, Position=2)]
        [string]$jobname,
    
    [Parameter(Mandatory=$True, Position=3)]
        [string[]]$locations,

    [Parameter(Mandatory=$True, Position=4)]
        [string]$collectpath,

    [Parameter(Position=5)]
        [string]$configpath="$env:SystemDrive\PSTCollector",  #This is obsoleted. Due to the hardlinks used in the Agent, we are forcing the Agent to use C:\PSTCollector.
  
    [Parameter(Position=6)]
        [switch]$forcerestart=$false,
    
    [Parameter(Position=7)]
        [switch]$noskipcommon=$false,

    [Parameter(Position=8)]
        [string]$ipg="1"
)

function InitializeXML
{
    $XML=$null
    if (!$forcerestart)
    {
        #Config file exists?
        if (Test-Path "$configpath\$jobname.xml")
        {
            #Load it
            [System.Xml.XmlDocument]$XML = get-content "$configpath\$jobname.xml" 
                
            try
            {
                #Update our settings on this run
                $XML.SelectSingleNode("//Configuration/Parameters/Parameter[@JobName]").SetAttribute("JobName",$JobName) #Should never be different. But whatever.
                $XML.Configuration.SetAttribute("LastRunTime",(get-date).tostring("G"))
                $XML.SelectSingleNode("//Configuration/Parameters/Parameter[@Mode]").SetAttribute("Mode",$mode)
                $XML.SelectSingleNode("//Configuration/Parameters/Parameter[@ConfigPath]").SetAttribute("ConfigPath",$ConfigPath)
                $XML.SelectSingleNode("//Configuration/Parameters/Parameter[@CollectPath]").SetAttribute("CollectPath",$CollectPath)
                $XML.SelectSingleNode("//Configuration/Parameters/Parameter[@ForceRestart]").SetAttribute("ForceRestart",$ForceRestart)
                return $XML #We're all done here
            }
            catch
            {
                #If we get here then we failed to load something from the config properly. Dropping through will recreate the XML and start over.
                TeeLog -message "Failure loading existing configuration XML at '$ConfigPath' - ignoring." -Logfile $log
            }
        }
    }

    #Initialize XML Nodes
    TeeLog -message "Building initial config" -Logfile $log
    $XML=New-Object System.Xml.XmlDocument
    $XMLRoot=$XML.AppendChild($XML.CreateElement("Configuration"))
    $XMLParameters=$XMLRoot.AppendChild($XML.CreateElement("Parameters"))
    $XMLErrors=$XMLRoot.AppendChild($XML.CreateElement("Errors"))
    $XMLLocations=$XMLRoot.AppendChild($XML.CreateElement("Locations"))
        
    #Store run parameters
    $XMLRoot.SetAttribute("ComputerName", "$Env:ComputerName")
    $XMLRoot.SetAttribute("Status", "Incomplete")
    $XMLRoot.SetAttribute("Description", "This is the PST CollectorAgent configuration file.")
    $XMLRoot.SetAttribute("LastRunTime", (get-date).tostring("G"))
    $XMLParameter=$XMLParameters.AppendChild($XML.CreateElement("Parameter"))
        $XMLParameter.SetAttribute("Mode", $Mode)
    $XMLParameter=$XMLParameters.AppendChild($XML.CreateElement("Parameter"))
        $XMLParameter.SetAttribute("ConfigPath", $ConfigPath)
    $XMLParameter=$XMLParameters.AppendChild($XML.CreateElement("Parameter"))
        $XMLParameter.SetAttribute("CollectPath", $CollectPath)
    $XMLParameter=$XMLParameters.AppendChild($XML.CreateElement("Parameter"))
        $XMLParameter.SetAttribute("JobName", $JobName)
    $XMLParameter=$XMLParameters.AppendChild($XML.CreateElement("Parameter"))
        $XMLParameter.SetAttribute("ForceRestart", $ForceRestart)

    return $XML
}

function RecurseFolder($path)
{

    if ($path -eq "") { return } 
    $files=@()
    $filesTemp=@()
    $err=$null
    
    if (!$noskipcommon)
    {
        if ($path -eq ${env:ProgramFiles(x86)}) { return }
        if ($path -eq $Env:ProgramFiles) { return }
        if ($path -eq $Env:WinDir) { return }
        #if ($path -eq $Env:ProgramData) { return } #Seems to be questionable if valid PST files would EVER be here
        if ($path.EndsWith("System Volume Information", "CurrentCultureIgnoreCase")) { return }
        if ($path.EndsWith("`$Recycle.bin", "CurrentCultureIgnoreCase")) { return }
    }

    $directory = @(get-childitem -LiteralPath $path -Force -ErrorAction SilentlyContinue -ErrorVariable err | Select FullName,Attributes | Where-Object {$_.Attributes -like "*directory*" -and $_.Attributes -notlike "*reparsepoint*"})
    if ($err)
    {
        TeeLog -Message "RecurseFolder: ($path): $err" -Logfile $log -Errors $XMLErrors
        return
    }

    foreach ($folder in $directory) { $files+=@(RecurseFolder($folder.FullName)) } #Recursive folder lookup

    $filesTemp=@(get-childitem -LiteralPath $path -Filter "*.pst" -Force -ErrorAction SilentlyContinue | Where-Object {$_.Attributes -notlike "*directory*" -and $_.Attributes -notlike "*reparsepoint*"})
    
    Foreach ($file in $filesTemp)
    {
        if (!$File.Fullname) { Continue } #Fixing PS 2.0
        Teelog -message "File Found: $($file.FullName)" -Logfile $log
    }
    
    $files+=$filesTemp
    $files
}

function TeeLog
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
            [string]$message,
            [string]$logfile,
        [Parameter()]
            [System.Xml.XMLElement]$errors
    )
    
    if (!$message) { Return }    
    try
    {
        if ($errors) 
        {
            $XMLError=$errors.AppendChild($XML.CreateElement("Error"))
            $XMLError.SetAttribute("Message", $message)
            $XMLError.SetAttribute("Time", (get-date).tostring("G"))
        }
        $message = (get-date).ToUniversalTime().tostring("G") + ": " + $Env:ComputerName + ": " + $message 
        $message | out-file -FilePath $logfile -Append
        write-host $message
    }
    catch { }
}

function SnapShotVolume ([string]$Volume, $Location)
{
    $Volume = $Volume.TrimEnd("\").ToUpper()
    $s1 = $null
    $s2 = $null
    #Get "vss" service startup type
    $VssStartMode = (Get-WmiObject -Query "Select StartMode From Win32_Service Where Name='vss'").StartMode
    if ($VssStartMode -eq "Disabled")
    {
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
    if ($XMLsnapshot) #Try to get reference to existing snapshot
    {
        try
        {
            $s2 = Get-WmiObject Win32_ShadowCopy | Where-Object { $_.ID -eq $XMLSnapshot.ID }
            if (!$s2) { throw "No Object Returned" }
            get-childitem "$ConfigPath\$($s2.ID)" -ErrorAction Stop | Out-Null
        }
        catch
        {
            TeeLog -Message "SnapShotVolume: Error using snapshot ID '$($XMLSnapshot.ID)' for volume '$volume': $($_.Exception.Message)" -Logfile $log
            try { $Location.RemoveChild($Location.SelectSingleNode("Snapshot[@Volume='$Volume']")) | Out-Null } catch {}
            if ($XMLSnapshot.ID) { & 'cmd.exe' /c rmdir "$ConfigPath\$($XMLSnapshot.ID)" | Out-Null }
            $s2 = $null
        }
    }
    
    if (!$s2) #If no handle exists, then make a new one
    {
        TeeLog -Message "SnapShotVolume: No VSS Snapshot found for volume '$volume'" -Logfile $log
        
        try
        {
            TeeLog -Message "SnapShotVolume: Creating VSS snapshot for volume '$volume'" -Logfile $log
            $s1 = (Get-WmiObject -List Win32_ShadowCopy).Create("$Volume\", "ClientAccessible")
            if ($s1.ReturnValue -gt 0) { throw "Error calling Win32_ShadowCopy.Create - ReturnValue = $($s1.ReturnValue)" }
            $s2 = Get-WmiObject Win32_ShadowCopy | Where-Object { $_.ID -eq $s1.ShadowID }
            $XMLSnapshot = $Location.AppendChild($XML.CreateElement("Snapshot"))
            $XMLSnapshot.SetAttribute("Volume", $volume)
            $XMLSnapshot.SetAttribute("ID", $s2.ID)
            TeeLog -Message "SnapShotVolume: Snapshot for volume '$volume' created. ID = '$($s2.ID)'" -Logfile $log
        }
        Catch
        {
            TeeLog -Message "SnapShotVolume: Error creating VSS snapshot for volume '$volume': $($_.Exception.message)" -Logfile $log -Errors $XMLErrors
            return $false
        }
    }
    
    #Create hardlink
    try { get-childitem "$ConfigPath\$($s2.ID)" -ErrorAction Stop | Out-Null }
    Catch { & "cmd.exe" /c mklink /d "$Configpath\$($s2.ID)" "$($s2.DeviceObject)\" | Out-Null }

    #Test hardlink
    try { get-childitem "$ConfigPath\$($s2.ID)" -ErrorAction Stop | Out-Null }
    Catch
    {
        TeeLog -Message "SnapShotVolume: Failed to create valid hardlink to snapshot with ID '$($s2.ID)' for Volume '$volume'" -Logfile $log -Errors $XMLErrors
        try { $Location.RemoveChild($Location.SelectSingleNode("Snapshot[@ID='$($s2.ID)']")) | Out-Null } catch {}
        if ($XMLSnapshot.ID) { & 'cmd.exe' /c rmdir "$ConfigPath\$($s2.ID)" | Out-Null }
        return $false
    }

    return $true
}

function TrxFile($spath, $fname, $dpath, $ipg, $log, $errors)
{
    $results = & "RoboCopy.exe" "$spath" "$dpath" "$fname" /Z /DCOPY:T /COPY:DAT /IPG:$ipg /R:5 /W:30 /NP | Out-String
    TeeLog -Message "RoboCopy Results: $results" -Logfile $log
    TeeLog -Message "Verifying file: '$dpath\$fname'" -Logfile $log
    try 
    {
        $sfile = get-item "$spath\$fname" -ErrorAction Stop
        $dfile = get-item "$dpath\$fname" -ErrorAction Stop
        if (Compare-Object $sfile $dfile -Property Length,LastWriteTime)
        {
            Throw "Source and Destination files do not match."
        }
        TeeLog -Message "Verify success: Source and Destination files are a match." -Logfile $log
    }
    Catch
    {
        TeeLog -Message "Verify failed: '$dpath\$fname' - $($_.Exception.Message)" -Logfile $log -Errors $errors
        return $false
    }

    return $true
}

function DoFind
{
    $Complete = $true #Flag to track job status
    
    #Clear all previous locations and errors
    try { $XMLRoot.RemoveChild($XMLRoot.SelectSingleNode("Locations")) | Out-Null } catch {}
    $XMLLocations=$XMLRoot.AppendChild($XML.CreateElement("Locations"))
    try { $XMLRoot.RemoveChild($XMLRoot.SelectSingleNode("Errors")) | Out-Null } catch {}
    $XMLErrors=$XMLRoot.AppendChild($XML.CreateElement("Errors"))

    #Search each location for .PST files
    foreach ($location in $locations)
    {
        $Files=@()
        $location = $location.TrimEnd("\").ToLower()
        if ($location -eq "alllocal") { $location = "AllLocal" }

        $LocationComplete = $true #Flag to track location status

        $XMLLocation=$XMLLocations.AppendChild($XML.CreateElement("Location"))
        $XMLLocation.SetAttribute("Path", $location)
        $XMLLocation.SetAttribute("Status", "Incomplete")

        $XMLFiles=$XMLLocation.AppendChild($XML.CreateElement("Files"))

        if ($location -eq "AllLocal") #Scan all local drives
        {
            TeeLog -message "Scanning All Local Drives" -Logfile $log
            try
            {
                $Drives = @(get-wmiobject win32_volume -errorAction Stop | ? {$_.DriveType -eq 3 -or $_.DriveType -eq 2 } | Where-Object {$_.DriveLetter -ne $null} | %{get-psdrive $_.DriveLetter[0] -ErrorAction Stop})
                Teelog -message "DriveCount = $($Drives.count)" -Logfile $log

                foreach ($drive in $drives)
                {
                    Teelog "Scanning Drive: $($Drive.Root)" -Logfile $log
                    $files+=RecurseFolder($($Drive.Root))
                }
            }
            catch
            {
                TeeLog -Message "Error while recursing drive '$($Drive.Root)' : $($_.Exception.Message)" -LogFile $Log -Errors $XMLErrors
                $LocationComplete=$false
                $Complete=$false
            }
        }
        else #Scan path
        {
            TeeLog -Message "Scanning Location: '$location'" -Logfile $log
            try
            {
                if (!(Test-Path $location)) { Throw "Path Not Found" }
                $files+=RecurseFolder($location)
            }
            catch
            {
                TeeLog -Message "Error while recursing location '$location' : $($_.Exception.Message)" -LogFile $Log -Errors $XMLErrors
                $LocationComplete=$false
                $Complete=$false
            }
        }

        #Write each file to XML
        foreach ($file in $files)
        {
            if (!$file.FullName) { Continue } #Constantly fixing Powershell 2.0 issues
            
            #Skip special .PSTs
            if (!$file.FullName.ToLower() -like "sharepoint lists") { Continue }
            if (!$file.FullName.ToLower() -like "internet calendar subscriptions") { Continue }
            
            try
            {
                $XMLFile=$XMLFiles.AppendChild($XML.CreateElement("File"))
                $XMLFile.SetAttribute("Path", $file.FullName)
                $XMLFile.SetAttribute("Owner", $(get-acl $file.FullName).Owner)
                $XMLFile.SetAttribute("LastModified", $file.LastWriteTime)
                $XMLFile.SetAttribute("FileSize", $file.Length)
                $XMLFile.SetAttribute("LastProcessed", (get-date).tostring("G")) 
                $XMLFile.SetAttribute("Status", "Found")
            }
            catch
            {
                TeeLog -Message "Error while processing files, file='$($file.FullName)' : $($_.Exception.Message)" -LogFile $Log -Errors $XMLErrors
                $LocationComplete=$false
                $Complete=$false
            }
        }
        if ($LocationComplete)
            { $XMLLocation.SetAttribute("Status", "Found") }
        else
            { $XMLLocation.SetAttribute("Status", "FindError") }
    }

    return $Complete
}

Function DoCollect
{
    #Go through each location and transfer files to collectpath as defined in XML

    $Complete = $true #Flag to track job status

    foreach ($location in $locations)
    {
        $location = $location.trimend("\").ToLower()
        if ($location -eq "alllocal") { $location = "AllLocal" }
        TeeLog -Message "Collecting location '$location'" -Logfile $log
        $XMLLocation = $XMLLocations.SelectSingleNode("Location[@Path='$location']")
        if ($XMLLocation.Status -eq "Found" -or $XMLLocation.Status -eq "CollectError")
        {
            #Collect location
            $LocationComplete = $true #Flag to track status of location

            $XMLFiles = $XMLLocation.SelectSingleNode("Files")
            foreach ($file in $XMLFiles.SelectNodes("*"))
            {
                #Skip if processed already
                if ($file.status -eq "Collected" -or $file.status -eq "Removed" -or $file.status -eq "RemoveError" -or $file.status -eq "Void")
                {
                    TeeLog -Message "File '$($File.Path)' has status '$($file.Status)' - Skipping" -Logfile $log
                }
                else
                {
                    $spath = $null
                    $dpath = $null

                    #Collect file
                    TeeLog -Message "Collecting file: '$($File.path)'" -Logfile $log
                    $file.SetAttribute("LastProcessed", (get-date).tostring("G")) 
                    $fname = Split-Path -path $file.path -Leaf

                    if ($file.path.StartsWith("\\"))
                    {
                        #Build UNC path
                        $spath = Split-Path -path $file.path -Parent
                        $dpath = "$CollectPath\$jobname\" + $(Split-Path -path $file.path -Parent).ToLower().Replace("$location\", "").TrimEnd("\")
                    }
                    else
                    {
                        #Build local path and Snapshot volume
                        if (SnapShotVolume -Volume $(Split-Path -path $file.path -Qualifier) -Location $XMLLocation)
                        {
                            $XMLLocation = $XMLLocations.SelectSingleNode("Location[@Path='$location']")
                            $XMLSnapshot = $XMLLocation.SelectSingleNode("Snapshot[@Volume='$(Split-Path -path $file.path -Qualifier)']")
                            $spath = "$ConfigPath\$($XMLSnapshot.ID)" + $(Split-Path -path $(split-path -path $file.path -noQualifier) -Parent)
                            $dpath = "$CollectPath\$jobname\" + $(Split-Path -path $file.path -Parent).Replace(":","").TrimEnd("\")
                        }
                        else
                        {
                            #Shadowcopy failure
                            $XMLLocation = $XMLLocations.SelectSingleNode("Location[@Path='$location']")
                            TeeLog -Message "$($file.path) - ShadowCopy failure - Collecting without VSS." -Logfile $log
                            $spath = Split-Path -path $file.path -Parent
                            $dpath = "$CollectPath\$jobname\" + $(Split-Path -path $file.path -Parent).Replace(":","").TrimEnd("\")
                        }    
                    }

                    if ($spath -and $dpath)
                    {
                        #Save data
                        $XML.Save("$ConfigPath\$jobname.xml")
                    
                        TeeLog -Message "Transferring file: '$($File.path)'" -Logfile $log
                        if (TrxFile -spath $spath -dpath $dpath -fname $fname -ipg $ipg -log $log -errors $XMLErrors)
                        {
                            TeeLog -Message "Transfer successful: $($File.path)" -Logfile $log
                            $file.SetAttribute("Status", "Collected")
                        }
                        else
                        {
                            TeeLog -Message "Transfer failed: $($File.path)" -Logfile $log
                            $file.SetAttribute("Status", "CollectError")
                            $Complete = $false
                            $LocationComplete = $false
                        }
                    }
                    else
                    {
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
        }
        else
        {
            if ($XMLLocation.Status -eq "Collected")
            { TeeLog -Message "The location '$location' has status '$($XMLLocation.Status)' - Skipping." -Logfile $log }
            else
            {
                TeeLog -Message "The location '$location' has status '$($XMLLocation.Status)' which is other than FOUND, COLLECTED or COLLECTERROR - Skipping." -Logfile $log -Errors $XMLErrors
                $Complete=$false
            }
        }
    }
    
    return $Complete
}

Function DoRemove
{
    #Go through each location and remove files as defined in XML

    $Complete = $true #Flag to track job status

    foreach ($location in $locations)
    {
        $location = $location.trimend("\").ToLower()
        if ($location -eq "alllocal") { $location = "AllLocal" }
        TeeLog -Message "Removing location '$location'" -Logfile $log
        $XMLLocation = $XMLLocations.SelectSingleNode("Location[@Path='$location']")
        if ($XMLLocation.Status -eq "Collected" -or $XMLLocation.Status -eq "RemoveError")
        {
            #Remove location
            $LocationComplete = $true #Flag to track status of location

            $XMLFiles = $XMLLocation.SelectSingleNode("Files")
            foreach ($file in $XMLFiles.SelectNodes("*"))
            {
                #Skip if processed already
                if ($file.status -eq "Removed" -or $file.status -eq "Void")
                {
                    TeeLog -Message "File '$($File.Path)' has status '$($file.Status)' - Skipping" -Logfile $log
                }
                else
                {
                    #Remove file
                    TeeLog -Message "Removing file: '$($File.path)'" -Logfile $log
                    $file.SetAttribute("LastProcessed", (get-date).tostring("G")) 

                    try
                    {
                        $fname = $(Split-Path -path $file.path -Leaf).ToLower()
                        #Skip whitelist files - for backwards compatibility
                        if ($fname -like "*sharepoint lists*" -or $fname -like "*internet calendar subscriptions*")
                        {
                            TeeLog -Message "File Whitelisted - Setting status and skipping: $($File.path)" -Logfile $log
                        }
                        else
                        {
                            remove-item $file.path -Force -ErrorAction Stop
                            TeeLog -Message "Remove successful: $($File.path)" -Logfile $log
                        }
                        $file.SetAttribute("Status", "Removed")
                    }
                    catch
                    {
                        if ($_.Exception.Message -like "*cannot find path*")
                        {
                            TeeLog -Message "Remove successful: File no longer exists - $($File.path)" -Logfile $log
                            $file.SetAttribute("Status", "Removed")
                        }
                        else
                        {
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
        }
        else
        {
            if ($XMLLocation.Status -eq "Removed")
            { TeeLog -Message "The location '$location' has status '$($XMLLocation.Status)' - Skipping." -Logfile $log }
            else
            {
                TeeLog -Message "The location '$location' has status '$($XMLLocation.Status)' which is other than COLLECTED, REMOVED or REMOVEERROR - Skipping." -Logfile $log -Errors $XMLErrors
                $Complete=$false
            }
        }
    }
    
    return $Complete
    
}

#Restart as NT AUTHORITY\System if not
#if ($(whoami) -ne "NT AUTHORITY\System")
#{
#   $Arguments = "/s powershell.exe -ExecutionPolicy Bypass -File `"" + $myInvocation.MyCommand.Definition + "`" "
#   
#   foreach ($param in $PSBoundParameters.GetEnumerator())
#   {
#        $Arguments += "-" + $param.key + " `"" + $param.Value + "`" "
#   }
#
#   #Invoke-TokenManipulation 
#   Start-Process -FilePath "$PSScriptRoot\psexec.exe" -Verb runAS -ArgumentList $Arguments | out-null
#
#   # Exit from the current, unelevated, process
#   exit
#   }

#$configpath = $configpath.trimend("\") - Obsoleted
$configpath = "C:\PSTCollector" #New configpath is forced to C:\PSTCollector to support the hardlinks to shadowcopies.
$collectpath = $collectpath.trimend("\")
$log = $configpath + "\$jobname.log"

Teelog -message "Running As: $Env:Username" -Logfile $log

#Quick and dirty way to check if we might already be running (this happens when a system is put in to standby/hibernation during our collects)
if (get-process | where-object {$_.Name -eq "robocopy"})
{
    Teelog -message "Found process 'Robocopy' already running. This script may already be active - Quitting." -Logfile $log
    Exit 1
}

#Create config path if not exist
New-Item -ItemType Directory -Path $configpath -ErrorAction SilentlyContinue | out-null

$XML = InitializeXML
$XMLRoot = $XML.SelectSingleNode("//Configuration")
$XMLErrors = $XML.SelectSingleNode("//Configuration/Errors")
$XMLParameters = $XML.SelectSingleNode("//Configuration/Parameters")
$XMLLocations = $XML.SelectSingleNode("//Configuration/Locations")


Switch ($mode)
{
    FIND {
        #FIND Mode

        if ($XMLRoot.Status -eq "Found" -or $XMLRoot.Status -eq "Collected" -or $XMLRoot.Status -eq "Removed" -or $XMLRoot.Status -eq "RemoveError" -or $XMLRoot.Status -eq "CollectError")
        {
            TeeLog -message "This job has status '$($XMLRoot.Status)' - Quitting." -Logfile $log
        }
        else
        {
            TeeLog -message "Starting FIND job." -Logfile $log
            if (DoFind)
            {
                TeeLog -message "The FIND operationg completed successfully." -Logfile $log
                $XMLRoot.SetAttribute("Status", "Found")
            }
            else
            {
                TeeLog -message "The FIND operation failed." -Logfile $log
                $XMLRoot.SetAttribute("Status", "FindError")
            }
        }
    }

    COLLECT {
        if ($XMLRoot.Status -eq "FindError" -or $XMLRoot.Status -eq "Incomplete")
        {
            TeeLog -message "This job has status '$($XMLRoot.Status)' - Performing FIND first." -Logfile $log
            if (DoFind)
            {
                TeeLog -message "The FIND operationg completed successfully." -Logfile $log
                $XMLRoot.SetAttribute("Status", "Found")
            }
            else
            {
                TeeLog -message "The FIND operation failed." -Logfile $log
                $XMLRoot.SetAttribute("Status", "FindError")
            }
            $XML.Save("$ConfigPath\$jobname.xml")
            $XMLLocations = $XML.SelectSingleNode("//Configuration/Locations")
        }

        if ($XMLRoot.Status -eq "Found" -or $XMLRoot.Status -eq "CollectError")
        {
            TeeLog -message "Starting COLLECT job." -Logfile $log

            if (DoCollect)
            {
                TeeLog -message "COLLECT job finished successfully." -Logfile $log
                $XMLRoot.SetAttribute("Status", "Collected")
            }
            else
            {
                TeeLog -message "COLLECT job failed." -Logfile $log
                $XMLRoot.SetAttribute("Status", "CollectError")
            }
        }
        else
        {
            if ($XMLRoot.Status -eq "Collected" -or $XMLRoot.Status -eq "Removed" -or $XMLRoot.Status -eq "RemoveError" -or $XMLRoot.Status -eq "FindError")
                { TeeLog -message "This job has status '$($XMLRoot.Status)' - Quitting." -Logfile $log }
            else
            {
                TeeLog -message "This job has an unknown status '$($XMLRoot.Status)' - Quitting." -Logfile $log
            }
        }
    }
    REMOVE {
        if ($XMLRoot.Status -eq "FindError" -or $XMLRoot.Status -eq "Incomplete")
        {
            TeeLog -message "This job has status '$($XMLRoot.Status)' - Performing FIND first." -Logfile $log
            if (DoFind)
            {
                TeeLog -message "The FIND operationg completed successfully." -Logfile $log
                $XMLRoot.SetAttribute("Status", "Found")
            }
            else
            {
                TeeLog -message "The FIND operation failed." -Logfile $log
                $XMLRoot.SetAttribute("Status", "FindError")
            }
            $XML.Save("$ConfigPath\$jobname.xml")
            $XMLLocations = $XML.SelectSingleNode("//Configuration/Locations")
        }

        if ($XMLRoot.Status -eq "CollectError" -or $XMLRoot.Status -eq "Found")
        {
            TeeLog -message "This job has status '$($XMLRoot.Status)' - Performing COLLECT first." -Logfile $log
            if (DoCollect)
            {
                TeeLog -message "The COLLECT operationg completed successfully." -Logfile $log
                $XMLRoot.SetAttribute("Status", "Collected")
            }
            else
            {
                TeeLog -message "The COLLECT operation failed." -Logfile $log
                $XMLRoot.SetAttribute("Status", "CollectError")
            }
            $XML.Save("$ConfigPath\$jobname.xml")
            $XMLLocations = $XML.SelectSingleNode("//Configuration/Locations")
        }

        if ($XMLRoot.Status -eq "Collected" -or $XMLRoot.Status -eq "RemoveError")
        {
            TeeLog -message "Starting REMOVE job." -Logfile $log

            if (DoRemove)
            {
                TeeLog -message "REMOVE job finished successfully." -Logfile $log
                $XMLRoot.SetAttribute("Status", "Removed")
            }
            else
            {
                TeeLog -message "REMOVE job failed." -Logfile $log
                $XMLRoot.SetAttribute("Status", "RemoveError")
            }
        }
        else
        {
            if ($XMLRoot.Status -eq "Removed")
                { TeeLog -message "This job has status '$($XMLRoot.Status)' - Quitting." -Logfile $log }
            else
            {
                TeeLog -message "This job has an unknown status '$($XMLRoot.Status)' - Quitting." -Logfile $log
            }
        }
    }
    DEFAULT {
        TeeLog -message "This job has an unknown MODE '$mode' - Quitting." -Logfile $log
    }
}


try
{
    if ($XML -ne $null)
    {
        #Write data to disk
        $XML.Save("$ConfigPath\$jobname.xml")
 
        #Create collect path if not exist
        New-Item -ItemType Directory -Path "$CollectPath" -ErrorAction SilentlyContinue | out-null
        
        #Write data to collection
        $XML.Save("$CollectPath\$jobname.xml")
        Copy-Item "$Configpath\$jobname.log" "$collectpath" -ErrorAction Stop -Force | out-null
    }
}
catch
{
    TeeLog -Message "Error while saving XMLs during finalize : $($_.Exception.Message)" -LogFile $Log
    Exit 1
}

TeeLog -Message "CollectorAgent - Complete" -LogFile $Log
Exit 0


