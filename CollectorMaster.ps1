<#
.SYNOPSIS
    PST Collector Master - Orchestrates PST file collection across multiple workstations and file shares.

.DESCRIPTION
    This script coordinates the discovery, collection, and removal of PST files across
    an enterprise environment. It can target:
    - Domain-joined workstations via Active Directory OUs (using PsExec to run CollectorAgent)
    - Network file shares (running CollectorAgent locally via background jobs)
    
    After collection, generates a CSV mapping file compatible with Microsoft 365 PST Import.
    
    The script maintains centralized state in an XML configuration file, aggregating
    results from all CollectorAgent instances. Can be run multiple times to catch
    offline computers and resume interrupted operations.

.PARAMETER Mode
    Operation mode: FIND (discover PSTs), COLLECT (copy to collection path), or REMOVE (delete source files).

.PARAMETER JobName
    Unique identifier for this collection job. Must be consistent across all runs.

.PARAMETER Locations
    Array of locations to process:
    - OU paths: "OU=Computers,DC=domain,DC=local"
    - UNC paths: "\\server\share"

.PARAMETER CollectPath
    UNC path where collected PST files will be stored.

.PARAMETER ConfigPath
    Local path for master configuration and log files. Default: C:\PSTCollector

.PARAMETER ForceRestart
    Wipe existing configuration and start fresh. Use with caution.

.PARAMETER Noping
    Skip ping checks before processing computers. Useful when ICMP is blocked.

.PARAMETER ThrottleLimit
    Maximum concurrent background jobs. Default: 25

.PARAMETER NoSkipCommon
    Pass to CollectorAgent to scan all folders including system directories.

.PARAMETER IsArchive
    Set to $true to import PSTs to archive mailbox, $false for primary. Default: $true

.PARAMETER IPG
    RoboCopy interpacket gap in milliseconds. Default is 1ms to throttle network bandwidth.
    Higher values reduce network impact but slow transfers.

.PARAMETER EnabledOnly
    When processing OU locations, only target computers with Enabled=$True in Active Directory.
    By default, all computers in the OU are targeted regardless of enabled state.

.EXAMPLE
    .\CollectorMaster.ps1 -Mode FIND -JobName "Migration2024" `
        -Locations "OU=Workstations,DC=corp,DC=local" -CollectPath "\\server\PSTs"

.EXAMPLE
    .\CollectorMaster.ps1 -Mode COLLECT -JobName "Migration2024" `
        -Locations "\\fileserver\home$","\\fileserver\profiles$" -CollectPath "\\server\PSTs"

.NOTES
    Author: Appleoddity
    Requires: PowerShell 5.0+, Domain Admin privileges, PsExec.exe in script directory
    
    Output Files:
    - MASTER-{JobName}.xml: Aggregated configuration and status
    - MASTER-{JobName}.csv: Microsoft 365 PST Import mapping file
    - MASTER-{JobName}.log: Execution log
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
    [string]$configpath = "$env:SystemDrive\PSTCollector",
  
    [Parameter(Position = 6)]
    [switch]$forcerestart = $false,

    [Parameter(Position = 7)]
    [switch]$noping = $false,

    [Parameter(Position = 8)]
    [Int]$throttlelimit = 25,

    [Parameter(Position = 9)]
    [Switch]$NoSkipCommon = $false,

    [Parameter(Position = 10)]
    [Boolean]$IsArchive = $true,

    [Parameter(Position = 11)]
    [string]$ipg = "1",

    [Parameter(Position = 12)]
    [Switch]$EnabledOnly = $false
)

#region Functions

<#
.SYNOPSIS
    Converts a location path to a relative path for use in file system operations.
.PARAMETER Location
    The location string to sanitize.
.OUTPUTS
    String - Path with leading/trailing backslashes removed
#>
function Sanitize([String]$Location) {
    #Converts a 'location' to a relative path suitable for using in Collect or Config path parameters
    return $Location.TrimStart("\").TrimEnd("\")
}

<#
.SYNOPSIS
    Processes completed background jobs and aggregates results.
.DESCRIPTION
    Monitors PowerShell background jobs for each location/computer, collects their output,
    imports child XML results into the master configuration, and appends child logs.
    Called repeatedly during job execution to process finished work.
#>
function DoStatus {
    Get-Job -HasMoreData $true | Receive-Job | TeeLog -Logfile $log
    Get-Job |
        ForEach-Object {
            if ($_.State -ne "Completed" -and $_.State -ne "Failed") { return } #Only process completed and failed jobs
            
            #Identify existing XML data and prepare it
            $XMLLocation = $XML.SelectSingleNode("//Configuration/Locations/Location[@Path='$($_.Name)']")
            if (!$XMLLocation) { return } #Sanity check
            
            if ($_.State -eq "Completed") { #If job successful
                TeeLog -Message "Collector: Retrieving results for '$($_.Name)'" -logfile $log                
                try { $XMLLocation.RemoveChild($XMLLocation.SelectSingleNode("Errors")) | Out-Null } catch {}
                try { $XMLLocation.RemoveChild($XMLLocation.SelectSingleNode("Files")) | Out-Null } catch {}
                if ($_.Name -like "*\*") { $ChildLocation = $_.Name } else { $ChildLocation = "AllLocal" }
                try { #Pull the XML results from the collect path
                    #Use XmlDocument.Load() for proper Unicode encoding handling
                    $XMLChild = New-Object System.Xml.XmlDocument
                    $XMLChild.Load("$collectpath\$(Sanitize($_.Name))\$jobname.xml")
                     
                    $XMLLocation.AppendChild($XML.ImportNode($XMLChild.SelectSingleNode("//Configuration/Locations/Location[@Path='$ChildLocation']/Files"), $True)) | Out-Null
                    $XMLLocation.AppendChild($XML.ImportNode($XMLChild.SelectSingleNode("//Configuration/Errors"), $True)) | Out-Null
                    $XMLLocation.Status = $XMLChild.SelectSingleNode("//Configuration/Locations/Location[@Path='$ChildLocation']").Status
                    #$XMLLocation.Status=$XMLChild.Configuration.Status
                } catch {
                    $XMLLocation.Status = "Error" #If failed, then mark target as error
                    $Global:NotCompleteFlag = $true
                }
            } else {
                $XMLLocation.Status = "Error" #If failed, then mark target as error
                $Global:NotCompleteFlag = $true
            }
        
            if ($XMLLocation.Status -like "*error" -or $XMLLocation.Status -eq "Incomplete") { $Global:NotCompleteFlag = $true }

            if ($XMLLocation.Path.StartsWith("\\")) { $_ | Receive-Job -Wait } 
            $_ | Remove-Job
            $XML.Save("$configpath\MASTER-$jobname.xml") #Save the new config XML
        
            #Retrieve child's log and parse only current events
            $ChildLog = @()
            $FinishLog = $false
            Get-Content -LiteralPath "$collectpath\$(Sanitize($_.Name))\$jobname.log" -ErrorAction SilentlyContinue |
                foreach {
                    if (!$FinishLog) {
                        try {
                            $logdate = [datetime]::Parse(($_.Split(':', 4) | Select-Object -Index 0, 1, 2) -join ':')
                            $rundate = [datetime]::Parse($XMLRoot.LastRunTime).ToUniversalTime()
                            if ($logdate -ge $rundate) {
                                $FinishLog = $true
                                $ChildLog += $_
                            }
                        } catch { }
                    } else { $ChildLog += $_ }
                }
                if (!($XMLLocation.Path.StartsWith("\\"))) { $ChildLog } #Write child's log to screenSelect-Object
                $ChildLog | Out-File $Log -Append -ErrorAction SilentlyContinue #Append child log to master log
            }
}

<#
.SYNOPSIS
    Writes a message to the log file and console.
.PARAMETER message
    The message text to log.
.PARAMETER logfile
    Path to the log file.
.PARAMETER errors
    Optional XML element for error tracking.
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
        Write-Host $message
        $message = (Get-Date).tostring("G") + ": " + $message 
        $message | Out-File -FilePath $logfile -Append
        if ($errors) {
            $XMLError = $errors.AppendChild($XML.CreateElement("Error"))
            $XMLError.SetAttribute("Message", $message)
            $XMLError.SetAttribute("Time", (Get-Date).ToUniversalTime().tostring("G"))
        }
    } catch { }
}

<#
.SYNOPSIS
    Generates a Microsoft 365 PST Import CSV mapping file from collected PST data.
.DESCRIPTION
    Iterates through all collected PST files in the XML configuration and creates
    a CSV file in the format required by Microsoft 365 PST Import Service.
    
    For each PST:
    - Looks up the file owner's email address in Active Directory
    - Falls back to Administrator's email if lookup fails
    - Builds the Azure blob path based on the collection structure
    - Sets the target import folder path
    
    The resulting CSV can be used directly with the Microsoft 365 Compliance Portal
    PST Import feature after uploading files with AzCopy.
.PARAMETER XMLToExport
    The XML configuration document containing collected file information.
.OUTPUTS
    Boolean - True if any errors occurred, False on success
#>
function ConvertToCSV($XMLToExport) {
    $ErrorFlag = $False
    if (!$XMLToExport)
    { return $true }   #Sanity Check

    $Locations = $XMLToExport.SelectNodes("/Configuration/Locations/*[@Status='Collected']")
    if (!$Locations)
    { return $true } #Sanity Check

    #Export array
    $Export = @()
    
    foreach ($Location in $Locations) {
        if ($Location.Status -eq "Collected") {
            $Files = $Location.SelectNodes("Files/*[@Status='Collected']")
            foreach ($File in $Files) {
                #File Object
                $ExportedFile = New-Object PSObject
                $ExportedFile | Add-Member -MemberType NoteProperty -Name Workload -Value "Exchange"
                $ExportedFile | Add-Member -MemberType NoteProperty -Name FilePath -Value $null
                $ExportedFile | Add-Member -MemberType NoteProperty -Name Name -Value $null
                $ExportedFile | Add-Member -MemberType NoteProperty -Name Mailbox -Value $null
                $ExportedFile | Add-Member -MemberType NoteProperty -Name IsArchive -Value $IsArchive
                $ExportedFile | Add-Member -MemberType NoteProperty -Name TargetRootFolder -Value $null
                $ExportedFile | Add-Member -MemberType NoteProperty -Name ContentCodePage -Value $null
                $ExportedFile | Add-Member -MemberType NoteProperty -Name SPFileContainer -Value $null
                $ExportedFile | Add-Member -MemberType NoteProperty -Name SPManifestContainer -Value $null
                $ExportedFile | Add-Member -MemberType NoteProperty -Name SPSiteUrl -Value $null

                #Get mailbox - default unknowns to Administrator mailbox
                try {
                    $ExportedFile.Mailbox = (Get-AdUser -Identity $(Split-Path -Path $File.Owner -Leaf -ErrorAction SilentlyContinue) -Properties EmailAddress | Select-Object EmailAddress).EmailAddress
                    if (!$ExportedFile.Mailbox) { throw "Missing E-mail address" }
                } catch {
                    $ExportedFile.Mailbox = (Get-AdUser -Identity "Administrator" -Properties EmailAddress | Select-Object EmailAddress).EmailAddress
                }
        
                if ($file.path.StartsWith("\\")) {
                    #Build UNC path
                    $ExportedFile.FilePath = "pstcollector/" + $Location.Path.ToLower().TrimStart("\\").Replace("\", "/") + "/$jobname/" + $(Split-Path -Path $file.path -Parent).ToLower().Replace("$($Location.Path.ToLower())\", "").TrimEnd("\").Replace("\", "/")
                } else {
                    #Build local path
                    $ExportedFile.FilePath = "pstcollector/" + $location.Path + "/$jobname/" + $(Split-Path -Path $file.path -Parent).Replace(":", "").TrimEnd("/").Replace("\", "/")
                }
            
                $ExportedFile.Name = $(Split-Path -Path $File.path -Leaf)
                $ExportedFile.TargetRootFolder = "/ImportedPST/$(Split-Path -Path $ExportedFile.FilePath -Leaf)/" + $(Split-Path -Path $File.path -Leaf).ToLower().Replace(".pst", "")
    
                if ($ExportedFile.Mailbox -and $ExportedFile.FilePath -and $ExportedFile.Name) {
                    $Export += $ExportedFile
                } else {
                    TeeLog -Message "Collector: Can't Export '$($Location.Path)' : '$($File.Path)' to CSV - Missing required info" -logfile $log
                    $ErrorFlag = $true
                }
            }
        }
    }
    try {
        $Export | Export-Csv -NoTypeInformation -Encoding UTF8 "$ConfigPath\MASTER-$jobname.csv"
    } catch {
        TeeLog -Message "Collector: Failed saving export file '$ConfigPath\MASTER-$jobname.csv': $($_.Exception.Message)" -logfile $log
        $ErrorFlag = $true
    }
    return $ErrorFlag
}

#endregion Functions

#region Main Script
# ==================================================================================
# MAIN SCRIPT EXECUTION
# ==================================================================================

# Clean up any orphaned jobs from previous runs
Get-Job | Remove-Job
    
$XML = $null

# Normalize paths
$configpath = $configpath.trimend("\")
$collectpath = $collectpath.trimend("\")
$log = $configpath + "\MASTER-$jobname.log"

# Create config directory if it doesn't exist
New-Item -ItemType Directory -Path $configpath -ErrorAction SilentlyContinue | Out-Null

TeeLog -Message "Collector: CollectorMaster STARTING" -logfile $log

# Load existing configuration if available and ForceRestart not specified
if (!$forcerestart) {
    if (Test-Path -LiteralPath "$configpath\MASTER-$jobname.xml") {
        # Use XmlDocument.Load() for proper Unicode encoding handling
        TeeLog -Message "Collector: Loading existing config $configpath\MASTER-$jobname.xml" -logfile $log                
        $XML = New-Object System.Xml.XmlDocument
        $XML.Load("$configpath\MASTER-$jobname.xml")
    }
}

# Initialize new XML if none loaded
if (!$XML) {
    TeeLog -Message "Collector: Initializing new config $configpath\MASTER-$jobname.xml" -logfile $log                
    $XML = New-Object System.Xml.XmlDocument
}

$XMLRoot = $XML.SelectSingleNode("//Configuration")
if (!$XMLRoot) { $XMLRoot = $XML.AppendChild($XML.CreateElement("Configuration")) }
    
$XMLParameters = $XML.SelectSingleNode("//Configuration/Parameters")
if (!$XMLParameters) { $XMLParameters = $XMLRoot.AppendChild($XML.CreateElement("Parameters")) }

$XMLLocations = $XML.SelectSingleNode("//Configuration/Locations")
if (!$XMLLocations) { $XMLLocations = $XMLRoot.AppendChild($XML.CreateElement("Locations")) }

$XMLErrors = $XML.SelectSingleNode("//Configuration/Errors")
if (!$XMLErrors) { $XMLErrors = $XMLRoot.AppendChild($XML.CreateElement("Errors")) }

if (($XMLRoot.Status -eq "Collected" -or $XMLRoot.Status -eq "Removed") -and $mode -eq "collect") {
    TeeLog -Message "Collector: Job has status '$($XMLRoot.Status)' - Quitting." -logfile $log                
    exit
}

if (($XMLRoot.Status -eq "Found" -or $XMLRoot.Status -eq "Collected" -or $XMLRoot.Status -eq "Removed") -and $mode -eq "find") {
    TeeLog -Message "Collector: Job has status '$($XMLRoot.Status)' - Quitting." -logfile $log                
    exit
}

if ($XMLRoot.Status -eq "Removed" -and $mode -eq "remove") {
    TeeLog -Message "Collector: Job has status '$($XMLRoot.Status)' - Quitting." -logfile $log                
    exit
}

$XMLRoot.SetAttribute("ComputerName", "$Env:ComputerName")
$XMLRoot.SetAttribute("Status", "Incomplete")
$XMLRoot.SetAttribute("Description", "This is the PST CollectorMaster configuration file.")
$XMLRoot.SetAttribute("LastRunTime", (Get-Date).tostring("G"))

try { $XMLParameters.SelectSingleNode("Parameter[@Jobname]").SetAttribute("Jobname", $JobName) }
catch { 
    $XMLParameter = $XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("Jobname", $Jobname)
}

try { $XMLParameters.SelectSingleNode("Parameter[@Mode]").SetAttribute("Mode", $mode) }
catch {
    $XMLParameter = $XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("Mode", $Mode)
}

try { $XMLParameters.SelectSingleNode("Parameter[@ConfigPath]").SetAttribute("ConfigPath", $ConfigPath) }
catch {
    $XMLParameter = $XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("ConfigPath", $ConfigPath)
}

try { $XMLParameters.SelectSingleNode("Parameter[@CollectPath]").SetAttribute("CollectPath", $CollectPath) }
catch {
    $XMLParameter = $XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("CollectPath", $CollectPath)
}

try { $XMLParameters.SelectSingleNode("Parameter[@ForceRestart]").SetAttribute("ForceRestart", $ForceRestart) }
catch {
    $XMLParameter = $XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("ForceRestart", $ForceRestart)
}

try { $XMLParameters.SelectSingleNode("Parameter[@ThrottleLimit]").SetAttribute("ThrottleLimit", $ThrottleLimit) }
catch {
    $XMLParameter = $XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("ThrottleLimit", $ThrottleLimit)
}

try { $XMLParameters.SelectSingleNode("Parameter[@NoSkipCommon]").SetAttribute("NoSkipCommon", $NoSkipCommon) }
catch {
    $XMLParameter = $XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("NoSkipCommon", $NoSkipCommon)
}

try { $XMLParameters.SelectSingleNode("Parameter[@IPG]").SetAttribute("IPG", $ipg) }
catch {
    $XMLParameter = $XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("IPG", $ipg)
}

try { $XMLParameters.SelectSingleNode("Parameter[@EnabledOnly]").SetAttribute("EnabledOnly", $EnabledOnly) }
catch {
    $XMLParameter = $XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("EnabledOnly", $EnabledOnly)
}

[boolean]$Global:NotCompleteFlag = $false

foreach ($location in $locations) {
    $location = $location.TrimEnd("\").tolower()

    #Network location
    if ($location.StartsWith("\\")) {
        TeeLog -Message "Collector: Processing location '$location'" -Logfile $log

        #Check status of location
        $XMLLocation = $XML.SelectSingleNode("//Configuration/Locations/Location[@Path='$location']")
        
        if ($XMLLocation) {
            #Skip if status VOID - a way to disable locations that should not be looked at anymore
            if ($XMLLocation.Status -eq "void") {
                TeeLog -Message "Collector: '$location' has status '$($XMLLocation.Status)'. Skipping." -Logfile $log
                continue
            }

            #Skip if found, collected, or removed already
            if (($XMLLocation.Status -eq "found" -or $XMLLocation.Status -eq "collected" -or $XMLLocation.Status -eq "removed" -or $XMLLocation.Status -eq "CollectError" -or $XMLLocation.Status -eq "RemoveError") -and $Mode -eq "Find") {
                TeeLog -Message "Collector: '$location' has status '$($XMLLocation.Status)'. Skipping." -Logfile $log
                continue
            }
            if (($XMLLocation.Status -eq "collected" -or $XMLLocation.Status -eq "removed" -or $XMLLocation.Status -eq "RemoveError" ) -and $mode -eq "Collect") {
                TeeLog -Message "Collector: '$location' has status '$($XMLLocation.Status)'. Skipping." -Logfile $log
                continue
            }
            if ($XMLLocation.Status -eq "removed" -and $mode -eq "Remove") {
                TeeLog -Message "Collector: '$location' has status '$($XMLLocation.Status)'. Skipping." -Logfile $log
                continue
            }
        } else { $XMLLocation = $XMLLocations.AppendChild($XML.CreateElement("Location")) }

        $XMLLocation.SetAttribute("Path", $location)
        $XMLLocation.SetAttribute("Status", "Incomplete")
        $XMLLocation.SetAttribute("LastAttempt", (Get-Date).tostring("G"))

        if (!$noping) {
            if (!(Test-Connection -ComputerName $($location.Split('\')[2]) -Quiet -Count 1)) {
                TeeLog -Message "Collector: '$($location.Split('\')[2])' is offline. Skipped." -Logfile $log
                $XMLLocation.SetAttribute("Status", "Offline")
                $Global:NotCompleteFlag = $true
                continue
            }
        }             

        New-Item -ItemType Directory -Path "c:\PSTCollector" -ErrorAction SilentlyContinue | Out-Null
        Copy-Item -LiteralPath "$PSScriptRoot\CollectorAgent.ps1" -Destination "c:\PSTCollector" -ErrorAction Stop -Force | Out-Null

        if ($ForceRestart -or ($XMLLocation.Status -eq "restart")) {
            Start-Job -Name $location -FilePath c:\PSTCollector\CollectorAgent.ps1 -ArgumentList $mode, $jobname, $location, "$collectpath\$(Sanitize($location))", "$ConfigPath\$(Sanitize($location))", $True, $false, $ipg | Out-Null
        } else {
            Start-Job -Name $location -FilePath c:\PSTCollector\CollectorAgent.ps1 -ArgumentList $mode, $jobname, $location, "$collectpath\$(Sanitize($location))", "$ConfigPath\$(Sanitize($location))", $false, $false, $ipg | Out-Null
        }
    }

    #Organizational Unit
    if ($location.StartsWith("ou=")) {
        
        if ($EnabledOnly) {
            $Computers = @(Get-ADComputer -Filter { Enabled -eq $True } -SearchBase $location)
        } else {
            $Computers = @(Get-ADComputer -Filter * -SearchBase $location)
        }
        TeeLog -Message "Collector: Processing location '$location' with '$($Computers.count)' computers." -Logfile $log

        foreach ($computer in $computers) {
            #Check status of location
            $XMLLocation = $XML.SelectSingleNode("//Configuration/Locations/Location[@Path='$($Computer.Name)']")

            if ($XMLLocation) {
                #Skip if status VOID - a way to disable locations that should not be looked at anymore
                if ($XMLLocation.Status -eq "void") {
                    #TeeLog -Message "Collector: '$($Computer.Name)' has status '$($XMLLocation.Status)'. Skipping." -Logfile $log
                    continue
                }

                #Skip if found or collected already
                if (($XMLLocation.Status -eq "found" -or $XMLLocation.Status -eq "collected" -or $XMLLocation.Status -eq "removed" -or $XMLLocation.Status -eq "CollectError" -or $XMLLocation.Status -eq "RemoveError") -and $Mode -eq "Find") {
                    #TeeLog -Message "Collector: '$($Computer.Name)' has status '$($XMLLocation.Status)'. Skipping." -Logfile $log
                    continue
                }
                if (($XMLLocation.Status -eq "collected" -or $XMLLocation.Status -eq "removed" -or $XMLLocation.Status -eq "RemoveError") -and $mode -eq "Collect") {
                    #TeeLog -Message "Collector: '$($Computer.Name)' has status '$($XMLLocation.Status)'. Skipping." -Logfile $log
                    continue
                }
                if ($XMLLocation.Status -eq "removed" -and $mode -eq "Remove") {
                    #TeeLog -Message "Collector: '$($Computer.Name)' has status '$($XMLLocation.Status)'. Skipping." -Logfile $log
                    continue
                }
            } else { $XMLLocation = $XMLLocations.AppendChild($XML.CreateElement("Location")) }
            
            TeeLog -Message "Collector: Processing computer '$($Computer.Name)'" -Logfile $log
            $XMLLocation.SetAttribute("Path", $computer.name)
            $XMLLocation.SetAttribute("Status", "Incomplete")
            $XMLLocation.SetAttribute("LastAttempt", (Get-Date).tostring("G"))

            if (!$noping) {
                if (!(Test-Connection -ComputerName $Computer.Name -Quiet -Count 1)) {
                    TeeLog -Message "Collector: '$($Computer.Name)' is offline. Skipped." -Logfile $log
                    $XMLLocation.SetAttribute("Status", "Offline")
                    $Global:NotCompleteFlag = $true
                    continue
                }
            }  
            Start-Job -Name $computer.Name -ScriptBlock {
                param ($Root, $CompName, $ChildMode, $ChildJob, $ChildRestart, $ChildConfig, $ChildCollect, $ChildNoSkipCommon, $ChildIPG)
                try {
                    New-Item -ItemType Directory -Path "\\$CompName\c$\PSTCollector" -ErrorAction SilentlyContinue | Out-Null
                    Copy-Item -LiteralPath "$Root\CollectorAgent.ps1" -Destination "\\$CompName\c$\PSTCollector" -ErrorAction Stop -Force | Out-Null
                    if ($ChildRestart) { $ExtraParam += "-ForceRestart " }
                    if ($ChildNoSkipCommon) { $ExtraParam += "-NoSkipCommon " }
                    $ExtraParam += "-ipg $ChildIPG"
                    & "$Root\psexec.exe" \\$($CompName) -nobanner -accepteula -s powershell.exe -ExecutionPolicy Bypass  -Command c:\PSTCollector\CollectorAgent.ps1 -mode $ChildMode -JobName $ChildJob -Locations AllLocal -ConfigPath $ChildConfig -CollectPath $ChildCollect $ExtraParam 2>$null
                } catch { }
                if ($LastExitCode -ne 0) { throw "Error returned during agent process: $LastExitCode" }
            } -ArgumentList $PSScriptRoot, $($Computer.Name), $mode, $jobname, $($ForceRestart -or ($XMLLocation.Status -eq "restart")), $ConfigPath\$($Computer.Name), $CollectPath\$($Computer.Name), $NoSkipCommon, $ipg | Out-Null
        
            if ($XMLLocation.Path.StartsWith("\\")) { Get-Job -HasMoreData $true | Receive-Job | TeeLog -Logfile $log }
            DoStatus
        
            while (@(Get-Job -State Running).count -ge $throttlelimit) {
                Write-Host "Throttling..."
                if ($XMLLocation.Path.StartsWith("\\")) { Get-Job -HasMoreData $true | Receive-Job | TeeLog -Logfile $log }
                Start-Sleep -s 10 #Throttle maximum number of simultaenous jobs.
                DoStatus
            }
        }
    }

    if ($XMLLocation.Path.StartsWith("\\")) { Get-Job -HasMoreData $true | Receive-Job | TeeLog -Logfile $log }
    DoStatus
        
    while (@(Get-Job -State Running).count -ge $throttlelimit) {
        Write-Host "Throttling..."
        if ($XMLLocation.Path.StartsWith("\\")) { Get-Job -HasMoreData $true | Receive-Job | TeeLog -Logfile $log }
        Start-Sleep -s 10 #Throttle maximum number of simultaenous jobs.
        DoStatus
    }
}

while ($(Get-Job -State Running).Count -gt 0) {
    if ($XMLLocation.Path.StartsWith("\\")) { Get-Job -HasMoreData $true | Receive-Job | TeeLog -Logfile $log }
    DoStatus
    Sleep -Seconds 10
    Write-Host "Waiting for jobs to complete..."
    Get-Job -State Running | Select-Object Name, Status | Format-Table -AutoSize
}

DoStatus
if ($mode -eq "collect") {
    if (ConvertToCsv($XML)) {
        $Global:NotCompleteFlag = $true
    }
}

if (!$Global:NotCompleteFlag) {
    if ($mode -eq "find") { $XMLRoot.Status = "Found" }
    if ($mode -eq "collect") { $XMLRoot.Status = "Collected" }
    if ($mode -eq "remove") { $XMLRoot.Status = "Removed" }
}
$XML.Save("$configpath\MASTER-$jobname.xml") #Save the config XML
Write-Host "Completed."