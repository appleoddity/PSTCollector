# This is the PST Collector Master script. It handles the firing of the CollectorAgent and network scans. It also gathers and centralizes information.
#
# Usage: CollectorMaster.ps1 -Mode <mode> -JobName <jobname> -Locations <locations> -CollectPath <path> [-ConfigPath <path>] [-ForceRestart] [-noping] [-throttlelimit <xx>] [-NoSkipCommon] [-IsArchive <True | False>]
#
# Specify locations as an Organizational Unit path (i.e. OU=COMPUTERS,DC=DOMAIN,DC=LOCAL) or network file path (i.e. \\server\share\path) separted by commas.
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
        [string]$configpath="$env:SystemDrive\PSTCollector",
  
    [Parameter(Position=6)]
        [switch]$forcerestart=$false,

    [Parameter(Position=7)]
        [switch]$noping=$false,

    [Parameter(Position=8)]
        [Int]$throttlelimit=25,

    [Parameter(Position=9)]
        [Switch]$NoSkipCommon=$false,

    [Parameter(Position=10)]
        [Boolean]$IsArchive=$true
)

function Sanitize([String]$Location)
{
    #Converts a 'location' to a relative path suitable for using in Collect or Config path parameters
    Return $Location.TrimStart("\").TrimEnd("\")
}

function DoStatus
{
    Get-Job -HasMoreData $true | Receive-Job | TeeLog -Logfile $log
    Get-Job |
    Foreach-Object {
        if ($_.State -ne "Completed" -and $_.State -ne "Failed") { Return } #Only process completed and failed jobs
            
        #Identify existing XML data and prepare it
        $XMLLocation = $XML.SelectSingleNode("//Configuration/Locations/Location[@Path='$($_.Name)']")
        if (!$XMLLocation) { Return } #Sanity check
            
        if ($_.State -eq "Completed") #If job successful
        {
            TeeLog -Message "Collector: Retrieving results for '$($_.Name)'" -logfile $log                
            try { $XMLLocation.RemoveChild($XMLLocation.SelectSingleNode("Errors")) | Out-Null } catch {}
            try { $XMLLocation.RemoveChild($XMLLocation.SelectSingleNode("Files")) | Out-Null } catch {}
            if ($_.Name -like "*\*") { $ChildLocation = $_.Name } else { $ChildLocation = "AllLocal" }
            Try #Pull the XML results from the collect path
                {
                    [System.Xml.XmlDocument]$XMLChild = get-content "$collectpath\$(Sanitize($_.Name))\$jobname.xml"
                     
                    $XMLLocation.AppendChild($XML.ImportNode($XMLChild.SelectSingleNode("//Configuration/Locations/Location[@Path='$ChildLocation']/Files"), $True)) | Out-Null
                    $XMLLocation.AppendChild($XML.ImportNode($XMLChild.SelectSingleNode("//Configuration/Errors"), $True)) | Out-Null
                    $XMLLocation.Status=$XMLChild.SelectSingleNode("//Configuration/Locations/Location[@Path='$ChildLocation']").Status
                    #$XMLLocation.Status=$XMLChild.Configuration.Status
                }
                catch
                {
                    $XMLLocation.Status = "Error" #If failed, then mark target as error
                    $Global:NotCompleteFlag = $true
                }
        }
        else
        {
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
        Get-Content("$collectpath\$(Sanitize($_.Name))\$jobname.log") -ErrorAction SilentlyContinue |
        foreach {
            if (!$FinishLog)
            {
                try
                {
                    $logdate = [datetime]::Parse(($_.Split(':',4) | Select -Index 0,1,2) -join ':')
                    $rundate = [datetime]::Parse($XMLRoot.LastRunTime).ToUniversalTime()
                    if ($logdate -ge $rundate)
                    {
                        $FinishLog = $true
                        $ChildLog+=$_
                    }
                }
                catch { }
            }
            else { $ChildLog+=$_ }
        }
        if (!($XMLLocation.Path.StartsWith("\\"))) { $ChildLog } #Write child's log to screen
        $ChildLog | Out-File $Log -Append -ErrorAction SilentlyContinue #Append child log to master log
    }
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
        write-host $message
        $message = (get-date).tostring("G") + ": " + $message 
        $message | out-file -FilePath $logfile -Append
        if ($errors) 
        {
            $XMLError=$errors.AppendChild($XML.CreateElement("Error"))
            $XMLError.SetAttribute("Message", $message)
            $XMLError.SetAttribute("Time", (get-date).ToUniversalTime().tostring("G"))
        }
    }
    catch { }
}

function ConvertToCSV($XMLToExport)
{
    $ErrorFlag = $False
    if (!$XMLToExport)
    { return $true }   #Sanity Check

    $Locations = $XMLToExport.SelectNodes("/Configuration/Locations/*[@Status='Collected']")
    if (!$Locations)
    { return $true } #Sanity Check

    #Export array
    $Export = @()
    
    Foreach($Location in $Locations)
    {
        if ($Location.Status = "Collected")
        {
            $Files = $Location.SelectNodes("Files/*[@Status='Collected']")
            Foreach($File in $Files)
            {
                #File Object
                $ExportedFile = new-object PSObject
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
                Try
                {
                    $ExportedFile.Mailbox = (Get-AdUser -Identity $(Split-Path -Path $File.Owner -Leaf) -Properties EmailAddress | Select EmailAddress).EmailAddress
                    if (!$ExportedFile.Mailbox) { Throw "Missing E-mail address" }
                }
                catch
                {
                $ExportedFile.Mailbox = (Get-AdUser -Identity "Administrator" -Properties EmailAddress | Select EmailAddress).EmailAddress
                }
        
                if ($file.path.StartsWith("\\"))
                {
                    #Build UNC path
                    $ExportedFile.FilePath = "pstcollector/" + $Location.Path.ToLower().TrimStart("\\").Replace("\","/") + "/$jobname/" + $(Split-Path -path $file.path -Parent).ToLower().Replace("$($Location.Path.ToLower())\", "").TrimEnd("\").Replace("\","/")
                }
                else
                {
                #Build local path
                    $ExportedFile.FilePath = "pstcollector/" + $location.Path + "/$jobname/" + $(Split-Path -path $file.path -Parent).Replace(":","").TrimEnd("/").Replace("\", "/")
                }
            
                $ExportedFile.Name = $(Split-Path -Path $File.path -Leaf)
                $ExportedFile.TargetRootFolder = "/ImportedPST/$(Split-Path -path $ExportedFile.FilePath -Leaf)/" + $(Split-Path -Path $File.path -Leaf).ToLower().Replace(".pst","")
    
                if ($ExportedFile.Mailbox -and $ExportedFile.FilePath -and $ExportedFile.Name)
                {
                    $Export+=$ExportedFile
                }
                else
                {
                    TeeLog -Message "Collector: Can't Export '$($Location.Path)' : '$($File.Path)' to CSV - Missing required info" -logfile $log
                    $ErrorFlag = $true
                }
            }
        }
    }
    Try
    {
        $Export | Export-Csv -NoTypeInformation "$ConfigPath\MASTER-$jobname.csv"
    }
    catch
    {
        TeeLog -Message "Collector: Failed saving export file '$ConfigPath\MASTER-$jobname.csv': $($_.Exception.Message)" -logfile $log
        $ErrorFlag = $true
    }
    Return $ErrorFlag
}

Get-Job | Remove-Job #Clear old jobs
    
$XML = $null
$Jobs=@()

$configpath = $configpath.trimend("\")
$collectpath = $collectpath.trimend("\")
$log = $configpath + "\MASTER-$jobname.log"

#Create config path if not exist
New-Item -ItemType Directory -Path $configpath -ErrorAction SilentlyContinue | out-null

TeeLog -Message "Collector: CollectorMaster STARTING" -logfile $log

if (!$forcerestart)
{
    #Config file exists?
    if (Test-Path "$configpath\MASTER-$jobname.xml")
    {
        #Load it
        TeeLog -Message "Collector: Loading existing config $configpath\MASTER-$jobname.xml" -logfile $log                
        [System.Xml.XmlDocument]$XML = get-content "$configpath\MASTER-$jobname.xml" 
    }
}

#Initialize XML
if (!$XML)
{
    TeeLog -Message "Collector: Initializing new config $configpath\MASTER-$jobname.xml" -logfile $log                
    $XML=New-Object System.Xml.XmlDocument
}

$XMLRoot=$XML.SelectSingleNode("//Configuration")
if (!$XMLRoot) { $XMLRoot = $XML.AppendChild($XML.CreateElement("Configuration")) }
    
$XMLParameters=$XML.SelectSingleNode("//Configuration/Parameters")
if (!$XMLParameters) { $XMLParameters = $XMLRoot.AppendChild($XML.CreateElement("Parameters")) }

$XMLLocations=$XML.SelectSingleNode("//Configuration/Locations")
if (!$XMLLocations) { $XMLLocations = $XMLRoot.AppendChild($XML.CreateElement("Locations")) }

$XMLErrors=$XML.SelectSingleNode("//Configuration/Errors")
if (!$XMLErrors) { $XMLErrors = $XMLRoot.AppendChild($XML.CreateElement("Errors")) }

if (($XMLRoot.Status -eq "Collected" -or $XMLRoot.Status -eq "Removed") -and $mode -eq "collect")
{
    TeeLog -Message "Collector: Job has status '$($XMLRoot.Status)' - Quitting." -logfile $log                
    exit
}

if (($XMLRoot.Status -eq "Found" -or $XMLRoot.Status -eq "Collected" -or $XMLRoot.Status -eq "Removed") -and $mode -eq "find")
{
    TeeLog -Message "Collector: Job has status '$($XMLRoot.Status)' - Quitting." -logfile $log                
    exit
}

if ($XMLRoot.Status -eq "Removed" -and $mode -eq "remove")
{
    TeeLog -Message "Collector: Job has status '$($XMLRoot.Status)' - Quitting." -logfile $log                
    exit
}

$XMLRoot.SetAttribute("ComputerName", "$Env:ComputerName")
$XMLRoot.SetAttribute("Status", "Incomplete")
$XMLRoot.SetAttribute("Description", "This is the PST CollectorMaster configuration file.")
$XMLRoot.SetAttribute("LastRunTime", (get-date).tostring("G"))

Try { $XMLParameters.SelectSingleNode("Parameter[@Jobname]").SetAttribute("Jobname",$JobName) }
catch
{ 
    $XMLParameter=$XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("Jobname", $Jobname)
}

Try { $XMLParameters.SelectSingleNode("Parameter[@Mode]").SetAttribute("Mode",$mode) }
Catch
{
    $XMLParameter=$XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("Mode", $Mode)
}

Try { $XMLParameters.SelectSingleNode("Parameter[@ConfigPath]").SetAttribute("ConfigPath",$ConfigPath) }
Catch
{
    $XMLParameter=$XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("ConfigPath", $ConfigPath)
}

Try { $XMLParameters.SelectSingleNode("Parameter[@CollectPath]").SetAttribute("CollectPath",$CollectPath) }
Catch
{
    $XMLParameter=$XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("CollectPath", $CollectPath)
}

Try { $XMLParameters.SelectSingleNode("Parameter[@ForceRestart]").SetAttribute("ForceRestart",$ForceRestart) }
Catch
{
    $XMLParameter=$XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("ForceRestart", $ForceRestart)
}

Try { $XMLParameters.SelectSingleNode("Parameter[@ThrottleLimit]").SetAttribute("ThrottleLimit",$ThrottleLimit) }
Catch
{
    $XMLParameter=$XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("ThrottleLimit", $ThrottleLimit)
}

Try { $XMLParameters.SelectSingleNode("Parameter[@NoSkipCommon]").SetAttribute("NoSkipCommon",$NoSkipCommon) }
Catch
{
    $XMLParameter=$XMLParameters.AppendChild($XML.CreateElement("Parameter"))
    $XMLParameter.SetAttribute("NoSkipCommon", $NoSkipCommon)
}

[boolean]$Global:NotCompleteFlag = $false

foreach ($location in $locations)
{
    $location = $location.TrimEnd("\").tolower()

    #Network location
    if ($location.StartsWith("\\"))
    {
        TeeLog -Message "Collector: Processing location '$location'" -Logfile $log

        #Check status of location
        $XMLLocation = $XML.SelectSingleNode("//Configuration/Locations/Location[@Path='$location']")
        
        if ($XMLLocation) 
        {
            #Skip if status VOID - a way to disable locations that should not be looked at anymore
            if ($XMLLocation.Status -eq "void")
            {
                TeeLog -Message "Collector: '$location' has status '$($XMLLocation.Status)'. Skipping." -Logfile $log
                Continue
            }

            #Skip if found, collected, or removed already
            if (($XMLLocation.Status -eq "found" -or $XMLLocation.Status -eq "collected" -or $XMLLocation.Status -eq "removed" -or $XMLLocation.Status -eq "CollectError" -or $XMLLocation.Status -eq "RemoveError") -and $Mode -eq "Find")
            {
                TeeLog -Message "Collector: '$location' has status '$($XMLLocation.Status)'. Skipping." -Logfile $log
                Continue
            }
            if (($XMLLocation.Status -eq "collected" -or $XMLLocation.Status -eq "removed" -or $XMLLocation.Status -eq "RemoveError" ) -and $mode -eq "Collect")
            {
                TeeLog -Message "Collector: '$location' has status '$($XMLLocation.Status)'. Skipping." -Logfile $log
                Continue
            }
            if ($XMLLocation.Status -eq "removed" -and $mode -eq "Remove")
            {
                TeeLog -Message "Collector: '$location' has status '$($XMLLocation.Status)'. Skipping." -Logfile $log
                Continue
            }
        }
        else { $XMLLocation=$XMLLocations.AppendChild($XML.CreateElement("Location")) }

        $XMLLocation.SetAttribute("Path", $location)
        $XMLLocation.SetAttribute("Status", "Incomplete")
        $XMLLocation.SetAttribute("LastAttempt", (get-date).tostring("G"))

        if (!$noping)
        {
            if (!(test-connection -ComputerName $($location.Split('\')[2]) -Quiet -Count 1))
            {
                TeeLog -Message "Collector: '$($location.Split('\')[2])' is offline. Skipped." -Logfile $log
                $XMLLocation.SetAttribute("Status", "Offline")
                $Global:NotCompleteFlag=$true
                continue
            }
        }             

        New-Item -ItemType Directory -Path "c:\PSTCollector" -ErrorAction SilentlyContinue | out-null
        Copy-Item "$PSScriptRoot\CollectorAgent.ps1" "c:\PSTCollector" -ErrorAction Stop -Force | out-null

        if ($ForceRestart)
        {
            Start-Job -Name $location -FilePath c:\PSTCollector\CollectorAgent.ps1 -ArgumentList $mode, $jobname, $location, "$collectpath\$(Sanitize($location))", "$ConfigPath\$(Sanitize($location))", $True | Out-Null
        }
        else
        {
            Start-Job -Name $location -FilePath c:\PSTCollector\CollectorAgent.ps1 -ArgumentList $mode, $jobname, $location, "$collectpath\$(Sanitize($location))", "$ConfigPath\$(Sanitize($location))" | Out-Null
        }
    }

    #Organizational Unit
    if ($location.StartsWith("ou=")) {
        
        $Computers = @(Get-ADComputer -filter * -SearchBase $location)
        TeeLog -Message "Collector: Processing location '$location' with '$($Computers.count)' computers." -Logfile $log

        foreach ($computer in $computers)
        {
            #Check status of location
            $XMLLocation = $XML.SelectSingleNode("//Configuration/Locations/Location[@Path='$($Computer.Name)']")

            if ($XMLLocation) 
            {
                #Skip if status VOID - a way to disable locations that should not be looked at anymore
                if ($XMLLocation.Status -eq "void")
                {
                    #TeeLog -Message "Collector: '$($Computer.Name)' has status '$($XMLLocation.Status)'. Skipping." -Logfile $log
                    Continue
                }

                #Skip if found or collected already
                if (($XMLLocation.Status -eq "found" -or $XMLLocation.Status -eq "collected" -or $XMLLocation.Status -eq "removed" -or $XMLLocation.Status -eq "CollectError" -or $XMLLocation.Status -eq "RemoveError") -and $Mode -eq "Find")
                {
                    #TeeLog -Message "Collector: '$($Computer.Name)' has status '$($XMLLocation.Status)'. Skipping." -Logfile $log
                    Continue
                }
                if (($XMLLocation.Status -eq "collected" -or $XMLLocation.Status -eq "removed" -or $XMLLocation.Status -eq "RemoveError") -and $mode -eq "Collect")
                {
                    #TeeLog -Message "Collector: '$($Computer.Name)' has status '$($XMLLocation.Status)'. Skipping." -Logfile $log
                    Continue
                }
                if ($XMLLocation.Status -eq "removed" -and $mode -eq "Remove")
                {
                    #TeeLog -Message "Collector: '$($Computer.Name)' has status '$($XMLLocation.Status)'. Skipping." -Logfile $log
                    Continue
                }
            }
            else { $XMLLocation=$XMLLocations.AppendChild($XML.CreateElement("Location")) }
            
            TeeLog -Message "Collector: Processing computer '$($Computer.Name)'" -Logfile $log
            $XMLLocation.SetAttribute("Path", $computer.name)
            $XMLLocation.SetAttribute("Status", "Incomplete")
            $XMLLocation.SetAttribute("LastAttempt", (get-date).tostring("G"))

            if (!$noping)
            {
                if (!(test-connection -ComputerName $Computer.Name -Quiet -Count 1))
                {
                    TeeLog -Message "Collector: '$($Computer.Name)' is offline. Skipped." -Logfile $log
                    $XMLLocation.SetAttribute("Status", "Offline")
                    $Global:NotCompleteFlag=$true
                    continue
                }
            }  
            Start-Job -Name $computer.Name -ScriptBlock {
                param ($Root, $CompName, $ChildMode, $ChildJob, $ChildRestart, $ChildConfig, $ChildCollect, $ChildNoSkipCommon)
                try
                {
                    New-Item -ItemType Directory -Path "\\$CompName\c$\PSTCollector" -ErrorAction SilentlyContinue | out-null
                    Copy-Item "$Root\CollectorAgent.ps1" "\\$CompName\c$\PSTCollector" -ErrorAction Stop -Force | out-null
                    if ($ChildRestart) { $ExtraParam += "-ForceRestart " }
                    if ($ChildNoSkipCommon) { $ExtraParam += "-NoSkipCommon" }
                    & "$Root\psexec.exe" \\$($CompName) -nobanner -accepteula -s powershell.exe -ExecutionPolicy Bypass  -Command c:\PSTCollector\CollectorAgent.ps1 -mode $ChildMode -JobName $ChildJob -Locations AllLocal -ConfigPath $ChildConfig -CollectPath $ChildCollect $ExtraParam 2>$null
                }
                catch { }
                if ($LastExitCode -ne 0) { throw "Error returned during agent process: $LastExitCode" }
            } -ArgumentList $PSScriptRoot,$($Computer.Name), $mode, $jobname, $ForceRestart, $ConfigPath\$($Computer.Name), $CollectPath\$($Computer.Name), $NoSkipCommon | out-null
        
            if ($XMLLocation.Path.StartsWith("\\")) { Get-Job -HasMoreData $true | Receive-Job | TeeLog -Logfile $log }
            DoStatus
        
            while (@(Get-Job -State Running).count -ge $throttlelimit)
            {
                Write-Host "Throttling..."
                if ($XMLLocation.Path.StartsWith("\\")) { Get-Job -HasMoreData $true | Receive-Job | TeeLog -Logfile $log }
                Start-Sleep -s 10 #Throttle maximum number of simultaenous jobs.
                DoStatus
            }
        }
    }

    if ($XMLLocation.Path.StartsWith("\\")) { Get-Job -HasMoreData $true | Receive-Job | TeeLog -Logfile $log }
    DoStatus
        
    while (@(Get-Job -State Running).count -ge $throttlelimit)
    {
        Write-Host "Throttling..."
        if ($XMLLocation.Path.StartsWith("\\")) { Get-Job -HasMoreData $true | Receive-Job | TeeLog -Logfile $log }
        Start-Sleep -s 10 #Throttle maximum number of simultaenous jobs.
        DoStatus
    }
}

While ($(Get-Job -State Running).Count -gt 0) {
    if ($XMLLocation.Path.StartsWith("\\")) { Get-Job -HasMoreData $true | Receive-Job | TeeLog -Logfile $log }
    DoStatus
    Sleep -Seconds 10
    Write-Host "Waiting for jobs to complete..."
    Get-Job -State Running | select Name,Status | format-table -AutoSize
}

DoStatus
if ($mode -eq "collect")
{
    if (ConvertToCsv($XML))
    {
        $Global:NotCompleteFlag = $true
    }
}

if (!$Global:NotCompleteFlag)
{
    if ($mode -eq "find") { $XMLRoot.Status = "Found" }
    if ($mode -eq "collect") { $XMLRoot.Status = "Collected" }
    if ($mode -eq "remove") { $XMLRoot.Status = "Removed" }
}
$XML.Save("$configpath\MASTER-$jobname.xml") #Save the config XML
Write-Host "Completed."

