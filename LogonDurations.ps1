<#
.SYNOPSIS
    This script gets the logon duration for endpoints.

.DESCRIPTION
    This script gets the logon duration for endpoints.

    Pseudo-code:
    Given a list of Sites.
    Foreach Site:        
        Find all computers in AD for this Site
        Randomise the list of computers from AD (this is because this script will be run multiple times a day, we want a different computer with no one logged on each time we run this script)
        Foreach computer:
            Get the currently logged on user
                Is there no one currently logged on?
                    Yes: Add this computer to a list and break out of the for loop
                    No: Do nothing
    Foreach computer in the list:
        Copy LockComputer.bat to \\$computerName\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup
        Open up an SCCM remote control window for that computer
        Prompt the user to press enter once he/she has typed in the username/password and pressed enter to begin the logon
        Once the user has pressed enter in the script, record the current time and add this to a list called ComputersWithNoOneLoggedOn with the computername as the key
        Once the computer has finished logging in, it will run the batch file and lock automatically as a security mechanism in case the network crashes and we lose connection to that computer.
    Foreach computer name key in ComputersWithNoOneLoggedOn:
        Delete LockComputer.bat from \\$computerName\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup
        Get the last modified time of '\\<computerName>\C$\Users\<userNameOfUserRunningThisScript>\Tracing\Communicator-uccapi-0.uccapilog'
        Get the logon start time from ComputersWithNoOneLoggedOn
        Find the difference between these times and record this as the logon duration
        Restart the computer (to force you to log out)
           
.INPUTS
    See Param input parameters below.

.OUTPUTS
    A csv file.

.EXAMPLE    

.NOTES
    This script needs to be run from a epd server using your epd priv account.

    Author: dklempfner@gmail.com
    Date: 04/07/2017

    Updates:
    Date: 04/08/2017
    Now copying over a batch file to the computer's startup folder to automatically lock the computer once logging in has finished.
    This is a security mechanism in case we lose connection to that computer, we don't want it just sitting there for anyone to come and use the priv account.
#>

Param([String]$InputFilePath = 'C:\InputFile.csv',
      [String]$OutputFilePath = 'C:\OutputFile.csv',
      [String]$CmRcViewerFilePath = 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe', #This doesn't need to be changed.
      [String]$LockComputerBatchFilePath = 'C:\GetLogonDurations\LockComputer.bat',
      [String]$LdapPath = 'LDAP://DC=epd,DC=def,DC=ghi,DC=au')

function WriteProgress
{
    Param([Parameter(Mandatory = $true)][Int32]$Numerator,
          [Parameter(Mandatory = $true)][Int32]$Denominator,
          [Parameter(Mandatory = $true)][String]$Status)

    $percentComplete = [int](($Numerator/$Denominator)*100)
    if($percentComplete -le 100)
    {
        Write-Progress -Activity $Status -Status "$Numerator/$Denominator -  $percentComplete%" -PercentComplete $percentComplete
    }
    else
    {
        Write-Warning "Cannot call Write-Progress. $percentComplete% is not valid."
    }
}

function ValidateInputFilePath
{
    param([Parameter(Mandatory=$true)][String]$InputFilePath,
          [Parameter(Mandatory=$true)][String]$DesiredFileExtension)

    if(!(Test-Path $InputFilePath))
    {
        Write-Error "Please make sure $InputFilePath exists and then rerun the script."
    }
    if([System.IO.Path]::GetExtension($InputFilePath) -ne $DesiredFileExtension)
    {
        Write-Error "Please make sure $InputFilePath is a .csv file and then rerun the script."
    }
}

function GetDirectorySearcher
{
    param([Parameter(Mandatory=$true)][String]$LdapPath,
          [Parameter(Mandatory=$false)][Object[]]$Properties)
        
    $directoryEntry = New-Object 'System.DirectoryServices.DirectoryEntry'    
    $directoryEntry.Path = $LdapPath
    $searcher = New-Object 'System.DirectoryServices.DirectorySearcher'($directoryEntry)    
    $searcher.SearchScope = [System.DirectoryServices.SearchScope]::SubTree    
    $searcher.PageSize = [System.Int32]::MaxValue
      
    if($Properties)
    {        
        foreach($property in $Properties)
        {
            $searcher.PropertiesToLoad.Add($property) | Out-Null
        }
    }
    
    return $searcher
}

function GetAllComputersFromSite
{
    param([Parameter(Mandatory=$true)][String]$LdapPath,
          [Parameter(Mandatory=$true)][String]$Site)

    $propertiesToLoad = @('name')
    $directorySearcher = GetDirectorySearcher $LdapPath $propertiesToLoad
    $directorySearcher.Filter = "(&(objectCategory=computer)(name=$($Site)PWE7*))"
    $results = $directorySearcher.FindAll()
    return $results
}

function IsNoOneLoggedOn
{
    param([Parameter(Mandatory=$true)][String]$ComputerName)

    try
    {
        $computerSystemInfo = Get-WmiObject -ComputerName $ComputerName -Class Win32_ComputerSystem
        $currentlyLoggedOnUser = $computerSystemInfo.UserName
        $couldGetCompSysInfo = $true
    }
    catch
    {
        $erroMsg = $_.Exception.Message
        Write-Warning "Couldn't get currently logged on user for $ComputerName - $erroMsg"
        $couldGetCompSysInfo = $false
    }

    $isNoOneLoggedOn = $couldGetCompSysInfo -and !$currentlyLoggedOnUser
    return $isNoOneLoggedOn
}

function GetLogonFinishTime
{
    param([Parameter(Mandatory=$true)][String]$ComputerName)

    $lyncLogFile = "\\$ComputerName\C$\Users\$env:USERNAME\Tracing\Communicator-uccapi-0.uccapilog"
    if(Test-Path $lyncLogFile)
    {
        try
        {
            $fileInfo = New-Object System.IO.FileInfo($lyncLogFile)
            return $fileInfo.LastWriteTime
        }
        catch
        {
            $errorMsg = $_.Exception.Message
            Write-Warning "Could not get LastWriteTime for $lyncLogFile - $errorMsg"
        }                
    }
    else
    {
        Write-Warning "$lyncLogFile does not exist. Cannot determine logon finish time."
    }

    return $null
}

function ShuffleTheResults
{
    param([Parameter(Mandatory=$true)][Object[]]$Results)

    $resultsShuffled = $Results | Sort-Object {Get-Random}
    return $resultsShuffled
}

function GetListOfComputerCustomObjects
{
    param([Parameter(Mandatory=$true)][Object[]]$Sites)

    $computerCustomObjects = @()

    for($i = 0; $i -lt $Sites.Count; $i++)
    {
        $site = $Sites[$i].Site

        WriteProgress ($i+1) $Sites.Count "Finding computers in $site"

        $numOfComputersFoundWithNoOneLoggedOn = 0
   
        $results = GetAllComputersFromSite $LdapPath $site        

        if(!$results -or ($results -and $results.Count -eq 0))
        {
            $computerCustomObjects += [PSCustomObject]@{Site = $site; ComputerName = ''; LogonStartTime = ''; LogonFinishTime = ''; LogonDuration = ''; WereEndpointsFoundInAdForThisSite = $false; WasAtLeastOneEndpointFoundWithNoOneLoggedOnForThisSite = $false; WasLockComputerBatchFileCopiedSuccessfully = ''; WasLockComputerBatchFileDeletedSuccessfully = ''}
        }
        else
        {
            $results = ShuffleTheResults $results
            $wasAComputerFoundWithNoOneLoggedOn = $false
            for($j = 0; $j -lt $results.Count; $j++)
            {                        
                WriteProgress ($j+1) $results.Count "Finding computers with no one logged on at $site"

                $result = $results[$j]

                if($result.Properties.name.Count -gt 0)
                {
                    $computerName = $result.Properties.name[0]
                    $isNoOneLoggedOn = IsNoOneLoggedOn $computerName
                    if($isNoOneLoggedOn)
                    {
                        $wasAComputerFoundWithNoOneLoggedOn = $true
                        $customObject = [PSCustomObject]@{Site = $site; ComputerName = $computerName; LogonStartTime = ''; LogonFinishTime = ''; LogonDuration = ''; WereEndpointsFoundInAdForThisSite = $true; WasAtLeastOneEndpointFoundWithNoOneLoggedOnForThisSite = $true; WasLockComputerBatchFileCopiedSuccessfully = ''; WasLockComputerBatchFileDeletedSuccessfully = ''}
                        $computerCustomObjects += $customObject
                        break
                    }
                }
            }
            if(!$wasAComputerFoundWithNoOneLoggedOn)
            {
                $computerCustomObjects += [PSCustomObject]@{Site = $site; ComputerName = ''; LogonStartTime = ''; LogonFinishTime = ''; LogonDuration = ''; WereEndpointsFoundInAdForThisSite = $true; WasAtLeastOneEndpointFoundWithNoOneLoggedOnForThisSite = $false; WasLockComputerBatchFileCopiedSuccessfully = ''; WasLockComputerBatchFileDeletedSuccessfully = ''}
            }
        }
    }

    return $computerCustomObjects
}

function DeleteLockComputerBatchFileFromComutersWithNoOneLoggedOn
{
    param([Parameter(Mandatory=$true)][Object[]]$ComputersWithNoOneLoggedOn,
          [Parameter(Mandatory=$true)][String]$LockComputerBatchFilePath)
    
    for($o = 0; $o -lt $ComputersWithNoOneLoggedOn.Count; $o++)
    {
        $computerWithNoOneLoggedOn = $ComputersWithNoOneLoggedOn[$o]
        $computerName = $computerWithNoOneLoggedOn.ComputerName
        if($computerName)
        {
            $computerStartupFolderPath = "\\$computerName\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"

            WriteProgress ($o+1) $ComputersWithNoOneLoggedOn.Count "Deleting lock computer batch file from $computerStartupFolderPath"

            $lockComputerBatchFileName = [System.IO.Path]::GetFileName($LockComputerBatchFilePath)
            $lockComputerBatchFilePathOnComputer = [System.IO.Path]::Combine($computerStartupFolderPath, $lockComputerBatchFileName)

            if(Test-Path $lockComputerBatchFilePathOnComputer)
            {
                try
                {
                    Remove-Item $lockComputerBatchFilePathOnComputer
                    $computerWithNoOneLoggedOn.WasLockComputerBatchFileDeletedSuccessfully = $true
                }
                catch
                {
                    $errorMsg = $_.Exception.Message
                    Write-Warning "Could not delete $lockComputerBatchFilePathOnComputer - $errorMsg"
                    $computerWithNoOneLoggedOn.WasLockComputerBatchFileDeletedSuccessfully = $false
                }
            }
            else
            {
                Write-Warning "Tried to delete the batch file $lockComputerBatchFilePathOnComputer but it didn't exist."
            }
        }
        else
        {
            Write-Warning "The element at index $o in the collection had a blank computer name. This should not happen."
        }
    }
}

function CopyLockComputerBatchFileToComutersWithNoOneLoggedOn
{
    param([Parameter(Mandatory=$true)][Object[]]$ComputersWithNoOneLoggedOn,
          [Parameter(Mandatory=$true)][String]$LockComputerBatchFilePath)
    
    for($n = 0; $n -lt $ComputersWithNoOneLoggedOn.Count; $n++)
    {
        $computerWithNoOneLoggedOn = $ComputersWithNoOneLoggedOn[$n]
        $computerName = $computerWithNoOneLoggedOn.ComputerName
        if($computerName)
        {
            $computerStartupFolderPath = "\\$computerName\c$\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup"

            WriteProgress ($n+1) $ComputersWithNoOneLoggedOn.Count "Copying lock computer batch file to $computerStartupFolderPath"

            if(Test-Path $computerStartupFolderPath)
            {
                try
                {
                    Copy-Item $LockComputerBatchFilePath $computerStartupFolderPath
                    $computerWithNoOneLoggedOn.WasLockComputerBatchFileCopiedSuccessfully = $true
                }
                catch
                {
                    $errorMsg = $_.Exception.Message
                    Write-Warning "Could not copy batch file to $computerStartupFolderPath - $errorMsg"
                    $computerWithNoOneLoggedOn.WasLockComputerBatchFileCopiedSuccessfully = $false
                }
            }
            else
            {
                Write-Warning "Could not find start folder for $computerName. This computer will not be automatically locked when logged on to."
            }
        }
        else
        {
            Write-Warning "The element at index $n in the collection had a blank computer name. This should not happen."
        }
    }
}

function OpenUpSccmRemoteControlWindows
{
    param([Parameter(Mandatory=$true)][Object[]]$ComputersWithNoOneLoggedOn)

    for($k = 0; $k -lt $ComputersWithNoOneLoggedOn.Count; $k++)
    {
        $computerWithNoOneLoggedOn = $ComputersWithNoOneLoggedOn[$k]
        $computerName = $computerWithNoOneLoggedOn.ComputerName
        if($computerName)
        {
            WriteProgress ($k+1) $computersWithNoOneLoggedOn.Count "Opening SCCM remote control window for $computerName"

            & $CmRcViewerFilePath $computerName #Opens up the SCCM remote control window
            Write-Verbose "$computerName's remote control window should now be open."
            Read-Host "Quickly press enter here as soon as you have typed in your username/password and clicked enter"
            $computerWithNoOneLoggedOn.LogonStartTime = Get-Date
        }
        else
        {
            Write-Warning "The element at index $k in the collection had a blank computer name. This should not happen."
        }
    }
}

function CalculateLogonDuration
{
    param([Parameter(Mandatory=$true)][Object[]]$ComputersWithNoOneLoggedOn)

    for($l = 0; $l -lt $ComputersWithNoOneLoggedOn.Count; $l++)
    {
        $computerWithNoOneLoggedOn = $ComputersWithNoOneLoggedOn[$l]
        $computerName = $computerWithNoOneLoggedOn.ComputerName

        WriteProgress ($l+1) $computersWithNoOneLoggedOn.Count "Calculating the logon duration for $computerName"
        
        $computerWithNoOneLoggedOn.LogonDuration = 0

        $computerWithNoOneLoggedOn.LogonFinishTime = GetLogonFinishTime $computerName
        if($computerWithNoOneLoggedOn.LogonFinishTime)
        {
            $computerWithNoOneLoggedOn.LogonDuration = $computerWithNoOneLoggedOn.LogonFinishTime - $computerWithNoOneLoggedOn.LogonStartTime
        }
    }
}

function RestartComputers
{
    param([Parameter(Mandatory=$true)][Object[]]$ComputersWithNoOneLoggedOn)

    for($m = 0; $m -lt $ComputersWithNoOneLoggedOn.Count; $m++)
    {
        $computerWithNoOneLoggedOn = $ComputersWithNoOneLoggedOn[$m]
        $computerName = $computerWithNoOneLoggedOn.ComputerName

        WriteProgress ($m+1) $computersWithNoOneLoggedOn.Count "Restarting $computerName"

        RestartComputer $computerName
    }
}

function RestartComputer
{
    param([Parameter(Mandatory=$true)][String]$ComputerName)

    shutdown -r -t 0 -m "\\$ComputerName"
}

cls

$VerbosePreference = 'Continue'
$ErrorActionPreference = 'Stop'

try
{
    $dateTimeFormat = 'yyyy-MM-dd HH:mm:ss'
    $startTime = Get-Date
    $startTimeString = $startTime.ToString($dateTimeFormat)
    $timeZone = [System.TimeZone]::CurrentTimeZone.StandardName
    $sixNewLines = "$([Environment]::NewLine)$([Environment]::NewLine)$([Environment]::NewLine)$([Environment]::NewLine)$([Environment]::NewLine)$([Environment]::NewLine)"
    $startText = "$($sixNewLines)Get Logon durations script started at $startTimeString $timeZone"
    Write-Verbose $startText

    ValidateInputFilePath $CmRcViewerFilePath '.exe'

    ValidateInputFilePath $InputFilePath '.csv'

    ValidateInputFilePath $LockComputerBatchFilePath '.bat'

    $sites = @(Import-Csv $InputFilePath)

    $computerCustomObjects = GetListOfComputerCustomObjects $sites
    
    $computersWithNoOneLoggedOn = $computerCustomObjects | Where-Object { $_.WasAtLeastOneEndpointFoundWithNoOneLoggedOnForThisSite -eq $true }

    if($computersWithNoOneLoggedOn)
    {
        CopyLockComputerBatchFileToComutersWithNoOneLoggedOn $computersWithNoOneLoggedOn $LockComputerBatchFilePath

        OpenUpSccmRemoteControlWindows $computersWithNoOneLoggedOn

        Read-Host 'Press enter once all the computers have finished logging in.'

        DeleteLockComputerBatchFileFromComutersWithNoOneLoggedOn $computersWithNoOneLoggedOn $LockComputerBatchFilePath

        CalculateLogonDuration $computersWithNoOneLoggedOn

        RestartComputers $computersWithNoOneLoggedOn
    }
    else
    {
        Write-Warning 'For all Sites, either no endpoints were found in AD, or there were endpoints in AD but all of them had someone logged on.'
    }
}
catch
{
    $errorMsg = $_.Exception.Message
    Write-Warning $errorMsg
}
finally
{
    $endTime = Get-Date
    $duration = $endTime.Subtract($startTime)
    $endTimeString = $endTime.ToString($dateTimeFormat)

    $computerCustomObjects | Export-Csv $OutputFilePath -NoTypeInformation

    $finishedText = "Logon duration script finished at $endTimeString $timeZone. Output file is located at $OutputFilePath.$([Environment]::NewLine)"
    $finishedText += "Run by $($env:USERDOMAIN)/$($env:USERNAME) on $($env:COMPUTERNAME)$([Environment]::NewLine)"
    $finishedText += "Execution time: $($duration.Hours) Hours : $($duration.Minutes) Minutes : $($duration.Seconds) Seconds$([Environment]::NewLine)"

    Write-Verbose $finishedText
}