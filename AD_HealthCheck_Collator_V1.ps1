#--------------------------------------------------------------------------------------------------------------------------------------------------------#
#                                                                                                                                                        #
#                                              AD Health Check Script                                                                                    #
#                                                                                                                                                        #
#                                       For Any debugging Check Script Transcript located in                                                             # 
#                                           folder "Script directory\Collator\transcripts"                                                               #
#                                                                                                                                                        #
#                                                                                                                                                        #
#--------------------------------------------------------------------------------------------------------------------------------------------------------#


<#

NAME:
      AD_HealthCheck_Collator_V1.ps1

AUTHOR Email :


SYNOPSIS:

      Active Directory Health Check solution will run on Active Directory server 
      and will make the automatic discovery of machines in active directory and 
      will check the below parameters of available domain controllers.

USAGE:
      If script name ends with .txt then first change the file exetension from .txt to .ps1
      - Open Powershell
      - Type cd < Directory where script is kept>
      - Press ENTER key
      - & '.\AD_HealthCheck_Collator_V1.ps1'    
      Example: & '.\AD_HealthCheck_Collator_V1.ps1'
      - Press ENTER key

      Change Details: Optimized the Script

Last Execution Result:      
      Successfully tested

#>


#---------------------------------------------------------------------------------------------------------------------------
# Read Parameters from CSV File
#---------------------------------------------------------------------------------------------------------------------------

# Stop any previous transcript running in the session
try{
  stop-transcript|out-null
}
catch [System.InvalidOperationException]{
    Write-host "No Transcripts Running"
}

# First Clear any variables
Remove-Variable * -ErrorAction SilentlyContinue


# Start script execution time calculation
$ScrptStartTime = (Get-Date).ToString('dd-MM-yyyy hh:mm:ss')
$sw = [Diagnostics.Stopwatch]::StartNew()

# Get Script Directory
$Scriptpath = $($MyInvocation.MyCommand.Path)
$Dir = $(Split-Path $Scriptpath);

# Remove if an previous log files are present
Remove-Item "$Dir\Collator\log\*" -Force -ErrorAction Continue

# Create New Collator Log Folder if not exists

if (!(Test-Path "$Dir\Collator\log")) {New-Item "$Dir\Collator\log" -type Directory -Force }

# Name of Agent Task on Each Domain Controller
$taskname = 'ADHealthCheck'

# Report
$runntime= (get-date -format dd_MM_yyyy-HH_mm_ss)-as [string]
$HealthReport = "$dir\Reports" + "$runntime" + ".htm"

# Logfile 
$Logfile = "$dir\Log" + "$runntime" + ".htm"

$params = import-csv "$Dir\config.csv"

# E-mail report details
$SendEmail     = $params.SendEmail.Trim()
$emailFrom     = $params.EmailFrom.Trim()
$emailTo       = $params.EmailTo.Trim()
#[string]$To =    $emailTo.Split(';')
$smtpServer    = $params.SmtpServer.Trim()
$emailSubject  = $params.EmailSubject.Trim()
$retentionDays = -($params.RetentionDays).Trim()
[int]$diskspacethresold = $params.diskspacethresold.Trim()

[string]$date = Get-Date

# Wait time in SECONDS for KEY PRESS EVENT to get credential
[int]$KeyPressWaitTime = $params.KeyPressWaitTime

# Maximum Number of times Script Should Check the Scheduled tasks on domain controllers.
[int]$MaxCheckLimit = $params.MaxCheckLimit

# Wait time for which Script should wait before Checking again. Value should be in Seconds
[int]$WaitTime = $params.WaitTime

# Wait time (when file for a machine is not found on Network Path) for which Script should wait before Checking again. Value should be in Seconds
$FileNotFoundWaitTime = $params.FileNotFoundWaitTime

# Array Will auto be populated by discovered domain controller machines
$DCList = @()

# Array Will be auto populated where permission will be denied to create the task
$machines_with_failed_task_creation = @()
$machines_with_failed_task_creation_list = @()

# Array Will be auto populated where Necessary ADHealthCheck Directory Structure won't be able to be created
$machines_with_failed_directoy_creation = @()
$machines_with_failed_directoy_creation_list = @()

# Array Will be auto populated where Log files will not be found and will need to be rechecked
$FileRecheckMachines = @()

# Array Will be auto populated where Log files will not be found
$machines_with_file_not_found = @()
$machines_with_file_not_found_list = @()










#---------------------------------------------------------------------------------------------------------------------------------------------
# Functions Section
#---------------------------------------------------------------------------------------------------------------------------------------------


#-------------------------
# Function CleanFiles
#-------------------------

Function CleanFiles { 
    <#
        .SYNOPSIS
            Function Clean Transcript, Log File and Reports older than retentiondays
        .EXAMPLE
            CleanFiles -p "$Dir" -d 2 -f ".*.txt"

            This shows the help for the example function
    #>
    [cmdletbinding()]
    Param 
    (
        # Path to be cleaned like C:\temp, C:\User\Documents etc.
        [Parameter(Mandatory=$True)]
        [Alias("p")]
        $Path,
           
        # Days of file retention like 3, 30 or 60
        [Parameter(Mandatory=$True)]
        [Alias("d")]
        $retentionDays,
        
        #Filename Regex like *.txt, *.html etc.
        [Parameter(Mandatory=$True)]
        [Alias("f")]
        [string[]]$filenameregex
    )

    "Path: $Path"
    "Retention Days: $retentionDays"
    "File name Regex $filenameregex"
    $CurrentDate = Get-Date
    $DatetoDelete = $CurrentDate.AddDays($retentionDays)
    "Date to Delete: $DatetoDelete"
     Get-ChildItem $Path | Where-Object {$_.Name -match $filenameregex} | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item
}



#-------------------------
# Function CheckTaskStatus
#-------------------------

Function CheckTaskStatus
{
   # Function Checks status of ADHealth Check Task on each domain controller

   [cmdletbinding()]
    Param 
    (
     [Parameter(Mandatory=$True)]
     $DCList,   
     [Parameter(Mandatory=$True)]
     [string[]]$PendingMachines
    )

   foreach ($machine in $PendingMachines) 
   {
      $task = schtasks /Query /S $machine /TN $taskname /fo list
      if($?)
      {
     
        [string]$Name = $task | Select-String -Pattern "(TaskName:\s+\\)(.*)" | %{$_.Matches} | %{ $_.Groups[2]} | %{$_.Value}
        [string]$Status = $task | Select-String -Pattern "(Status:\s+)(.*)" | %{$_.Matches} | %{ $_.Groups[2]} | %{$_.Value}
        if($Status -eq 'Running')
        {
        Write-Host -ForegroundColor Cyan "`nMachine: $machine"
        Write-Host -ForegroundColor Yellow "Task $Name is found but still task status is $Status`n"
        } 
        else 
        {
        Write-Host -ForegroundColor Cyan "`nMachine: $machine"
        Write-Host -ForegroundColor Green "Task $Name is found and Status is $Status`n"
        $DCList += $machine
        $Script:PendingMachines = $Script:PendingMachines -ne $machine
        }
      }      
      else 
      {
         Continue
      }
   }
}


#---------------------------
# Function CheckandCopyFile
#---------------------------

Function CheckandCopyFile { 
    <#
        .SYNOPSIS
            Checks file on each domain controller and Copies to machine running collator script
        .EXAMPLE
            CheckandCopyFile -m $dc
    #>
    [cmdletbinding()]
    Param 
    (
        # Machine i.e. Domain Contoller FQDN
        [Parameter(Mandatory=$True)]
        [Alias("m")]
        $machine
    )
		$src = "\\$machine\c$\ADHealthCheck\log\"
		$dest = "$Dir\Collator\log\"

		$files = Get-ChildItem $src -Filter "*.txt"
		
		if($files) {
			$files | Copy-Item -Destination $dest -force
			return $true
		}
		else 
		{
			return $false
		}  
}


Write-host -ForegroundColor Cyan "`nDirectory of Script is: $Dir`n"

# Check and Create Transcript Path
if (!(Test-path "$Dir\Collator\transcripts\")) 
{
   New-Item "$Dir\Collator\transcripts" -ItemType directory
}

$today = (Get-Date).ToString('MM_d_yyyy_hhmmss')

# Let's start transcripting everything
$trnPath = "$Dir\Collator\transcripts\transcript_$today.txt"
Start-Transcript -Path $trnPath -NoClobber
Write-Host "$trnPath`n"



#---------------------------------------------------------------------------------------------------------------------------
# Start Creating Log File
#---------------------------------------------------------------------------------------------------------------------------

"<!DOCTYPE html>
<html>
<head>
<style>
.loginfo 
{
    color: #2437FC;
}
.loghostname
{
    color: #8C280D;
    font-weight: bold;
}
.logwarning
{
    color: 3B5AA27;
}
.logerror
{
    color: #FF3018
}
.logsuccess
{
    color: #058B15
}
.loghighlight
{
font-size: 20px;
color: #DCA53D;
}
</style>
</head>
<body>" | Out-File $Logfile




#---------------------------------------------------------------------------------------------------------------------------
# Script wait to get credentials  
#---------------------------------------------------------------------------------------------------------------------------

Write-Host -ForegroundColor Cyan "`nPlease Press ANY key to provide the Credentials. Script will wait for $KeyPressWaitTime seconds. In case of no key press it will continue `n"

$secondsRunning = 0
while($secondsRunning -lt $KeyPressWaitTime) 
{
   if ($host.UI.RawUI.KeyAvailable) 
   {
      $Key = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp,IncludeKeyDown")
      Break
   }
   Write-Host ("Waiting for: " + ($KeyPressWaitTime-$secondsRunning) + " Seconds" )
   Start-Sleep -Seconds 1
   $secondsRunning++
}

if($Key) 
{
   $keypressed = $Key.Character
   Write-Host -ForegroundColor Green "`n As you have pressed a key (Key: $keypressed). So now you will be asked for the Administrator Credentials`n"
   $User = Get-Credential
   $Username = $User.UserName
   $SecurePassword = $User.Password
   $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
   $Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)   
} 
else 
{
   Write-Host -ForegroundColor Green "`n As no input received from the user. So script will proceed without asking for Administrator Credentials`n"
}




#---------------------------------------------------------------------------------------------------------------------------
# Collator Task Existence Confirmation and Creation
#---------------------------------------------------------------------------------------------------------------------------

$CollatorTask = schtasks /Query /TN 'ADHealthCheckCollator' /fo list

if (!$CollatorTask) 
{
   Write-host 'No Collator task exist.`nScript will now try to create the ADHealthCheckCollator Task'
   "<p><span class=loginfo>No Collator task exist.`nScript will create ADHealthCheckCollator Task</span></p>" | Out-File -Append $Logfile

   $CollatorTaskStartTime = Get-Date -Format HH:mm
   $CollatorTaskStartDate = Get-Date -Format MM/dd/yyyy
   $CollatorTaskFrequency = 'Daily'
   $CollatorTasklevel = 'Highest'
   $TaskAction = "PowerShell.exe -ExecutionPolicy Bypass -NonInteractive -File $Dir\AD_HealthCheck_Collator_V1.ps1"

   if($UserName) 
   {
      schtasks /Create /S $env:COMPUTERNAME /RU $UserName /RP $Password /RL $CollatorTasklevel /TN 'ADHealthCheckCollator' /TR $TaskAction /ST $CollatorTaskStartTime /SD $CollatorTaskStartDate /SC $CollatorTaskFrequency /F

      # Confirm if Collator Task is Created
      $CollatorTaskReConfirm = schtasks /Query /TN 'ADHealthCheckCollator' /fo list
      if (!$CollatorTaskReConfirm) 
      {
        "<p><span class=logerror>Script Couldn't Create the Collator task on machine:</span><span class=loghostname>$env:COMPUTERNAME</span></p>" | Out-File -Append $Logfile
      } 
      else 
      {
        "<p><span class=logsuccess>Task Created Successfully on:</span><span class=loghostname>$env:COMPUTERNAME</span></p>" | Out-File -Append $Logfile
      }
   } 
   else 
   {
      Write-Host -ForegroundColor Yellow "`nAs No Credentials have been provided. `nSo script will not be able to create ADHealthCheckCollator Scheduled Task on this machine`n"
      Write-Host -ForegroundColor Yellow "`n`nPlease Create the ADHealthCheckCollator scheduled task manually `nor please RERUN the script and provide the credntial when asked`n"
      "<p><span class=logwarning>No Credentials have been provided. `nSo script will not be able to create ADHealthCheckCollator Scheduled Task</span></p>" | Out-File -Append $Logfile
   }
} 
else 
{
   if($UserName) 
   {
	schtasks /Change /S $env:COMPUTERNAME /RU $UserName /RP $Password /RL $CollatorTasklevel /TN 'ADHealthCheckCollator' /TR $TaskAction /ST $CollatorTaskStartTime /SD $CollatorTaskStartDate /SC $CollatorTaskFrequency /F
   }
   
   Write-host -foregroundcolor Green "`nADHealthCheckCollator Task Exist. Script will proceed further`n"
   "<p><span class=loginfo>ADHealthCheckCollator Task Exist. Script will proceed further</span></p>" | Out-File -Append $Logfile
}




#---------------------------------------------------------------------------------------------------------------------------
# Setting the header for the Report
#---------------------------------------------------------------------------------------------------------------------------

[DateTime]$DisplayDate = ((get-date).ToUniversalTime())

$header = "
      <!DOCTYPE html>
		<html>
		<head>
        <link rel='shortcut icon' href='favicon.png' type='image/x-icon'>
        <meta charset='utf-8'>
		<meta name='viewport' content='width=device-width, initial-scale=1.0'>		
		<title>AD health Check</title>
		<script type=""text/javascript"">
		  function Powershellparamater(htmlTable)
		  {

			 var myWindow = window.open('', '_blank');
			 myWindow.document.write(htmlTable);
		  }
		  window.onscroll = function (){


			 if (window.pageYOffset == 0) {
				document.getElementById(""toolbar"").style.display = ""none"";
			 }
			 else {
				if (window.pageYOffset > 150) {
				   document.getElementById(""toolbar"").style.display = ""block"";
				}
			 }
		  }

		  function HideTopButton() {
			 document.getElementById(""toolbar"").style.display = ""none"";
		  }
		</script>
		<style>
        <style>
		    #toolbar  
            {
				position: fixed;
				width: 100%;
				height: 25px;
				top: 0;
				left: 0;
				/**/
				text-align: right;
				display: none;
			}
			#backToTop  
            {
				font-family: Segoe UI;
				font-weight: bold;
				font-size: 20px;
				color: #9A2701;
				background-color: #ffffff;
			}


			#Reportrer  
            {
				width: 95%;
				margin: 0 auto;
			}


			body 
            {
				color: #333333;
				font-family: Calibri,Tahoma;
				font-size: 10pt;
				background-color: #616060;
			}

			.odd  
            {
				background-color: #ffffff;
			}

			.even  
            {
				background-color: #dddddd;
			}
			
			table
			{
				background-color: #616060;
				width: 100%;
				color: #fff;
				margin: auto;
				border: 1px groove #000000;
				border-collapse: collapse;
			}
			
			caption
			{
				background-color: #D9D7D7;
				color: #000000;
			}

			.bold_class
			{
				background-color: #ffffff;
				color: #000000;
				font-weight: 550;
			}

			
            td  
            {
				text-align: left;
				font-size: 14px;
				color: #000000;
				background-color: #F5F5F5;
				border: 1px groove #000000;
				
				-webkit-box-shadow: 0px 10px 10px -10px #979696;
				-moz-box-shadow: 0px 10px 10px -10px #979696;
				box-shadow: 0px 2px 2px 2px #979696;
            }
			
			td a
			{
				text-decoration: none;
				color:blue;
				word-wrap: Break-word;
			}
			
			th 
			{
				background-color: #7D7D7D;
				text-align: center;
				font-size: 14px;
				border: 1px groove #000000;
				word-wrap: Break-word;
				
				-webkit-box-shadow: 0px 10px 10px -10px #979696;
				-moz-box-shadow: 0px 10px 10px -10px #979696;
				box-shadow: 0px 2px 2px 2px #979696;
			}
			
			#dctable
			{
				width:98%;
				overflow-x: auto;
				overflow-y: auto;
				margin: 0px auto;
				margin: 0px auto;
				-webkit-box-shadow: 0px 10px 10px -10px #979696;
				-moz-box-shadow: 0px 10px 10px -10px #979696;
				box-shadow: 0px 2px 2px 2px #979696;
			}
			
			#cputable
			{
				width:80%;
				-webkit-box-shadow: 0px 10px 10px -10px #979696;
				-moz-box-shadow: 0px 10px 10px -10px #979696;
				box-shadow: 0px 2px 2px 2px #979696;
                
			}
			
			#container
			{
				width: 98%;
				background-color: #616060;
                margin: 0px auto;
				overflow-x:auto;
				margin-bottom: 20px;
			}
            
            #scriptexecutioncontainer
            {
				width: 80%;
				background-color: #616060;
				overflow-x:auto;
				margin-bottom: 30px;
				margin: auto;
			}
            
            #discovercontainer
            {
				width: 80%;
				background-color: #616060;
				overflow-x:auto;
				padding-top: 30px;
				margin-bottom: 30px;
				margin: auto;
			}

			#portsubcontainer
			{
				float: left;
				width: 48%;
				height: 400px;
				overflow-x:auto;
				overflow-y:auto;
			}
			#sysvolsubcontainer
			{
				float: right;
				width: 48%;
				height: 400px;
				overflow-x:auto;
				overflow-y:auto;
			}
			#disksubcontainer
			{
				float: left;
				width: 48%;
				height: 400px;
				overflow-x:auto;
				overflow-y:auto;
			}
            #servicessubcontainer
			{
				float: right;
				width: 48%;
				height: 400px;
				overflow-x:auto;
				overflow-y:auto;
			}
			#cputablecontainer{
				width:98%;
				margin: 0px auto;
				overflow-y: auto;
				overflow-x:auto;
				margin-bottom: 50px;
				height: 600px; 
			}
			#dctablecontainer{
				width:98%;
				margin: 0px auto;
				width: 100%;
				overflow-y: auto;
				overflow-x:auto;
				margin-bottom: 50px;
				height: 600px;
				
			}
			.error  
			{
				text-color: #FE5959;
				text-align: left;
			}

			#titleblock
			{
				display: block;
				float: center;
				margin-left: 25%;
				margin-right: 25%;
				width: 100%;
				position: relative;
				text-align: center
				background-image:
			}
			
			#header img {
			  float: left;
			  width: 190px;
			  height: 130px;
			  /*background-color: #fff;*/
			}

			.title_class 
			{
				color: #3B1400;
				text-shadow: 0 0 1px #F42121, 0 0 1px #0A8504, 0 0 2px white;
				font-size:58px;
				text-align: center;
			}
			.passed
			{
				background-color: #6CCB19;
                text-align: left;
                color: #000000;
			}
			.failed
			{
				background-color: #FA6E59;
				text-align: left;
                color: #000000;
				text-decoration: none;
			}
			#headingbutton
			{
				display: inline-block;
				padding-top: 8px;
				padding-bottom: 8px;
				background-color: #D9D7D7;
				font-size: 16px
				font-family: 'Trebuchet MS', Arial, Helvetica, sans-serif;
				font-weight: bold;
				color: #000;
				
				width: 12%;
				text-align: center;
				
				-webkit-box-shadow: 0px 1px 1px 1px #979696;
				-moz-box-shadow: 0px 1px 1px 1px #979696;
				box-shadow: 0px 1px 1px 1px #979696;
			}
			
			#headingtabsection
			{
				width: 96%;
				margin-right: 50px;
				margin-left: 55px;
				margin-bottom: 30px;
				margin-bottom: 50px;
			}
			
			#headingbutton:active
			{
				background-color: #7C2020;
			}

			#headingbutton:hover
			{
				background-color: #7C2020;
				color: #ffffff;
			}
			
			#headingbutton:hover
			{
				background-color: #ffffff;
				color: #000000;
			}
			 #headingbutton a
			{
				color: #000000;
				font-size: 16px;
				text-decoration: none;
				
			}
			 
			#header
			{
				width: 100%
				padding: 30px;
				text-align: center;
				color: #3B1400;
				color: white;
				text-shadow: 8px 8px 12px #000000;
				font-size:68px;
				background-color: #616060;
			}
			#headerdate
			{
				color: #ffffff;
				font-size:16px;
				font-weight: bold;
				margin-bottom: 5px;		
			}
			/* Tooltip container */
			.tooltip {
			  position: relative;
			  display: inline-block;
			  border-bottom: 1px dotted black; /* If you want dots under the hoverable text */
			}

			/* Tooltip text */
			.tooltip .tooltiptext {
			  visibility: hidden;
			  width: 180px;
			  background-color: black;
			  color: #fff;
			  text-align: center;
			  padding: 5px 0;
			  border-radius: 6px;
			 
			  /* Position the tooltip text - see examples below! */
			  position: absolute;
			  z-index: 1;
			}

			/* Show the tooltip text when you mouse over the tooltip container */
			.tooltip:hover .tooltiptext {
			  visibility: visible;
			  right: 105%; 
			}
		</style>
	</head>
	<body>
	    <div id=header>
            
            AD Health Check Report
            <br />
            <span id=headerdate>$DisplayDate</span>
        </div> 
        <!--<div id='toolbar'><a href='#' id='backToTop' onclick='HideTopButton()'>TOP</a></div>-->
        <section id=headingtabsection>
            <div id=headingbutton><a href='#Connectivity Status'>Connectivity Status</a></div>
			<div id=headingbutton><a href='#Sysvol'>SysVol and NetLogon</a></div>
            <div id=headingbutton><a href='#Disk Space'>Disk Space</span></a></div>
            <div id=headingbutton><a href='#Services'>Services</span></a></div>
            <div id=headingbutton><a href='#Performance Check'>Performance Check</a></div>
            <div id=headingbutton><a href='#DC Health Check'>DC Health Check</span></a></div>
            <div id=headingbutton><a href='#Script Execution Time'>Execution Details</span></a></div>
			<div id=headingbutton><a href='#DC Discovery details'>DC Discovery details</span></a></div>
        </section>
   "
Add-Content $HealthReport $header




#---------------------------------------------------------------------------------------------------------------------------
# Get all DCs in the Domain
#---------------------------------------------------------------------------------------------------------------------------


try 
{ 
      $Forest = [system.directoryservices.activedirectory.Forest]::GetCurrentForest()    
} 
catch 
{ 
      Write-output "Cannot connect to current forest."
      "<p><span class=logerror>Cannot connect to current forest.</span></p>" | Out-File -Append $Logfile
      Break;
} 

 
$Forest.domains | ForEach-Object {$_.DomainControllers} | ForEach-Object {
$DCList += $_.Name
}

if(!$DCList)
{
   Write-Host -ForegroundColor Yellow "No Domain Controller found. Run this solution on AD server. Please try again."
   "<p><span class=logerror>No Domain Controller found. Run this solution on AD server. Please try again.</span></p>" | Out-File -Append $Logfile   
   Stop-Transcript
   Break;
}

$total_machines_discovered = @($DCList)
$total_machines_discovered_count = $total_machines_discovered.Count
$total_machines_discovered_list = @()
foreach ($i in $total_machines_discovered) {
    $total_machines_discovered_list += "${i}<br />"
}
$total_machines_discovered_list
                                   
Write-Host -ForegroundColor White "`nList of Domain Controllers Discovered`n"
"<p><span class=loginfor>List of Domain Controllers Discovered</span></p>" | Out-File -Append $Logfile 

# List out all machines discovered in Log File and Console
foreach ($D in $DCList) 
{

Write-host "$D"
"<p><span class=loginfor>$D</span></p>" | Out-File -Append $Logfile
}

Add-Content $HealthReport $dataRow

Write-output "`n Starting Task Scheduling Part `n"
"<p><span class=loghighlight>Starting Task Scheudling Part</span></p>" | Out-File -Append $Logfile  





#---------------------------------------------------------------------------------------------------------------------------
# ADHealthCheck Directory Structure Check and Creation on Domain Controllers
#---------------------------------------------------------------------------------------------------------------------------

Write-Host -ForeGroundColor Cyan " `n ******************************************ADHealthCheck Directory Structure Check and Creation Part********************************************** `n "
"<p><span class=loginfo>******************************************ADHealthCheck Directory Structure Check and Creation Part**********************************************</span></p>" | Out-File -Append $Logfile

foreach($dc in $DCList)
{
   Write-Host -ForeGroundColor Cyan " `n **************************************************************************************************** `n "
   "<p><span>****************************************************************************************************</span></p>" | Out-File -Append $Logfile
   
   Write-Host -ForeGroundColor Cyan " `n Script will check directory structure on $dc `n "
   "<p><span class=loginfo>Script will check directory structure on $dc</span></p>" | Out-File -Append $Logfile

   $machinepath = "\\$dc\c$"
   
# Check for the availability of directory structre on each domain controller----------#   
   if (Test-Path "$machinepath\ADHealthCheck") 
   {
		#------------Replace the old Scrpt on each domain controller-----------#
		Remove-item "$machinepath\ADHealthCheck\AD_HealthCheck_V5_1.ps1" -Force -ErrorAction Continue
		Copy-Item -Path "$Dir\AD_HealthCheck_V5_1.ps1" "$machinepath\ADHealthCheck" -Force -ErrorAction Continue
		
		if((Test-Path "$machinepath\ADHealthCheck\log"))
		{
			#------------If Directory structure is present the remove the old log files present in the log directory----------#  
			$removeoldstructure = Remove-Item "$machinepath\ADHealthCheck\log\*" -Force -ErrorAction Continue
		}
		else
		{
			try 
            {
                New-Item "$machinepath\ADHealthCheck\log" -ItemType directory -Force -ErrorAction Stop
            }
            catch
            {
                
                $exceptionmessage = $_.Exception.Message
                Write-host -ForegroundColor Yellow "${dc}: $exceptionmessage" 
                "<p><span class=logerror>${dc}: $exceptionmessage</span></p>" | Out-File -Append $Logfile
                # Add this in failed machine list where directory creation failed.
                $machines_with_failed_directoy_creation += $dc
                # Remove this machine from dc list
                $DCList = $DCList -ne $dc
                Continue
            }       
		}
   }
   else
   {
		try
        {
            New-Item "$machinepath\ADHealthCheck" -ItemType directory -ErrorAction Stop
		    New-Item "$machinepath\ADHealthCheck\log" -ItemType directory -ErrorAction Stop		
        }
        catch
        {
                $exceptionmessage = $_.Exception.Message
                Write-host -ForegroundColor Yellow "${dc}: $exceptionmessage" 
                "<p><span class=logerror>${dc}: $exceptionmessage</span></p>" | Out-File -Append $Logfile
                # Add this in failed machine list where directory creation failed.
                $machines_with_failed_directoy_creation += $dc
                # Remove this machine from dc list
                $DCList = $DCList -ne $dc
                Continue
        }
   }
   Write-Host -ForegroundColor Cyan "`nNow Script will confirm the existence of files required"      
   
   #---------- Confirm if required files are copied to machine -------------------------------------#
   
   $HealthCheckDirExist = Test-Path "$machinepath\ADHealthCheck"
   $LogDirExist = Test-Path "$machinepath\ADHealthCheck\log"
   $HealthCheckScriptExist = Test-Path "$machinepath\ADHealthCheck\AD_HealthCheck_V5_1.ps1"
   
   
   if (($HealthCheckDirExist) -and ($LogDirExist) -and ($HealthCheckScriptExist))
   {
      Write-Host -ForegroundColor Green "All required files copied to machine: $dc`nNow Script will proceed further"
      "<p><span class=logsuccess>All required files copied to machine</span></p>" | Out-File -Append $Logfile  
   } 
   else 
   {
      Write-Host -ForegroundColor yellow "`n ADHealthCheck Directory Structure Creation Error in machine: $dc `n"
      Write-Host -ForegroundColor yellow "Health Check Dir Exist: $HealthCheckDirExist"
      Write-Host -ForegroundColor yellow "Log Dir Exist: $LogDirExist"
      Write-Host -ForegroundColor yellow "AD Health Check Script Exist: $HealthCheckScriptExist"

      "<p><span class=logerror>ADHealthCheck Directory Structure Creation Error in machine: $dc </span></p>" | Out-File -Append $Logfile
      "<p><span class=logerror>Health Check Dir Exist: $HealthCheckDirExist</span></p>" | Out-File -Append $Logfile
      "<p><span class=logerror>Log Dir Exist: $LogDirExist</span></p>" | Out-File -Append $Logfile
      "<p><span class=logerror>AD Health Check Script Exist: $HealthCheckScriptExist</span></p>" | Out-File -Append $Logfile

        $exceptionmessage = $_.Exception.Message
        Write-host -ForegroundColor Yellow "${dc}: $exceptionmessage" 
        "<p><span class=logerror>${dc}: $exceptionmessage</span></p>" | Out-File -Append $Logfile
        # Add this in failed machine list where directory creation failed.
        $machines_with_failed_directoy_creation += $dc
        # Remove this machine from dc list
        $DCList = $DCList -ne $dc
        Continue
   }

}


#---------------------------------------------------------------------------
# List out the machines in which Active Directory Sturcture can't be created
#---------------------------------------------------------------------------

[int]$machines_with_failed_directoy_creation_count = $machines_with_failed_directoy_creation.Count

if ($machines_with_failed_directoy_creation_count -ne 0)
{
    Write-Host -ForegroundColor Yellow "Active Health Directory structure Couldn't be Created in below machines"
    "<p><span class=logerror>Active Health Directory structure Couldn't be Created in below machines</span></p>" | Out-File -Append $Logfile

    foreach ($m in $machines_with_failed_directoy_creation)
    {
        Write-Host -ForegroundColor Yellow "`n $m `n"
        "<span class=logerror>$m</span><br />" | Out-File -Append $Logfile
        $machines_with_failed_directoy_creation_list += "$m<br />"
    }
}
else
{
    Write-Host -ForegroundColor Green "AD Directory Structure Created on all machines successfully"
    "<p><span class=logsuccess>AD Directory Created on all machines successfully</span></p>" | Out-File -Append $Logfile
}






#---------------------------------------------------------------------------------------------------------------------------
# Scheduled Task Creation/Execution on Domain Controllers
#---------------------------------------------------------------------------------------------------------------------------

Write-Host -ForeGroundColor Cyan " `n ******************************************Scheduled Task Creation/Execution Part********************************************** `n "
	"<p><span class=loginfo>******************************************Scheduled Task Creation/Execution Part**********************************************</span></p>" | Out-File -Append $Logfile


# Check if any domain controllers left
if($DCList.Count -eq 0) {
    Write-host -ForegroundColor Yellow "As no machines left script won't continue further"
    "<p><span class=logerror>As no machines left script won't continue further</span></p>" | Out-File -Append $Logfile
    Stop-Transcript
    Break;
}


foreach ($dc in $DCList)
{
	
   
	Write-host -foregroundcolor white "Machine: $dc"
	"<p><span class=loginfo>Machine: $dc</span></p>" | Out-File -Append $Logfile
   
   $task = schtasks /Query /S $dc /TN $taskname /fo list
  
   if($task)
   {   

      Write-host -foregroundcolor white "TASK Details"
      "<p><span class=loginfo>TASK Details:<br /> $task</span></p>" | Out-File -Append $Logfile
	  
      #------------Check Satus of task----------# 
      [string]$Name = $task | Select-String -Pattern "(TaskName:\s+\\)(.*)" | %{$_.Matches} | %{ $_.Groups[2]} | %{$_.Value}
      [string]$Status = $task | Select-String -Pattern "(Status:\s+)(.*)" | %{$_.Matches} | %{ $_.Groups[2]} | %{$_.Value}
      if($Status -eq 'Running')
      {
         Write-Host -ForegroundColor Green "`n Task $Name is already $Status `n"
      }
      else
      {
         Write-Host -ForegroundColor White "`n`n Task $Name on $dc is found with $Status Status `n`n`n "
         if ($Username) {
			$schtaskrun = schtasks /change /u $UserName /p $Password /S "${dc}"
            [string]$localSystem = $env:COMPUTERNAME
            if($dc -match $localSystem)
			{
               # Local System where Collator Script is running
               Write-Host -ForegroundColor White "$dc is a Local Machine"
			   # Run Task
               $schtaskrun = schtasks /run /TN "ADHealthCheck" /S "${dc}"
               Write-host "`n $schtaskrun `n"
               "<p><span class=loginfo>$schtaskrun</span></p>" | Out-File -Append $Logfile
            } 
			else 
			{
               # Remote System any domain controller
               Write-Host -ForegroundColor White "$dc is a Non Local System: Domain Controller System: $dc"
               # Run Task
               $schtaskrun = schtasks /run /TN "ADHealthCheck" /u $UserName /p $Password /S "${dc}"
               Write-host "`n $schtaskrun `n"
               "<p><span class=loginfo>$schtaskrun</span></p>" | Out-File -Append $Logfile
            }
         } 
         else 
         {
            $schtaskrun = schtasks /run /TN "ADHealthCheck" /S "${dc}"
            Write-host "`n $schtaskrun `n"
            "<p><span class=loginfo>$schtaskrun</span></p>" | Out-File -Append $Logfile
         }
      }
   }
   else
   {

       if($Username)
       {
            Write-Host -ForegroundColor Yellow "`n Error in Querying the Scheduled task on $dc. Script will now try to create the task `n"
            "<p><span class=logwarning>Error in Querying the Scheduled task on $dc. Script will now try to create the task</span></p>" | Out-File -Append $Logfile
         
            [datetime]$TaskTrigger = (Get-Date)
            $TaskTrigger = $TaskTrigger.AddSeconds(30)
            $StartTime = Get-Date -Date $TaskTrigger -Format HH:mm
            $StartDate = Get-Date -Date $TaskTrigger  -Format MM/dd/yyyy
            $Frequency = 'ONCE'
            $level = 'Highest'
            $TaskAction = "PowerShell.exe -ExecutionPolicy Bypass -NonInteractive -File C:\ADHealthCheck\AD_HealthCheck_V5_1.ps1"
	
	
            [string]$localSystem = $env:COMPUTERNAME
            $taskCreateOutput = schtasks /Create /S $dc /RU $UserName /RP $Password /RL $level /TN $taskName /TR $TaskAction /ST $StartTime /SD $StartDate /SC $Frequency /F  | Out-String
	
	        Write-Host "$StartTime `n Start Date: $StartDate `n Frequency: $Frequency `n Level: $level `n Task Action: $TaskAction `n Task Creation Output: $taskCreateOutput `n"
            "<p><span class=loginfo>$StartTime<br />Start Date: $StartDate<br />Frequency: $Frequency<br />Level: $level<br />Task Action: $TaskAction<br />Task Creation Output: $taskCreateOutput</span></p>" | Out-File -Append $Logfile

            if (!$taskCreateOutput) {
        
                $StartDate = Get-Date -Date $TaskTrigger  -Format dd/MM/yyyy
                $taskCreateOutput = schtasks /Create /S $dc /RU $UserName /RP $Password /RL $level /TN $taskName /TR $TaskAction /ST $StartTime /SD $StartDate /SC $Frequency /F  | Out-String
        
		        Write-host -ForegroundColor White "Reattempting Task Creation with different Start Date Format"
                "<p><span class=loginfo>Reattempting Task Creation with different Start Date Format</span></p>" | Out-File -Append $Logfile
                Write-Host "$StartTime `n Start Date: $StartDate `n Frequency: $Frequency `n Level: $level `n Task Action: $TaskAction `n Task Creation Output: $taskCreateOutput `n"
		        "<p><span class=loginfo>$StartTime<br />Start Date: $StartDate<br />Frequency: $Frequency<br />Level: $level<br />Task Action: $TaskAction<br />Task Creation Output: $taskCreateOutput</span></p>" | Out-File -Append $Logfile
            }

            # Confirm if task is Created
            $task = schtasks /Query /S $dc /TN $taskname /fo list

            # Check if Task is Created
            if($task) 
            {
                Write-host -ForegroundColor Green "Successfully Created Task on $dc"
                "<p><span class=logsuccess>Successfully Created Task on $dc</span></p>" | Out-File -Append $Logfile
                # ADHealthCheck Task is Created, now script will run the task
                if($dc -match $localsystem) 
                {
                    # Local System where Collator Script is running
                    Write-Host -ForegroundColor Cyan "`n Local System: $localSystem `n"
                    $schtaskrun = schtasks /run /TN $taskname /S "${dc}"
                    Write-host "`n $schtaskrun `n"
                    "<p><span class=loginfo>$schtaskrun</span></p>" | Out-File -Append $Logfile  
                } 
                else 
                {
                    # Remote System any domain controller
                    Write-Host -ForegroundColor Cyan "`nNot a Local System.Domain Controller System: $dc `n"
                    $schtaskrun = schtasks /run /TN $taskname /u $UserName /p $Password /S "${dc}"
                    Write-host "`n $schtaskrun `n"
                    "<p><span class=loginfo>$schtaskrun</span></p>" | Out-File -Append $Logfile
                }     
            }
            else
            {
		        #Script Was Unable to Create the tasks
		        Write-host -ForegroundColor Green "`n Script was unable to create task on $dc. Please Check manually. `n"
		        "<p><span class=logerror><p>`n Script was unable to create task on $dc. Please Check manually. `n</span></p>" | Out-File -Append $Logfile

		        # Remove this machine from dc list
		        $DCList = $DCList -ne $dc
    
		        # Add the machine to Task Creation Failed Machine
		        $machines_with_failed_task_creation += $dc
            }          
       }
       else
       {
            #No Credentials provided so Script cannot create a task
		    Write-host -ForegroundColor Yellow "No Credentials provided so Script cannot create a task`n"
		    "<p><span class=logerror><p>No Credentials provided so Script cannot create a task</span></p>" | Out-File -Append $Logfile
       }
   
   } # Initial Task found if else ends here
   
Write-Host -ForeGroundColor Cyan " `n Task Schedeling Operation is done on $dc `n "
"<p><span class=loginfo>Task Schedeling Operation is done on $dc</span></p>" | Out-File -Append $Logfile

}

Write-Host -ForegroundColor White "`n Task Scheduling Part is tried on all domain controllers `n"
"<p><span class=loginfo>Task Scheduling Part is tried on all domain controllers</span></p>" | Out-File -Append $Logfile

Write-Host -ForeGroundColor Cyan " `n **************************************************************************************************** `n "
"<p><span>****************************************************************************************************</span></p>" | Out-File -Append $Logfile




#---------------------------------------------------------------------------
# List out all machines where task run is successfull
#---------------------------------------------------------------------------
Write-Host -ForegroundColor White "`n Task scheduled successfully in below machines: `n"
"<p><span class=loginfo>Task scheduled successfully in below machines:</span></p>" | Out-File -Append $Logfile

foreach ($d in $DCList)
{
    Write-Host -ForegroundColor White "`n $d `n"
    "<p><span class=logsuccess>$d<br /></span>" | Out-File -Append $Logfile
}




#---------------------------------------------------------------------------
# List out the machines in which task can't be created
#---------------------------------------------------------------------------
Write-Host -ForegroundColor White "`n Machines in which Scheduled Task can't be created: `n"
"<p><span class=loginfo>Machines in which Scheduled Task can't be created:</span></p>" | Out-File -Append $Logfile
[int]$machines_with_failed_task_creation_count = $machines_with_failed_task_creation.Count
if ($machines_with_failed_task_creation_count -ne 0)
{
    Write-Host -ForegroundColor Yellow "Task Couldn't be Created in below machines"
    "<p><span class=logerror>Task Couldn't be Created in below machines</span></p>" | Out-File -Append $Logfile

    foreach ($t in $machines_with_failed_task_creation)
    {
        Write-Host -ForegroundColor Yellow "`n $t `n"
        "<span class=logerror>$t</span><br />" | Out-File -Append $Logfile
        $machines_with_failed_task_creation_list += "$t<br />"
    }
}
else
{
    Write-Host -ForegroundColor Green "Task Created/executed on all machines successfully"
    "<p><span class=logsuccess>Task Created/executed on all machines successfully</span></p>" | Out-File -Append $Logfile
}






#---------------------------------------------------------------------------
# TCP and UDP Port Checking Part
#---------------------------------------------------------------------------


# Check if any domain controllers left
if($DCList.Count -eq 0) {
    Write-host -ForegroundColor Yellow "As no machines left script won't continue further"
    "<p><span class=logerror>As no machines left script won't continue further</span></p>" | Out-File -Append $Logfile
    Stop-Transcript
    Break
}



# Start Container Div and Sub container div
$dataRow = "<div id=container><div id=portsubcontainer>"
$dataRow += "<table border=1px>
<caption><h2><a name='Connectivity Status'>Connectivity Status</h2></caption>"

$dataRow += "<tr>
<th>Host Name</th>
<th>Reachable</th>
<th>Up Time (Days)</th></tr>"

foreach($dc in $total_machines_discovered)
{
   $os = Get-WMIObject -class Win32_OperatingSystem -computer $dc  
   $uptime = (Get-Date) - ($os.ConvertToDateTime($os.lastbootuptime))
   $last_boot_up_time = $os.ConvertToDateTime($os.lastbootuptime)
   $UpTimeData = "$($Uptime.Days) days $($Uptime.Hours) hours"



$dataRow += "<tr>
<td class=bold_class>$dc</td>
<td class=passed>Success</td>
<td><div class=tooltip>$UpTimeData<span class=tooltiptext>Last Boot Up Time: $last_boot_up_time</span></div></td></tr>"
}
Add-Content $HealthReport $dataRow
Add-Content $HealthReport "</table></div>" # End Sub Container Div


#---------------------------------------------------------------------------
# SysVol and NetLogon Checking Part
#---------------------------------------------------------------------------

# Start Sub Container
$sysvoltable= "<Div id=sysvolsubcontainer><table border=1px>
   <caption><h2><a name='Sysvol'>SysVol and NetLogon</h2></caption>
            <th>Host Name</th>
            <th>SysVol Status</th>
            <th>NetLogon Status</th><thead><tbody>
			<tr>
        "

Add-Content $HealthReport $SysVoltable

foreach ($machine in $total_machines_discovered)
{
	
   #---------------------------------------
   # SysVol Analysis
   #---------------------------------------  
   $sysvolrow = "<td class=bold_class>$machine</td>";
   $sysvolpath = "\\"+"$machine" + "\" + "SYSVOL"
   if(test-path $sysvolpath)
   {
        $sysvolrow += "<td class=passed><div class=tooltip>Sysvol Exists<span class=tooltiptext>$sysvolpath Exists</span></div></td>"
        #$sysvolrow += "<td class=passed>Sysvol Exists</td>";
        Add-Content $HealthReport $sysvolrow

        Write-Host -ForegroundColor Green "SysVol Analysis on ${machine}: $sysvolpath Exists"
        "<p><span class=logsuccess>SysVol Analysis on $machine :<br />$sysvolpath Exists</span></p>" | Out-File -Append $Logfile
    } 
    else
    {
        $sysvolrow += "<td class=failed><div class=tooltip>Sysvol does not Exists<span class=tooltiptext>$sysvolpath does not  Exists</span></div></td>"
        # $sysvolrow += "<td class=failed>Sysvol does not Exist</td>";
        Add-Content $HealthReport $sysvolrow

        Write-Host -ForegroundColor Red "`r SysVol Analysis on ${machine}: $sysvolpath does not  Exists"
        "<p><span class=logerror>SysVol Analysis on $machine :<br />$sysvolpath does not  Exists</span></p>" | Out-File -Append $Logfile
    }

    
   
   #----------------------------------------
   # Netlogon Analysis
   #----------------------------------------

   $Netlogonpath = "\\"+"$machine" + "\" + "NETLOGON"

   if(test-path $Netlogonpath)
   {
      $sysvolrow = "<td class=passed><div class=tooltip>Netlogon Exists<span class=tooltiptext>$Netlogonpath Exists</span></div></td></tr>"
      #$sysvolrow = "<td class=passed>Netlogon Exists</td></tr>";
      Add-Content $HealthReport $sysvolrow
      
      Write-Host -ForegroundColor Green "`r Netlogon Analysis on $machine `r $Netlogonpath Exists"
        "<p><span class=logsuccess>Netlogon Analysis on $machine :<br />$Netlogonpath Exists</span></p>" | Out-File -Append $Logfile   
   }
   else
   {
      $sysvolrow = "<td class=failed><div class=tooltip>Netlogon does not Exists<span class=tooltiptext>$Netlogonpath does not Exists</span></div></td></tr>"
      #$sysvolrow = "<td class=failed>Netlogon does not Exists</td></tr>";
      Add-Content $HealthReport $sysvolrow

      Write-Host -ForegroundColor Red "`r Netlogon Analysis on $machine `r $Netlogonpath  does not Exists"
      "<p><span class=logsuccess>Netlogon Analysis on $machine :<br />$Netlogonpath  does not Exists</span></p>" | Out-File -Append $Logfile
   }

}

Add-Content $HealthReport "</tbody></table></div></div>" # End Sub Container Div and Container Div


#-----------------------------------------------------------------------------------------------------------
# Script will now Check the status of schduled tasks
#-----------------------------------------------------------------------------------------------------------


Write-output "`n Starting Task Status Checking Part Now `n"
"<p><span class=loghighlight>Starting Task Status Checking Part Now</span></p>" | Out-File -Append $Logfile  

$PendingMachines = $DCList

# Script will keep checking pending tasks till MaxCheckLimit is not reached

for ([int]$i=0;$i -le $MaxCheckLimit; $i++) 
{
   Write-host " `n Starting Task Status Check Iteration-$i `n "
   "<p><span class=loginfo>****Starting Task Status Check Iteration-${i}****</span></p>" | Out-File -Append $Logfile
   
   # Write-host " `n Pending Machines:$PendingMachines `n "
   "<p><span class=loginfo>Pending Machines</span></p>" | Out-File -Append $Logfile
   foreach ($P in $PendingMachines) 
   {
    "$P<br />" | Out-File -Append $Logfile
   }

   if([int]$PendingMachines.Count -eq 0)
   {
		$i = $MaxCheckLimit
		Write-Host -ForegroundColor Green "`n$taskname task is complete on each machine"
   } 
   else
   {
		# Call Check Task Status function
		CheckTaskStatus $DCList $PendingMachines
		if([int]$PendingMachines.Count -eq 0)
		{
		$i = $MaxCheckLimit
		Write-Host -ForegroundColor Green "`n$taskname task is complete on each machine"
		} 
		else
		{
		Write-Host -ForegroundColor Yellow "Still tasks are running on some machines.`n"
		Write-Host -ForegroundColor Yellow "Script is now going to sleep for $WaitTime Seconds and then will recheck`n"

		Start-Sleep -s $WaitTime
      
      }
   }
}

# Display further message that tasks are complete or not after maximum checks

if([int]$PendingMachines.Count -eq 0)
{
   Write-Host -ForegroundColor Green " `n All Tasks are complete. Script will proceed further `n "
   "<p><span class=logsuccess>All Tasks are complete. Script will proceed further</span></p>" | Out-File -Append $Logfile
} 
else 
{
   Write-Host -ForegroundColor Yellow " `n Some Tasks are still not complete but as Max Check Limit is reached so script will proceed further `n "
   "<p><span class=logsuccess>Some Tasks are still not complete but as Max Check Limit is reached so script will proceed further</span></p>" | Out-File -Append $Logfile
}





#-----------------------------------------------------------------------------------------------------------
#Check for File Existence in each domain controller
#-----------------------------------------------------------------------------------------------------------


if ($DCLIST.Count -eq 0) 
{
    Write-host -foregroundcolor Yellow "`n No Domain Controller Left to checked `n"
	"<p><span class=logerror>No Domain Controller Left to checked</span></p>" | Out-File -Append $Logfile
    Break;
}



foreach($dc in $DCList)
{ 
	if(!(CheckandCopyFile -m $dc))
	{
		$FileRecheckMachines += $dc       
	}
	else{
		Write-host -foregroundcolor Green "File found for $machine. `n"
		"<span class=logsuccess>File found for $machine.</span><br />" | Out-File -Append $Logfile
	}
}

if($FileRecheckMachines.Count -eq 0)
{
	Write-host -foregroundcolor green "`n File is found for each machines. Script will proceed further."
	"<p><span class=logsuccess>File is found for each machines. Script will proceed further.</span></p>" | Out-File -Append $Logfile
        
}
else
{
	Write-host -foregroundcolor Yellow "`n File not found on some machine. Script will wait for $FileNotFoundWaitTime seconds and then will recheck. `n"
	"<p><span class=logwarning>File not found on some machine. Script will wait for $FileNotFoundWaitTime seconds and then will recheck.</span></p>" | Out-File -Append $Logfile
	
	Start-Sleep -s $FileNotFoundWaitTime
	
	foreach($m in $FileRecheckMachines)
	{   
		if(!(CheckandCopyFile -m $m))
		{
			Write-host -ForegroundColor Yellow "`n File has still not came on $machine. Script will now skip this machine"
			"<span class=logerror>File has still not came on $machine. Script will now skip this machine</span><br />" | Out-File -Append $Logfile
                
            # As file is not found in final attempt so remove it from DCLIst i.e. List of machines to be checked.
            $DCList = $DCList -ne $m

            # Add it to file not found list of machines
            $machines_with_file_not_found += $m


		}
	}	
}




# List out the machines where file is not found

$machines_with_file_not_found_count = $machines_with_file_not_found.Count

if ($machines_with_file_not_found_count -ne 0) 
{
    Write-Host -ForegroundColor Yellow "File is not found for below machines:"
    "<p><span class=logerror>File is not found for below machines:</span></p>" | Out-File -Append $Logfile

    foreach ($m in $machines_with_file_not_found) 
    {
        Write-Host -ForegroundColor Yellow "$m `n"
        "<span class=logerror>$t</span><br />" | Out-File -Append $Logfile
        $machines_with_file_not_found_list += "$m<br />"
    }
}



# Check if any domain controllers left
if($DCList.Count -eq 0) {

    Write-host -ForegroundColor Yellow "As no machines left script won't continue further"
    "<p><span class=logerror>As no machines left script won't continue further</span></p>" | Out-File -Append $Logfile
    Stop-Transcript
    Break

} else {

    Write-host -ForegroundColor Green "List of Successfull Machines Left"
    "<p><span class=logsuccess>List of Successfull Machines Left</span></p>" | Out-File -Append $Logfile
    foreach ($d in $DCList) {
        Write-host "$d`n"
        "<p><span>$d</span></p>" | Out-File -Append $Logfile
    }
}






#-----------------------
# DISK SPace Analysis
#-----------------------

#Start Container and Sub Container Div
$DiskdataRow = "<div id=container><div id=disksubcontainer><table border=1px>
            <caption><h2><a name='Disk Space'>Disk Space</h2></caption>
            <th>Host Name</th>
            <th>Name</th>
            <th>Free</th>
            <th>Total</th>
            <th>Free%</th>
            "

Add-Content $HealthReport $DiskdataRow


foreach($machine in $DCList)
{   
   Write-Host -ForegroundColor Cyan "`rMachine: $machine `r"

   $resultfile="$Dir\Collator\log\"+ $machine.Split(".")[0].trim() + "_adhealthcheck.txt"

   if(Test-Path $resultfile) 
   {
      $Disks = @()
      $DChealth= Get-Content $resultfile
      [int]$Diskstartline=(($DChealth | Select-String -pattern "Disk Test Starts" | Select-Object LineNumber).LineNumber+4)
      [int]$DiskEndline=(($DChealth | Select-String -pattern "Disk Test Ends" | Select-Object LineNumber).LineNumber-2)

      for($i=$Diskstartline; $i -lt $DiskEndline; $i+=1)
      {
         $Disks +=$DCHealth | Select-Object -Index $i
      }
      [int]$DiskRowSpan = ($DiskEndline - $Diskstartline)
   
      $DiskdataRow = "<tr>
               <td class=bold_class rowspan='$DiskRowSpan'>$machine</td>"
		   Add-Content $HealthReport $DiskdataRow
      foreach($disk in $Disks)
      {
         $disktext=$Disk.split("") | ?{$_ -ne ""}
         if([int]$disktext[3] -lt $diskspacethresold)
         {
            $class = 'failed'
         } else {
            $class = 'passed'
         }
         $name=$disktext[0] -replace "\\$",''
         $total=$disktext[1] -replace ",","."
         $FreeSpace=$disktext[2] -replace ",","."
         $FreeSpaceCent="$($disktext[3])" + " $($disktext[4])"

         $DiskdataRow="<td>$Name</td>
               <td>$Total</td>
               <td>$FreeSpace</td>
               <td class=$class>$FreeSpaceCent</td>
            </tr>"
            Add-Content $HealthReport $DiskdataRow
      }
   }
   else
   {
            $DiskdataRow = "<tr bgcolor=""#FF0000""><td class=bold_class>$machine</td><td class=failed colspan=3>File Not Found</td></tr>"   
   }

}
Add-Content $HealthReport "</table></Div>" #End Sub Container

#---------------------------------------------------------------------------------------------------------------------------------------------
# Services Analysis
#---------------------------------------------------------------------------------------------------------------------------------------------

# Start Sub Container
$dataRow = "<Div id=servicessubcontainer><table border=1px>
   <caption><h2><a name='Services'>Services</h2></caption>
            <th>Host Name</th>
            <th>Status</th>
         "
Add-Content $HealthReport $dataRow

foreach($machine in $DCList)
{
   $resultfile="$Dir\Collator\log\"+ $machine.Split(".")[0].trim() + "_adhealthcheck.txt" 

   if(Test-Path $resultfile) 
   {
      $DChealth=Get-Content $resultfile
      $Services = @()
      #---------------------------------------------------------------------------------------------------------------------------------------------
      # SVC Analysis
      #---------------------------------------------------------------------------------------------------------------------------------------------

      [int]$Svcstartline=(($DChealth | Select-String -pattern "Service Test Starts" | Select-Object LineNumber).LineNumber+2)
      [int]$SvcEndline=(($DChealth | Select-String -pattern "Service Test Ends" | Select-Object LineNumber).LineNumber-1)
      [int]$PerfEndline=(($DChealth | Select-String -pattern "Performance Test Ends" | Select-Object LineNumber).LineNumber-3)
      for($i=$Svcstartline; $i -lt $SvcEndline; $i+=1)
      {
         $Services +=$DCHealth | Select-Object -Index $i
      }
      $svcdataRow=@()
      $svcfailed=@()
      if($Services | Select-String -Pattern "All services are up")
      {
         $svcdataRow = "<tr><td class=bold_class>$machine</td><td class=passed>All Services are up</td></tr>"
      }
      else
      {
         foreach($obj in $services)
         {
            #$svcfailed += ($obj -split ',')[1]
			$svcfailed += "$(($obj -split ',')[1])<br/>"
         }
		 $svcdataRow = "<tr><td class=bold_class>$machine</td><td class=failed><div class=tooltip>Failed<span class=tooltiptext>Services Failed: $svcfailed</span></div></td></tr>"
      }

   }
   else
   {
            $svcdataRow = "<tr><td class=bold_class>$machine</td><td class=failed colspan=3>File Not Found</td></tr>"   
   }
   Add-Content $HealthReport $svcdataRow
}

Add-Content $HealthReport "</table></Div></Div>" # End Sub Container and Container Div

#-----------------------------
# Performance Check Analysis
#-----------------------------

$pdataRow ="<div id=cputablecontainer><table>
                <caption><h2><a>Performance Check</h2></caption>
                     <th>DC Name</th>
                     <th>CPU Processor Time<BR><=80 %</th>
                     <th>CPU Priviledge Time<BR><=30 %</th>
                     <th>RAM Available MB<BR>>=100 MB</th>
                     <th>RAM in use<BR><=80 %</th>
                     <th>RAM Pages/sec<BR><1000</th>
                     <th>RAM Page Faults/sec<BR><=2500</th>
                     <th>AD Bind Time<BR><=50 ms</th>
                     <th>DISK Read/sec<BR><=20 ms</th>
                     <th>DISK Transfer/sec<BR><=20 ms</th>
                     <th>DISK Write/sec<BR><=20 ms</th>
                     <th>DISK Queue Length<BR><=2</th>
                     <!--<th>DISK Time %<BR><=50 %</th>-->
					"

Add-Content $HealthReport $pdataRow

foreach($machine in $DCList)
{

   $Performance = @()
   $prfdataRow=@()
   $resultfile="$Dir\Collator\log\"+ $machine.Split(".")[0].trim() + "_adhealthcheck.txt"

   $pdataRow = "<tr><td class=bold_class>$machine</td>"
   if(Test-Path $resultfile)
   {

   $DChealth = Get-Content $resultfile

      #-------------------------------
      # Perf Analysis
      #-------------------------------

      [int]$Perfstartline=(($DChealth | Select-String -pattern "Performance Test Starts" | Select-Object LineNumber).LineNumber+1)
      [int]$PerfEndline=(($DChealth | Select-String -pattern "Performance Test Ends" | Select-Object LineNumber).LineNumber-3)
      for($i=$Perfstartline; $i -lt $PerfEndline; $i+=1)
      {
         $Performance +=$DCHealth | Select-Object -Index $i
      }
      $res=$null
      $Percent = '%'
	  Add-Content $HealthReport $pdataRow
      foreach($perf in $Performance)
      {
         $res = $perf.split("=")
         $resout=$null
      
         if ($res[1] -match "CPU") 
         {
         $res[1] = $res[1] -replace "CPU.*",""
         }
         elseif($res[1] -match "Disk")
         {
         $res[1] = $res[1] -replace "Disk.*",""
         }
         elseif($res[1] -match "Memory")
         {
         $res[1] = $res[1] -replace "Memory.*",""
         }
         "Res1: $res[1]"
         $resout=[math]::Round($($res[1]),3)
         if($res[0] -eq "CPU | \processor(_total)\% processor time "){if($res[1] -le 80){$pdataRow = "<td>$resout %</td>";Add-Content $HealthReport $pdataRow}Else{$pdataRow = "<td>$resout %</td>";
         Add-Content $HealthReport $pdataRow}}
         elseif($res[0] -eq "CPU | \Processor(_total)\% Privileged Time "){if($res[1] -le 30){$pdataRow = "<td>$resout %</td>"
         Add-Content $HealthReport $pdataRow}Else{$pdataRow = "<td>$resout %</td>";
         Add-Content $HealthReport $pdataRow}}
         elseif($res[0] -eq "Memory | \Memory\Available MBytes "){if($res[1] -ge 100){$pdataRow = "<td>$resout MB</td>";Add-Content $HealthReport $pdataRow}else{$pdataRow = "<td>$resout MB</td>";Add-Content $HealthReport $pdataRow}}
         elseif($res[0] -eq "Memory | \Memory\% Committed Bytes In Use "){if($res[1] -le 80){$pdataRow = "<td>$resout %</td>";Add-Content $HealthReport $pdataRow}else{$pdataRow = "<td>$resout %</td>";Add-Content $HealthReport $pdataRow}}
         elseif($res[0] -eq "Memory | \Memory\Pages/sec "){if($res[1] -lt 1000){$pdataRow = "<td>$resout </td>";Add-Content $HealthReport $pdataRow}else{$pdataRow = "<td>$resout </td>";Add-Content $HealthReport $pdataRow}}
         elseif($res[0] -eq "Memory | \Memory\Page Faults/sec "){if($res[1] -le 2500){$pdataRow = "<td>$resout </td>";Add-Content $HealthReport $pdataRow}else{$pdataRow = "<td>$resout </td>";Add-Content $HealthReport $pdataRow}}
         elseif($res[0] -eq "Directory Services | LDAP Bind Time "){if($res[1] -le 50){$pdataRow = "<td>$resout ms</td>";Add-Content $HealthReport $pdataRow}else{$pdataRow = "<td>$resout ms</td>";Add-Content $HealthReport $pdataRow}}
         elseif($res[0] -eq "Disk | \PhysicalDisk(_total)\Avg. Disk sec/Read "){if($res[1] -le 20){$pdataRow = "<td>$resout ms</td>";Add-Content $HealthReport $pdataRow}else{$pdataRow = "<td>$resout ms</td>";Add-Content $HealthReport $pdataRow}}
         elseif($res[0] -eq "Disk | \PhysicalDisk(_total)\Avg. Disk sec/Transfer "){if($res[1] -le 20){$pdataRow = "<td>$resout ms</td>";Add-Content $HealthReport $pdataRow}else{$pdataRow = "<td>$resout ms</td>";Add-Content $HealthReport $pdataRow}}
         elseif($res[0] -eq "Disk | \PhysicalDisk(_total)\Avg. Disk sec/Write "){if($res[1] -le 20){$pdataRow = "<td>$resout ms</td>";Add-Content $HealthReport $pdataRow}else{$pdataRow = "<td>$resout ms</td>";Add-Content $HealthReport $pdataRow}}
         elseif($res[0] -eq "Disk | \PhysicalDisk(_total)\Avg. Disk Queue Length "){if($res[1] -le 2){$pdataRow = "<td>$resout </td>";Add-Content $HealthReport $pdataRow}else{$pdataRow = "<td>$resout </td>";Add-Content $HealthReport $pdataRow}}
         <#elseif($res[0] -eq "Disk | \PhysicalDisk(_total)\% Disk Time ")
         {
         if($res[1] -le 50)
         {
            $pdataRow = "<td>$resout</td></tr>";
            Add-Content $HealthReport $pdataRow
         }
         else
         {
            $pdataRow = "<td>$resout</td></tr>";
            Add-Content $HealthReport $pdataRow
         }
         }#>
      }
   }
   else
   {
        $pdataRow += "<td class=failed colspan=11>File Not Found</td></tr>"
   }
}
Add-Content $HealthReport "</table></div>"

#-------------------------------------
# DC Check Analysis
#-------------------------------------

$dcdataRow = "<div id=dctablecontainer><table id=dctable>
   <caption><h2><a name='DC Health Check'>DC Health Check</h2></caption>
        <thead>
		<th>Host Name</th>
		<th>DNS Diagnostic Check</th>
		<th>DC Diagnostic Test</th>
		<th>DC Backup</th>
		<th>Replication Monitoring</th>
        </thead>
	        "

Add-Content $HealthReport $dcdataRow
foreach($machine in $DCList)
{
   Write-Host -ForegroundColor Cyan "Machine $Machine"
   $resultfile="$Dir\Collator\log\"+ $machine.Split(".")[0].trim() + "_adhealthcheck.txt"


   
   $dcdataRow = "<tr><td class=bold_class>$machine</td>"
   Add-Content $HealthReport $dcdataRow
   #-----------------------------------------
   # DNS Analysis
   #-----------------------------------------

   if(Test-Path $resultfile) 
   {
      $dataRowDNS = @()
      $DcdiagDNS = @()
      $DChealth=Get-Content $resultfile
      [int]$DNSstartline=(($DChealth | Select-String -pattern "DNSDIAG Starts" | Select-Object LineNumber).LineNumber+1)
      [int]$DNSEndline=(($DChealth | Select-String -pattern "DNSDIAG Ends" | Select-Object LineNumber).LineNumber-3)
      for($i=$DNSstartline; $i -lt $DNSEndline; $i+=1)
      {
         $DcdiagDNS +=$DCHealth | Select-Object -Index $i
      }
   
      If (($DcdiagDNS -eq $Null) -or ($DcdiagDNS | Select-String -pattern "Ldap search capabality attribute search failed on server"))
      {
         $dataRowDNS = "<td class=failed>Command failed to run</td>"
         Add-Content $HealthReport $dataRowDNS
      }
      Else
      {
   
         $TestLine = 0

         $TestLine = (($DcdiagDNS | Select-String -pattern "Summary of DNS test results:" | Select-Object LineNumber).linenumber) + 6
         $TextObject = @($DcdiagDNS);
         $RqStatusLine = $($TextObject[$TestLine]) 
         
         $Auth = ($RqStatusLine.split()| where {$_})[1]
         $Basc = ($RqStatusLine.split()| where {$_})[2]
         $Forw = ($RqStatusLine.split()| where {$_})[3]
         $Del = ($RqStatusLine.split()| where {$_})[4]
         $Dyn = ($RqStatusLine.split()| where {$_})[5]
         $RReg = ($RqStatusLine.split()| where {$_})[6]
         $Ext = ($RqStatusLine.split()| where {$_})[7]

         $FailedDCDNSTest = @()
         if($Auth -eq "PASS" -or $Auth -eq "n/a"){}Else{$FailedDCDNSTest += "TEST: Authentication;TEST: Basic;Summary of DNS test results:"; }
         if($Basc -eq "PASS" -or $Basc -eq "n/a"){}Else{$FailedDCDNSTest += "TEST: Basic;TEST: Forwarders/Root hints;Summary of DNS test results:"}
         if($Forw -eq "PASS" -or $Forw -eq "n/a"){}Else{$FailedDCDNSTest += "TEST: Forwarders/Root hints;TEST: Delegations;Summary of DNS test results:"}
         if($Del -eq "PASS" -or $Del -eq "n/a"){}Else{$FailedDCDNSTest += "TEST: Delegations;TEST: Dynamic update;Summary of DNS test results:"}
         if($Dyn -eq "PASS" -or $Dyn -eq "n/a"){}Else{$FailedDCDNSTest += "TEST: Dynamic update;TEST: Records registration;Summary of DNS test results:"}
         if($RReg -eq "PASS" -or $RReg -eq "n/a"){}Else{$FailedDCDNSTest += "TEST: Records registration;Summary of test results;Summary of DNS test results:"}
      
         $DcdiagDNSCount = 0
      
         $DcdiagDNSCount = $FailedDCDNSTest.count
         $dataRowDNS = "<td class=failed>"
         if($DcdiagDNSCount -gt 0)
         {
         
            foreach($Lines in $FailedDCDNSTest)
            {
               $FailRows = @();
               $FromLine = ""
               $ToLine = ""
               $from = 0
               $to = 0

               $FromLine = ($Lines.split(";")[0])
               $ToLine = ($Lines.split(";")[1])
               $ToLine1 = ($Lines.split(";")[2])
				            
               [int]$from =  ($DcdiagDNS | Select-String -pattern $FromLine | Select-Object LineNumber).LineNumber
               [int]$to   =  ($DcdiagDNS  | Select-String -pattern $ToLine | Select-Object LineNumber).LineNumber
            
               if($to -le 0)
               {
            	   [int]$to   =  ($DcdiagDNS  | Select-String -pattern $ToLine1 | Select-Object LineNumber).LineNumber
               }
            
               If($from -ge '1'){$from = $from - 1}
               If($to -ge '2'){$to  = $to - 2}
							
               $DcdiagDNS | Select-Object -Index ("$from".."$to") | Foreach-Object {$FailRows = $FailRows + "<br>" + ($_).replace("'","").replace('"',"")}
               $dataRowDNS += "<a href='javascript:void(0)' onclick=""Powershellparamater('"+ $FailRows +"')"">$FromLine</a> | "
                        
            }
            $dataRowDNS = $dataRowDNS.Trim() -replace "\|$", ""
            $dataRowDNS +="</td>"
            Add-Content $HealthReport $dataRowDNS
         }
         else
         {
           $dataRowDNS = "<td class=passed>All DNSDiag tests passed</td>"
   	       Add-Content $HealthReport $dataRowDNS
         }
      }
   }
   

   #--------------------------------------
   # DC Analysis
   #--------------------------------------

   if(Test-Path $resultfile) 
   {
      $dataRow1 = @()
	  $DcdiagRowspan = 0
	  $Dcdiag = @()
      $AllFailedDCDiags = @()
      [int]$DCstartline=(($DChealth | Select-String -pattern "DCDIAG Starts" | Select-Object LineNumber).LineNumber+1)
      [int]$DCEndline=(($DChealth | Select-String -pattern "DCDIAG Ends" | Select-Object LineNumber).LineNumber-3)
      
      for($i=$DCstartline; $i -lt $DCEndline; $i+=1)
      {
         $DCDiag +=$DCHealth | Select-Object -Index $i
      }

	  
      If (($Dcdiag -eq $Null) -or ($Dcdiag | Select-String -pattern "Ldap search capabality attribute search failed on server"))
      {
         $dataRow1 = "<td class=failed>Command failed to run</td>"
         Add-Content $HealthReport $dataRow1
      }
      else
      {
	     $failedTestList = @( $Dcdiag |select-string -pattern "failed test")
         foreach($Names in $failedTestList)
         {
            $TestName = ($Names -split "test")[1]
  	         $AllFailedDCDiags += $TestName.Trim()
   	     }

         $DcdiagRowspan = $AllFailedDCDiags.count
         $dataRow1 = "<td class=failed><div>"
         if($DcdiagRowspan -gt 0)
         {
            foreach($TestName in $AllFailedDCDiags)
            {
               $TestName = ($TestName).trim()
               $FailedRecords = @()
               $from = 0
               $to   = 0
               $FromLine = ""
               $ToLine = ""
			   $TestName1 = ""
               $FromLine = "Starting test: $TestName"
               $ToLine = "failed test $TestName"

               [int]$from =  (($Dcdiag | Select-String -pattern $FromLine | Select-Object LineNumber).LineNumber)-1
               [int]$to   =  ($Dcdiag  | Select-String -pattern $ToLine | Select-Object LineNumber).LineNumber

               $Dcdiag | Select-Object -Index ("$from".."$to") | Foreach-Object {$FailedRecords = $FailedRecords + "<br>" + ($_).replace('"',"").replace("'","")}
               for($i=$from; $i -lt $to; $i+=1)
               {
                  $DCDiag +=$DCHealth | Select-Object -Index $i
               }
               $dataRow1 += "<a href='javascript:void(0)' onclick=""Powershellparamater('"+ $FailedRecords +"')""><span>test: $TestName</span></a> | "
   		   
            }
            $dataRow1 = $dataRow1.Trim() -replace "\|$", ""
            $dataRow1 +="</div></td>"
            Add-Content $HealthReport $dataRow1
        }
        else
        {
            $dataRow1 = "<td class=passed>All DC tests passed</td>"	    
            $dataRow1 +="</div></td>"
            Add-Content $HealthReport $dataRow1
            
         }
         
      }
   }


   #-----------------------------------------
   # DC BackUp Analysis
   #-----------------------------------------
   if(Test-Path $resultfile) 
   {
      $dcbkupcell = ''
      $DCBKP = ''
      [int]$BKPstartline=(($DChealth | Select-String -pattern "Backup Test Starts" | Select-Object LineNumber).LineNumber+1)
      [int]$BKPEndline=(($DChealth | Select-String -pattern "Backup Test Ends" | Select-Object LineNumber).LineNumber-2)
      for($i=$BKPstartline; $i -lt $BKPEndline; $i+=1)
      {
         $DCBKP +=$DCHealth | Select-Object -Index $i
      }
      if($DCBKP | Select-String -pattern "Test is Success")
      {
         $dcdataRow = "<td class=passed>$($DCBKP.split("|")[1])</td>";
         Add-Content $HealthReport $dcdataRow
      }
      elseif($DCBKP | Select-String -pattern "Test Failed")
      {
         $bkupresult = $($DCBKP.split("|")[1]) -replace "Test Failed",""
         $dcdataRow = "<td class=failed>$bkupresult</td>";
         Add-Content $HealthReport $dcdataRow
      }
      else
      {
         $dcdataRow = "<td class=failed>Never Backed up</td>";
         Add-Content $HealthReport $dcdataRow
      }
   }
   else
   {
        $dcdataRow = "<td class=failed>File Not Found</td>"
        Add-Content $HealthReport $dcdataRow 
   }


   #----------------------------------------------
   # DC Replication Analysis
   #----------------------------------------------
   if(Test-Path $resultfile) 
   {
      #$DCREPL = @()
      $dcdataRow = @()
      $finalrepl = @()
      [int]$REPLstartline=(($DChealth | Select-String -pattern "Replication Test Starts" | Select-Object LineNumber).LineNumber+1)
      [int]$REPLEndline=(($DChealth | Select-String -pattern "Replication Test Ends" | Select-Object LineNumber).LineNumber-2)
      for($i=$REPLstartline; $i -lt $REPLEndline; $i+=1)
      {
         $DCREPL +=$DCHealth | Select-Object -Index $i
      } 
   
      if($DCREPL | Select-String -pattern "AD Replication is a success")
      {
         $dcdataRow = "<td class='passed'>AD Replication is Clear</td></tr>";
         Add-Content $HealthReport $dcdataRow
      }
      else
      {
         $n=0
         $outvar=$null
         $seperator="`r"
         $repls=$DCREPL.Split($seperator)
         $repls.Count
         foreach($repl in $repls)
         {
            if($repl -like "showrepl_COLUMNS*"){break}
            $n++
         }

         $outvar = "<table><tr><td>DestSite</td><td>DestDC</td><td>Partition</td><td>SrcSite</td><td>SrcDC</td><td>Failures</td><td>FailureTime</td><td>SuccessTime</td><td>FailureCode</td></tr>"
         $finalrepl +=$outvar
         while($n -lt $repls.count)
         {
            $destsite=$repls[$n+1].split(':')[1];$destdc=$repls[$n+2].split(':')[1];$partition=$repls[$n+3].split(':')[1];$SrcSite=$repls[$n+4].Split(':')[1];$srcDC=$repls[$n+5].Split(':')[1];$failures=$repls[$n+7].Split(':')[1];$FailureTime = $($repls[$n+8].Split(':')[1]+"-"+$repls[$n+8].Split(':')[2];);$SuccessTime = $($repls[$n+9].Split(":")[1]+"-"+$repls[$n+9].Split(":")[2]);$FailureCode = $repls[$n+10].split(":")[1]
            $outvar ="<tr><td>$destsite</td><td>$destdc</td><td>$partition</td><td>$srcsite</td><td>$srcdc</td><td>$failures</td><td>$FailureTime</td><td>$SuccessTime</td><td>$FailureCode</td></tr>"
            $finalrepl+=$outvar
            $n=$n+12
            $n
         }
         $finalrepl +="</table>"
         $dcdataRow = "<td class=failed><a href='javascript:void(0)' target=_blank onclick=""Powershellparamater('"+ $finalrepl +"')"">Failed</a></td></tr>"
         Add-Content $HealthReport $dcdataRow
      }
   }

}
Add-Content $HealthReport "</table></div>"

Stop-Transcript



#---------------------------------------------------------------------------------------------------------------------------------------------
# Script Execution Time
#---------------------------------------------------------------------------------------------------------------------------------------------
$myhost = $env:COMPUTERNAME

$ScriptExecutionRow = "<div id=scriptexecutioncontainer><table>
   <caption><h2><a name='Script Execution Time'>Execution Details</h2></caption>
      <th>Start Time</th>
      <th>Stop Time</th>
		<th>Days</th>
      <th>Hours</th>
      <th>Minutes</th>
      <th>Seconds</th>
      <th>Milliseconds</th>
      <th>Script Executed on</th>
	</th>"

# Stop script execution time calculation
$sw.Stop()
$Days = $sw.Elapsed.Days
$Hours = $sw.Elapsed.Hours
$Minutes = $sw.Elapsed.Minutes
$Seconds = $sw.Elapsed.Seconds
$Milliseconds = $sw.Elapsed.Milliseconds
$ScriptStopTime = (Get-Date).ToString('dd-MM-yyyy hh:mm:ss')
$Elapsed = "<tr>
               <td>$ScrptStartTime</td>
               <td>$ScriptStopTime</td>
               <td>$Days</td>
               <td>$Hours</td>
               <td>$Minutes</td>
               <td>$Seconds</td>
               <td>$Milliseconds</td>
               <td>$myhost</td>
               
            </tr>
         "
$ScriptExecutionRow += $Elapsed
Add-Content $HealthReport $ScriptExecutionRow
Add-Content $HealthReport "</table></div>"






#---------------------------------------------------------------------------------------------------------------------------------------------
# DC Discovery Details
#---------------------------------------------------------------------------------------------------------------------------------------------



$ExecutionDetailsRow = "<div id=discovercontainer><table>
   <caption><h2><a name='DC Discovery details'>DC Discovery details</h2></caption>
    <thead><th>Total DC Discovered</th>
    <th>DC Successfully Checked</th>
    <th>DC with Scheduled task creation failed</th>
    <th>DC with Folder Creation Failed</th>
    <th>DC with File Not found</th>
    </thead><tbody>
    "

[int]$Successfull_machines_count = $DCList.Count


$DCList_list = foreach ($dc in $DCList) {
                "$dc<br/>"
                }


$ExecutionDetails = 
			"<tr>
               <!--Total Discovered Machines i.e. DCList fetched intially in the begining-->
               <td><a href='javascript:void(0)' target=_blank onclick=""Powershellparamater('"+ $total_machines_discovered_list +"')"">$total_machines_discovered_count</a></td>
               <!--Successfull machines i.e. DCList remaining after the removal of unsuccessful machines-->
               <td><a href='javascript:void(0)' target=_blank onclick=""Powershellparamater('"+ $DCList_list +"')"">$Successfull_machines_count</a></td>"
               
if($machines_with_failed_task_creation_count -ne 0)
{           
    $ExecutionDetails += "<td class=failed><a href='javascript:void(0)' target=_blank onclick=""Powershellparamater('"+ $machines_with_failed_task_creation_list +"')"">$machines_with_failed_task_creation_count</a></td>"
} 
else 
{
    $ExecutionDetails += "<td>$machines_with_failed_task_creation_count</td>"
}


               

if($machines_with_failed_directoy_creation_count  -ne 0)
{
    $ExecutionDetails += "<td class=failed><a href='javascript:void(0)' target=_blank onclick=""Powershellparamater('"+ $machines_with_failed_directoy_creation_list +"')"">$machines_with_failed_directoy_creation_count</a></td>"
} 
else 
{
    $ExecutionDetails += "<td class=passed>$machines_with_failed_directoy_creation_count</td>"
}


if($machines_with_file_not_found_count  -ne 0)
{
    $ExecutionDetails += "<td class=failed><a href='javascript:void(0)' target=_blank onclick=""Powershellparamater('"+ $machines_with_file_not_found_list +"')"">$machines_with_file_not_found_count</a></td>"
} 
else 
{
    $ExecutionDetails += "<td class=passed>$machines_with_file_not_found_count</td>"
}
                                         
            "</tr>"

$ExecutionDetailsRow += $ExecutionDetails
Add-Content $HealthReport $ExecutionDetailsRow
Add-Content $HealthReport "</table></div>"






#---------------------------------------------------------------------------------------------------------------------------------------------
# Cleaning Old Files
#---------------------------------------------------------------------------------------------------------------------------------------------


# Clean Transcripts, Reports and Log Files Older than Retention Days

if($retentionDays)
{
    # Claan HTML Reports
    CleanFiles -p "$Dir" -d $retentionDays -f "^Reports.*"
    
    # Clean Transcripts
    CleanFiles -p "$Dir\Collator\transcripts" -d $retentionDays -f "^transcript_.*"

    # Clean Log Files
    CleanFiles -p "$Dir" -d $retentionDays -f "^Log.*"
}
else
{
	Write-host 'Retention Days value not found in Config File. So not deleting any files'
}





#---------------------------------------------------------------------------------------------------------------------------------------------
# Sending Mail
#---------------------------------------------------------------------------------------------------------------------------------------------

if($SendEmail -eq 'Yes' ) {

    # Send ADHealthCheck Report
    if(Test-Path $HealthReport) 
    {
        try {
            $body = "Please find AD Health Check report attached."
            $port = "25"
            Send-MailMessage -Priority High -Attachments $HealthReport -To $emailTo -From $emailFrom -SmtpServer $smtpServer -Body $Body -Subject $emailSubject -Port $port -ErrorAction Stop
        } catch {       
            Write-host -ForegroundColor Yellow 'Error in sending AD Health Check Report'
            "<p><span class=logerror>Error in sending AD Health Check Report</span></p>" | Out-File -Append $Logfile       
        }
    }

    
    #Send an ERROR mail if Report is not found 
    if(!(Test-Path $HealthReport)) 
    {

        try {
            $body = "ERROR: NO AD Health Check report"
            $port = "25"
            Send-MailMessage -Priority High -To $emailTo -From $emailFrom -SmtpServer $smtpServer -Body $Body -Subject $emailSubject -Port $port -ErrorAction Stop
        } catch {
            Write-host -ForegroundColor Yellow 'Unable to send Error mail.'
            "<p><span class=logerror>Unable to send Error mail.</span></p>" | Out-File -Append $Logfile
        }
    }

}
else
{
    Write-Host "As Send Email is NO so report through mail is not being sent. Please find the report in Script directory."
    "<p><span class=loginfo>As Send Email is NO so report through mail is not being sent. Please find the report in Script directory.</span></p>" | Out-File -Append $Logfile  

}



#--------------------------------------------------------------------------------------------------------------------------------------------------------#
#                                                                                                                                                        #
#                                                                                                                                                        #
#                                                                                                                                                        #
#                                                                                                                                                        #
#                                                                                                                                                        #
#                                                                                                                                                        #
#                                              End of AD Health Check Script                                                                             #
#                                                                                                                                                        #
#                                       For Any debugging Check Script Transcript located in                                                             # 
#                                           folder "Script directory\Collator\transcripts"                                                               #
#                                                                                                                                                        #
#                                              Also Check Log File LogDD_MM_YYYY-hh_mm_ss.html in HTML format in                                         #
#                                                                       Script Direcytory                                                                #
#                                                                                                                                                        #
#                                                                                                                                                        # 
#                                                                                                                                                        #
#                                                                                                                                                        #
#                                                                                                                                                        #
#                                                                                                                                                        #
#--------------------------------------------------------------------------------------------------------------------------------------------------------#