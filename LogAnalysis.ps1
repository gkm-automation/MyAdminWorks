
#Global Variable
$global:javafiledir = "D:\script"
$global:inputdir = "D:\Healthcheck"
$global:outputdir = "D:\output"
$global:resultflag = "true"
$global:results = @()
$global:csvpath = "D:\script\Info_sheet.csv"


#SMTP Inputs
$from = "admin@company.com"
$to = "user1@yourdomain","user2@yourdomain"
$smtpserver = "smtp.gmail.com"



<# 
.Synopsis 
   Write-Log writes a message to a specified log file with the current time stamp. 
.DESCRIPTION 
   The Write-Log function is designed to add logging capability to other scripts. 
   In addition to writing output and/or verbose you can write to a log file for 
   later debugging. 
#> 
function Write-Log 
{ 
    [CmdletBinding()] 
    Param 
    ( 
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)] 
        [ValidateNotNullOrEmpty()] 
        [Alias("LogContent")] 
        [string]$Message, 
 
        [Parameter(Mandatory=$false)] 
        [Alias('LogPath')] 
        [string]$Path="$javafiledir\PowerShellLog.log", 
         
        [Parameter(Mandatory=$false)] 
        [ValidateSet("Error","Warn","Info")] 
        [string]$Level="Info", 
         
        [Parameter(Mandatory=$false)] 
        [switch]$NoClobber 
    ) 
 
    Begin 
    { 
        # Set VerbosePreference to Continue so that verbose messages are displayed. 
        $VerbosePreference = 'Continue' 
    } 
    Process 
    { 
         
        # If the file already exists and NoClobber was specified, do not write to the log. 
        if ((Test-Path $Path) -AND $NoClobber) { 
            Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name." 
            Return 
            } 
 
        # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path. 
        elseif (!(Test-Path $Path)) { 
            Write-Verbose "Creating $Path." 
            $NewLogFile = New-Item $Path -Force -ItemType File 
            } 
 
        else { 
            # Nothing to see here yet. 
            } 
 
        # Format Date for our Log File 
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss" 
 
        # Write message to error, warning, or verbose pipeline and specify $LevelText 
        switch ($Level) { 
            'Error' { 
                Write-Error $Message 
                $LevelText = 'ERROR:' 
                } 
            'Warn' { 
                Write-Warning $Message 
                $LevelText = 'WARNING:' 
                } 
            'Info' { 
                Write-Verbose $Message 
                $LevelText = 'INFO:' 
                } 
            } 
         
        # Write log entry to $Path 
        "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append 
    } 
    End 
    { 
    } 
}

<# 
.Synopsis 
   Get-TextParser Executing java program by taking TXT input and parse the output file based on rule. 
.DESCRIPTION 
   The Write-Log function helps to run java application based on given input files and generate the output file.Then it will
   parse the output file and send HTML output report in Email.
#> 


function Get-TextParser {
$ltime = (Get-Date -Format "yy/mm/dd:hh:mm:ss")   
       
 try {        
            
	    ##Verifying Required Files Exists For Execution
	    if(Test-Path -Path "$javafiledir\LogAnalysis.java"){
            Write-Log -Level Info -Message "[$ltime]:File LogAnalysis.java Was Found"
            Get-ChildItem -Path $javafiledir -Filter *.txt | Remove-Item
            #Get-ChildItem -Path $javafiledir -Filter *.class | Remove-Item
            if(!(Test-Path "$javafiledir\Archive")){ New-Item -ItemType Directory -Path "$javafiledir\Archive" | Out-Null }
            $temphtm = "$javafiledir\healthreport.html"
            if(Test-Path $temphtm) {
            $tempvar = Get-ChildItem $temphtm 
            $modified = (Get-Date $tempvar.LastAccessTime -Format MMddyyyy_hhmmss).ToString()
            Move-Item -Path $temphtm -Destination "$javafiledir\Archive\healthreport_$($modified).html"
            Write-Log "[$ltime]:Cleaned Logs & Class files in $javafiledir"
            Write-Log "[$ltime]:Archived Reports in $javafiledir"
            }
            }
        else{ 
            Write-Log -Level Warn -Message "[$ltime]:File LogAnalysis.java Was Not Found..Exit from Script"  -ErrorAction Stop
            $global:resultflag = $false
            return;
            }
        if(Test-Path -Path "$inputdir\*.log"){
            Write-Log -Level Info -Message "[$ltime]:Input Directory $inputdir Was Found"
            }
        else{ 
            Write-Log -Level Warn -Message "[$ltime]:Input Directory $inputdir Was Not Found..Exit from Script" 
            $global:resultflag = "false"
            return;
            }
         if(Test-Path -Path $outputdir){
            Write-Log -Level Info -Message "[$ltime]:Output Directory $outputdir Was Found"
            if(!(Test-Path "$outputdir\Archive")){ New-Item -ItemType Directory -Path "$outputdir\Archive" | Out-Null }
            if(Test-Path "$outputdir\Health_CheckResults.txt") {
            $tempvar = Get-ChildItem "$outputdir\Health_CheckResults.txt" 
            $modified = (Get-Date $tempvar.LastAccessTime -Format MMddyyyy_hhmmss).ToString()
            Move-Item -Path "$outputdir\Health_CheckResults.txt" -Destination "$outputdir\Archive\Health_CheckResults_$($modified).txt"
            Write-Log -Message "[$ltime]:Archived existing output file from $outputdir"
            }
            }
            else{ 
            Write-Log -Level Warn -Message "[$ltime]:Output Directory $outputdir Was Not Found..Exit from Script" 
            $global:resultflag = "false"
            return;
            }
                       
            Write-Log -Message "[$ltime]:Running Java Program"
            $javarun = Start-Process -WorkingDirectory $javafiledir -FilePath java.exe -ArgumentList .\LogAnalysis.java -Wait -NoNewWindow -PassThru -RedirectStandardError "$javafiledir\runerr.log" -RedirectStandardOutput "$javafiledir\runoutput.log" -ErrorAction Stop
            if($javarun.ExitCode -ne 0 -or (Get-Content "$javafiledir\runerr.log")) { 
            Write-Log -Level Warn "[$ltime]:Error in Java Execuation ..Quit Script"
            $global:resultflag = "false"
            return; 
            }
            else { 
            Write-Log -Message "[$ltime]:Java Execution Completed Successfully.."
            Get-Content "$javafiledir\runoutput.log" | foreach { Write-Log "[JAVA_INPUT_FILES]:$_" }
            }
  
            Write-Log -Level Info -Message "[$ltime]:Parsing Java Output file"
            
            $csv = Import-Csv -Path $global:csvpath -ErrorAction Stop
            $errorcode = @()
            $errorcode = $csv.ErrorCode
            Write-Log -Level Info -Message "*********ERROR CODES***************"
            $errorcode | foreach {  Write-Log -Level Info -Message "[ERROR CODES]: $_ " }               

#Check Output files exists
        $outputfile = "$outputdir\Health_CheckResults.txt"
        if(!(Test-Path $outputfile)) { 
         Write-Log -Message "[$ltime]:Output File not generated by Java Program.. Pls check"
         $global:resultflag = "false"
         return;
         }
                 

for($i=0;$i -lt $errorcode.Count; $i++)
{

            $pattern = $([regex]::escape($errorcode[$i]))+'\b'
            #foreach { $out += $_ | Select-String -Pattern $pattern }
            #Local variables declaration
            $Filteredline = Get-Content $outputfile -ErrorAction SilentlyContinue |  Select-String -Pattern $pattern 
            $temp = (($i/$errorcode.Count)*100)
            Write-Progress -Activity "Word extraction Status" -Status "$temp% Complete:" -PercentComplete $temp
            $wonum = @()
            $assign = @()
            $lcode =@()
            $tstmp = ""
            [int]$rcount = 0
    foreach($in in $Filteredline)
    {

            #Write-Host "%%%%%%%%%%"
            $in |  % { if($_ -match "wonum =\s\d{1,8}" -or $_ -match "WO #\d{1,8}") { $wonum += $matches[0] } }
            $in | % { if($_ -match "assignmentid =\s\d{1,9}") { $assign+= $matches[0]} }
            $in | % { if($_ -match "laborcode\s\s=\s\w+\d{1,9}" -or $_ -match "\b[A-Z]{2}\d{1,8}\b") { $lcode += $matches[0] } }
            $in | % { if($_ -match "\d{4}\s\d{2}\s\d{2}\s\d{2}:\d{2}:\d{2}\b") { if ([string]::IsNullOrEmpty($tstmp)) {$tstmp = $matches[0];} $rcount += 1 } }
           #Write-Host "%%%%%%%%%%"

    }
    #Remove words associated in variable
    $errorname = $errorcode[$i]
    $wonum = $wonum -replace "wonum = ",""
    $wonum = $wonum -replace "WO #",""
    $assign = $assign -replace "assignmentid = ",""
    $lcode = $lcode -replace "laborcode  = ",""
    $lcode = $lcode | sort -Unique
    
    #separated records by space
    $wo = $wonum -join ","
    $ac = $assign -join ","
    $lc = $lcode -join ","
        
    #Finding NULL Records
    if ([string]::IsNullOrEmpty($wo)) {$wo = "No"}
    if ([string]::IsNullOrEmpty($ac)) {$ac = "No"}
    if ([string]::IsNullOrEmpty($lc)) {$lc = "No"}

    #Heading for records
    $wo = $wo.Insert(0,"WO:")
    $ac = $ac.Insert(0,"AID:")
    $lc = $lc.Insert(0,"UID:")
    
    $errexp = $csv.Where({$PSItem.Errorcode -eq $errorname }).explanation
    #combine records WO ACODE LCODE
    $ErrorDetails = ($errorname +"=> "+$errexp + $wo +" "+ $ac +" "+ $lc)

    #Finding CSV file to update Records
    $errorimpact = $csv.Where({$PSItem.Errorcode -eq $errorname }).impact
    $errorreq = $csv.Where({$PSItem.Errorcode -eq $errorname }).Requestnumber
   

     
    $ourObject = New-Object -TypeName psobject
    $ourObject | Add-Member -MemberType NoteProperty -Name "Errorcode" -Value $ErrorDetails
    $ourObject | Add-Member -MemberType NoteProperty -Name "Impact" -Value $errorimpact
    $ourObject | Add-Member -MemberType NoteProperty -Name "Request Number" -Value $errorreq
    $ourObject | Add-Member -MemberType NoteProperty -Name "No. of Occurrences" -Value $rcount
    $ourObject | Add-Member -MemberType NoteProperty -Name "Time Stamp" -Value $tstmp

    $global:results += $ourObject
        
 }#for

 Write-Log -Message "[$ltime]:Parsing the file completed..."
 }#try

   catch{ 
           Write-Log -Message "$_.Exception.Message" -Level Error
           $global:resultflag = "false"
           return;
      }

}#function

#Calling function Log Parser

get-textparser

$a = "<style>"
$a = $a + "BODY{background-color:peachpuff;}"
$a = $a + "TABLE{table-layout: Fixed;width: 100%;border-width: 1px;border-style: solid;border-color:  black;border-collapse: collapse;}"
$a = $a + "TH{column-width: 20%;border-width: 1px;padding: 8px;border-style: solid;border-color: black;background-color:thistle}"
$a = $a + "TD{WORD-WRAP:  break-word;column-width: 20%;border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color:PaleGoldenrod;text-align: center}"
$a = $a + "</style>"

Write-Log "[TRIGGER EMAIL]"
if($global:resultflag -eq 'false')
{
Write-Log -Level Warn -Message "Script failed to execute. Pls check Log file.."
Send-MailMessage -To $to -From $from -Subject "[ERROR]Daily Health Check Report - Work Order App $(get-date)" -BodyAsHtml "Script failed to execute. Pls check Log file.." -Priority High -SmtpServer $smtpserver
}
else
{
$outreport = "$javafiledir\healthreport.html"
$results | ConvertTo-Html -Head $a -Body "<H2> SMP server log observations Analysis: $(get-date)</H2>" | Out-File $outreport
$body = Get-Content $outreport -Raw
Send-MailMessage -To $to -From $from -Subject "Daily Health Check Report - Work Order App $(get-date)" -BodyAsHtml $body  -SmtpServer $smtpserver -Attachments $outreport
}



#######################################