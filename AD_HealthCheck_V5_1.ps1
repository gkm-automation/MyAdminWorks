$StartTime = $(get-date)
$strComputer=$env:COMPUTERNAME

########## Ensure Script can generate the log ########
$healthcheck = "\\" + "$strComputer"+ "\ADHealthCheck"
if(!(Test-path $healthcheck) -or !(Test-path “$dir\ADHealthCheck\log"))
{
    New-Item “$dir\ADHealthCheck" –type "directory" -ErrorAction SilentlyContinue
    New-SMBShare –Name “ADHealthCheck” –Path “$dir\ADHealthCheck” –FullAccess "CIS\domain admins" -ReadAccess “authenticated users” -ErrorAction SilentlyContinue
    New-Item “$dir\ADHealthCheck\log" –type "directory" -ErrorAction SilentlyContinue
}
gci "$dir\ADHealthCheck\log" | Remove-Item
########## RUNNING DCDIAG ##############
$DCDIAGDATALOG=@()
$dcdiag=@()

$DCDIAGDATALOG += "`nDCDIAG Starts`n"

$dcdiag=(Dcdiag.exe /v /s:$strComputer)
If (($dcdiag -eq $null) -or ($Dcdiag | Select-String -pattern "Ldap search capability attribute search failed on server"))
{
    $DCDIAGDATALOG+="Could not run DC Diagnosis"
}
else
{
    $failedTestList = @( $Dcdiag |select-string -pattern "failed test")
    $allfailedtests=@()
    
    if($failedTestList)
    {
        foreach($test in $failedTestList) 
        {
            $obj=($test -split "test")[1].Trim()
            $allfailedtests+=$obj
        }
        foreach($test in $allfailedtests)
        {
            $test=$test.Trim()
            $FailedRecords = @()
            $datalog =@()
            $from = 0
            $to   = 0
            $FromLine = ""
            $ToLine = ""
		    $TestName1 = ""
            $FromLine = "Starting test: $Test"
            $ToLine = "failed test $Test"

            [int]$from =  (($Dcdiag | Select-String -pattern $FromLine | Select-Object LineNumber).LineNumber)-1
            [int]$to   =  ($Dcdiag  | Select-String -pattern $ToLine | Select-Object LineNumber).LineNumber
                
            $DCDIAGDATALOG += $Dcdiag | Select-Object -Index ("$from".."$to") 
        }
    }
    else
    {
        $DCDIAGDATALOG += "`All tests are passed"
    }
}
$DCDIAGDATALOG += "`nDCDIAG Ends`n"


########## RUNNING DNSDIAG ##############

$DcdiagDNS = @()
$DNSDIAGDATALOG=@()
$DNSDIAGDATALOG += "`nDNSDIAG Starts`n"
$DcdiagDNS = (Dcdiag.exe /test:dns /v /s:$strComputer)
If (($dcdiagdns -eq $Null) -or ($Dcdiagdns | Select-String -pattern "Ldap search capabality attribute search failed on server"))
{
    $DNSDIAGDATALOG+="Could not run DNS Diagnosis`n"
}
else
{
    $TestLine = ( ($DcdiagDNS |select-string -pattern "Summary of DNS test results:" | Select-Object LineNumber).linenumber) + 6
    $DcdiagDNSText=@($DcdiagDNS)
    $TestStatusline=$($DcdiagDNSText[$testline])
    
    $Auth = ($TestStatusline.split()| where {$_})[1]
    $Basc = ($TestStatusline.split()| where {$_})[2]
    $Forw = ($TestStatusline.split()| where {$_})[3]
    $Del = ($TestStatusline.split()| where {$_})[4]
    $Dyn = ($TestStatusline.split()| where {$_})[5]
    $RReg = ($TestStatusline.split()| where {$_})[6]
    $Ext = ($TestStatusline.split()| where {$_})[7]
    
    $FailedDCDNSTest=@()
    
    if($Auth -eq "PASS" -or $Auth -eq "n/a"){}Else{$FailedDCDNSTest += "TEST: Authentication;TEST: Basic;Summary of DNS test results:"; }
    if($Basc -eq "PASS" -or $Basc -eq "n/a"){}Else{$FailedDCDNSTest += "TEST: Basic;TEST: Forwarders/Root hints;Summary of DNS test results:"}
    if($Forw -eq "PASS" -or $Forw -eq "n/a"){}Else{$FailedDCDNSTest += "TEST: Forwarders/Root hints;TEST: Delegations;Summary of DNS test results:"}
    if($Del -eq "PASS" -or $Del -eq "n/a"){}Else{$FailedDCDNSTest += "TEST: Delegations;TEST: Dynamic update;Summary of DNS test results:"}
    if($Dyn -eq "PASS" -or $Dyn -eq "n/a"){}Else{$FailedDCDNSTest += "TEST: Dynamic update;TEST: Records registration;Summary of DNS test results:"}
    if($RReg -eq "PASS" -or $RReg -eq "n/a"){}Else{$FailedDCDNSTest += "TEST: Records registration;Summary of test results;Summary of DNS test results:"}
    #$FailedDCDNSTest
    
    $DcdiagDNSCount = 0
        
    $DcdiagDNSCount = $FailedDCDNSTest.count

    if($DcdiagDNSCount -gt 0)
    {
        foreach($lines in $FailedDCDNSTest)
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
							
            $DNSDIAGDATALOG+=$DcdiagDNS | Select-Object -Index ("$from".."$to") 
        }
    }
    else
    {
        $DNSDIAGDATALOG+="`nAll tests are passed`n"
    }
}

$DNSDIAGDATALOG+="`nDNSDIAG Ends`n"

#>
########## Running Backup Tests ###########

$BKPDATALOG=@()
$BKPDATALOG += "`nBackup Test Starts`n"
$LastBackUpTime = $null
$LastBackUpStatus = $null
$Pattern = "\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}";
$MAtches = $(@(Repadmin.exe /showbackup)| Select-String -Pattern $pattern   -AllMatches| Select-Object -First 1);
if($Matches -match  $pattern){$LastBackUpTime = $matches[0]}
	
	if($LastBackUpTime)
	{
		$Time1 = $LastBackUpTime
		$Time2 = Get-Date 
		$TimeDiff = New-TimeSpan $Time1 $Time2

		if(($TimeDiff.Days -lt 1) -or ($TimeDiff.Days -eq 1 -and $TimeDiff.Hours -eq 0 -and $TimeDiff.Minutes -eq 0 -and $TimeDiff.Seconds -eq 0))
		{
			$LastBackUpStatus += "Test is Success |Last backup on $LastBackUpTime"
		}
		Else
		{
			$LastBackUpStatus += "Test Failed |Last backup on $LastBackUpTime"
		}
		$BKPDATALOG += $LastBackUpStatus
	}
	Else
	{
		$BKPDATALOG += "Failed"
	}
$BKPDATALOG += "`n Backup Test Ends`n"


########## Running Replication Tests ########
$REPLDATALOG = @()
$REPLDATALOG += "`n Replication Test Starts`n"
$myReplInfo = @( repadmin /showrepl * /csv | ConvertFrom-Csv |?{($_.'Number of Failures' -ne 0) -or($_.'Last Failure Time' -ne 0) -or ($_.'Last Failure Status' -ne 0) })
if($myReplInfo.count)
{
    foreach($cmdline in $myReplInfo)
    {
        $CmdSuccess = 1
		$showrepl_COLUMNS = $cmdLine.'showrepl_COLUMNS'
		$DestinationDSASite = $cmdLine.'Destination DSA Site'
		$DestinationDSA = $cmdLine.'Destination DSA'
		$NamingContext = $cmdLine.'Naming Context'
		$SourceDSASite = $cmdLine.'Source DSA Site'
		$SourceDSA = $cmdLine.'Source DSA'
		$TransportType = $cmdLine.'Transport Type'
		$FailureNo = $cmdLine.'Number of Failures'
		$LFailureTime = $cmdLine.'Last Failure Time'
		$LastSuccessTime = $cmdLine.'Last Success Time'
		$FailureStatus = $cmdLine.'Last Failure Status'
        if($FailureNo)
        {
            $REPLDATALOG+=$cmdline
        }
    }
}
else
{
    $REPLDATALOG += "AD Replication is a success."
}

$REPLDATALOG += "`n Replication Test Ends`n"



########## Running Sysvol Checks ###########

$SYSVOLDATALOG =@()
$NETLOGONDATALOG =@()

$SYSVOLDATALOG += "`nSysvol Test Starts`n"
$NETLOGONDATALOG +="`nNetlogon Test Starts`n"
$sysvolstate = Get-WmiObject -Class Win32_Share -Filter "Name='SYSVOL'" -ComputerName "$strComputer"
$Netlogonstate = Get-WmiObject -Class Win32_Share -Filter "Name='Netlogon'" -ComputerName "$strComputer"
$sysvollogicalpath = "\\"+"$strComputer" + "\" + "$($sysvolstate.name)"
$Netlogonlogicalpath = "\\"+"$strComputer" + "\" + "$($netlogonstate.name)"
$sysvolphysicalpath = $($sysvolstate.path)
$Netlogonphysicalpath = $($Netlogonstate.path)

if(test-path $sysvolphysicalpath)
{
    if(Test-Path $sysvollogicalpath){$SYSVOLDATALOG += "Sysvol Exists and is Shared"}else{$SYSVOLDATALOG += "Sysvol Exists but not Shared"}
}
else
{
    $SYSVOLDATALOG += "Sysvol does not Exist"
}

if(test-path $netlogonphysicalpath)
{
    if(Test-Path $netlogonlogicalpath){$NETLOGONDATALOG += "netlogon Exists and is Shared"}else{$NETLOGONDATALOG += "netlogon Exists but not Shared"}
}
else
{
    $NETLOGONDATALOG += "Netlogon does not Exist"
}

Write-Host -ForegroundColor DarkYellow "##########################`n"
Write-Host -ForegroundColor DarkYellow "###### SYSVOL&NETLOGON Starts #####`n"
Write-Host -ForegroundColor DarkYellow "##########################`n"
$SYSVOLDATALOG += "`nSysvol Test Ends`n"
$NETLOGONDATALOG += "`nNetlogon Test Ends`n"
Write-Host -ForegroundColor DarkYellow "`n##########################`n"
Write-Host -ForegroundColor DarkYellow "###### SYSVOL&NETLOGON Ends #######`n"
Write-Host -ForegroundColor DarkYellow "##########################`n`n`n"

######### Disk Space Test ###########

$DISKDATALOG = @()
$DISKDATALOG += "`nDisk Test Starts`n"
$DISKDATALOG += Get-WmiObject Win32_LogicalDisk -Filter "DriveType = '3'" -Computer $strComputer |

    Select-Object DeviceID,
	@{ Name = "Size";       Expression = { "{0:N1}" -f ($_.size / 1GB) } },
	@{ Name = "FreeSpace"; Expression = { "{0:N1}" -f ($_.freespace / 1GB) } },
	@{ Name = "FreeSpaceCent";  Expression = { "{0:P2}" -f (($_.freespace / 1GB) / ($_.size / 1GB)) } }
$DISKDATALOG += "`nDisk Test Ends`n"
########## Services Test ##############

$SVCDATALOG=@()
$SNR=$null
$SVCDATALOG += "`nService Test Starts`n"
$RegKey="SYSTEM\CurrentControlSet\Services\DFSR\Parameters\SysVols\Migrating Sysvols"
$Service_WithoutNTFRS = "NTDS,kdc,w32time,DFSR,RpcSs,DnsCache,IsmServ,SamSs,LanmanServer,LanmanWorkstation,NETLOGON,EVENTSYSTEM,DNS"
$Value=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$strComputer)
$sysvolmigration = $value.OpenSubKey($RegKey)
#-------Include NTFRS?-------------#
if($sysvolmigration)
{
    if($sysvolmigration.GetValue('Local State') -eq "3")
    {
        $checkservices=$Service_WithoutNTFRS
    }
    else
    {
        $checkservices=$Service_WithoutNTFRS + ",ntfrs"
    }
}
else
{
    $checkservices=$Service_WithoutNTFRS + ",ntfrs"
}
#---------Include DHCP?-----------#
$DHCPcheck = Get-Service -Name "DHCPServer"
if($DHCPcheck)
{
    $checkservices = $checkservices + ",DHCPserver"
}
#---------Start the check---------#
$ServiceNames = $checkservices.Split(",")
foreach($service in $ServiceNames)
{
    if($Service.Trim())
    {
        $serviceState = get-service -name $service | Select-Object Name, Status
        $servicestate
        if($serviceState.status -ne "running")
        {
            $serviceState | %{$res="$($_.Status),"+"$($_.name)"}
            $SVCDATALOG +="`n"
            $SVCDATALOG += $res
            $SNR=1
        }
    }
}
if(!($SNR))
{
$SVCDATALOG = "`nAll services are up`n"
}

$SVCDATALOG += "`nService Test Ends`n"


############# Performance Check #########
$PERFDATALOG=@()
$PERFDATALOG += "`nPerformance Test Starts`n"
$Pro_PT=Get-Counter -computername $strComputer -Counter "\processor(_total)\% processor time" | Select-Object -ExpandProperty countersamples | Select-Object -Property cookedvalue 
$Pro_PrT = Get-Counter -computername $strComputer -Counter "\Processor(_total)\% Privileged Time" | Select-Object -ExpandProperty countersamples | Select-Object -Property cookedvalue
$Mem_AM = Get-Counter -computername $strComputer -Counter "\Memory\Available MBytes" | Select-Object -ExpandProperty countersamples | Select-Object -Property cookedvalue
$Mem_CBIU = Get-Counter -computername $strComputer -Counter "\Memory\% Committed Bytes In Use" | Select-Object -ExpandProperty countersamples | Select-Object -Property cookedvalue
$Mem_PS = Get-Counter -computername $strComputer -Counter "\Memory\Pages/sec" | Select-Object -ExpandProperty countersamples | Select-Object -Property cookedvalue
$Mem_PFS = Get-Counter -computername $strComputer -Counter "\Memory\Page Faults/sec" | Select-Object -ExpandProperty countersamples | Select-Object -Property cookedvalue
$NTDS_LBT = Get-Counter -computername $strComputer -Counter "\NTDS\LDAP Bind Time" | Select-Object -ExpandProperty countersamples | Select-Object -Property cookedvalue
$PD_AvgDiskSR = Get-Counter -computername $strComputer -Counter "\PhysicalDisk(_total)\Avg. Disk sec/Read" | Select-Object -ExpandProperty countersamples | Select-Object -Property cookedvalue
$PD_AvgDiskST = Get-Counter -computername $strComputer -Counter "\PhysicalDisk(_total)\Avg. Disk sec/Transfer" | Select-Object -ExpandProperty countersamples | Select-Object -Property cookedvalue
$PD_AvgDiskSW = Get-Counter -computername $strComputer -Counter "\PhysicalDisk(_total)\Avg. Disk sec/Write" | Select-Object -ExpandProperty countersamples | Select-Object -Property cookedvalue
$PD_AvgDiskQL = Get-Counter -computername $strComputer -Counter "\PhysicalDisk(_total)\Avg. Disk Queue Length" | Select-Object -ExpandProperty countersamples | Select-Object -Property cookedvalue
$PD_DT = Get-Counter -computername $strComputer -Counter "\PhysicalDisk(_total)\% Disk Time" | Select-Object -ExpandProperty countersamples | Select-Object -Property cookedvalue
$Sys_PrQL = Get-Counter -computername $strComputer -Counter "\System\Processor Queue Length" | Select-Object -ExpandProperty countersamples | Select-Object -Property cookedvalue

$PERFDATALOG+="CPU | "+"\processor(_total)\% processor time = " + $Pro_PT.CookedValue
$PERFDATALOG+="CPU | "+"\Processor(_total)\% Privileged Time = " + $Pro_PrT.CookedValue
$PERFDATALOG+="Memory | "+"\Memory\Available MBytes = " + $Mem_AM.CookedValue
$PERFDATALOG+="Memory | "+"\Memory\% Committed Bytes In Use = " + $Mem_CBIU.CookedValue
$PERFDATALOG+="Memory | "+"\Memory\Pages/sec = " + $Mem_PS.CookedValue
$PERFDATALOG+="Memory | "+"\Memory\Page Faults/sec = " + $Mem_PFS.CookedValue
$PERFDATALOG+="Directory Services | "+"LDAP Bind Time = " + $NTDS_LBT.CookedValue
$PERFDATALOG+="Disk | "+"\PhysicalDisk(_total)\Avg. Disk sec/Read = " + $PD_AvgDiskSR.CookedValue
$PERFDATALOG+="Disk | "+"\PhysicalDisk(_total)\Avg. Disk sec/Transfer = " + $PD_AvgDiskST.CookedValue
$PERFDATALOG+="Disk | "+"\PhysicalDisk(_total)\Avg. Disk sec/Write = " + $PD_AvgDiskSW.CookedValue
$PERFDATALOG+="Disk | "+"\PhysicalDisk(_total)\Avg. Disk Queue Length = " + $PD_AvgDiskQL.CookedValue
$PERFDATALOG+="Disk | "+"\PhysicalDisk(_total)\% Disk Time = " + $PD_DT.CookedValue
$PERFDATALOG += "`nPerformance Test Ends`n"

$outputfile = "$dir\ADHealthCheck\log\" + $strComputer + "_adhealthcheck.txt"

$DCDIAGDATALOG | Out-File -Append $outputfile
$DNSDIAGDATALOG | Out-File -Append $outputfile
$SYSVOLDATALOG | Out-File -Append $outputfile
$NETLOGONDATALOG | Out-File -Append $outputfile
$REPLDATALOG | Out-File -Append $outputfile
$BKPDATALOG | Out-File -Append $outputfile
$DISKDATALOG | Out-File -Append $outputfile
$SVCDATALOG | Out-File -Append $outputfile
$PERFDATALOG | Out-File -Append $outputfile
$ElapsedTime = $(get-date) - $starttime
$totalTime = "{0:HH:mm:ss}" -f ([datetime]$elapsedTime.Ticks)
$totalTime | Out-File -Append $outputfile