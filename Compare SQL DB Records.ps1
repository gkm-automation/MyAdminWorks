$Script:dbInstance = "NAKORE04"
$script:dbService = "MSSQLSERVER"
$Script:dbName = "crs5_oltp_replicated"
#$Script:tableName = 'Antivirus'
#Import SQL Server PS Module
if(!(Get-Module -Name SqlServer)){
Write-Host "SQL PS Module not found.Installing.."
Install-Module -Name "SqlServer" -Force -Confirm:$false
}
#Import-Module -Name "SqlServer"
Try{
#Test Database Connection
Write-Host -Message "Test DB Connection || Database Instance Name:$dbInstance ||Database Name: $databasename"
$GWMIParams = @{
Class = "Win32_Service"
ComputerName = $dbInstance
Filter = "name= '$dbService' and state = 'Running'"
}
if((Test-Connection -ComputerName $dbInstance -Count 1 -Quiet) -and (Get-WmiObject @GWMIParams).State -eq 'Running'){
Write-Host "Database Connectivity Successful...!"
}
else{
Write-Host "Database Connectivity Unsuccessful...!"
#Exit
}
#Execute Stored Procedure in SQL
$sp1Records = Invoke-Sqlcmd -ServerInstance $dbInstance -Database "crs5_oltp_replicated" -Query "exec sp_Rep_comp_wth_prod"
$sp2Records = Invoke-Sqlcmd -ServerInstance $dbInstance -Database "Crs5_oltp_Qlik" -Query "exec sp_Rep_comp_wth_prod_Qlik"
#Write-Output $sq2Records | Out-File  -FilePath D:\script_output\script_output.txt
 
$sp1Records | Where-Object {$PSItem.Count_Diffrence -gt 0} | Select-Object Replication_Table,Production_Table_Count,Prod_Time,Replicated_Table_Count,Replicated_Time,Count_Diffrence,'Time_Diffrence(Milliseconds)',@{N="Database";e={"crs5_oltp_replicated"}}
$sp2Records | Where-Object {$PSItem.Count_Diffrence -gt 0} | Select-Object Replication_Table,Production_Table_Count,Prod_Time,Replicated_Table_Count,Replicated_Time,Count_Diffrence,'Time_Diffrence(Milliseconds)',@{N="Database";e={"Crs5_oltp_Qlik"}}
 
}
Catch{
Write-Host $($_.Exception.Message)
}
if($isCountGreater){
Write-Host "n_______COUNT DIFFERENCE IS GREATER THAN 100_______" Write-Output $sp2Records >> D:\script_output\script_output.txt } else{ Write-Host "n_______COUNT DIFFERENCE IS LESSTER THAN 100_______"
Write-Host ($sqlRecords | FT | Out-String)
#Write-Host $sq2Records | Out-File -FilePath D:\script_output\ Get-Content -Path D:\script_output\script_output.txt
Write-Output $sp2Records >> D:\script_output\script_output.txt
}
 
$sp3Records =
Invoke-Sqlcmd -ServerInstance $dbInstance -Database "msdb" -Query "
 
DECLARE @MAILBODY NVARCHAR(MAX);
DECLARE @MAILSUBJECT NVARCHAR(255);
DECLARE @MAILPROFILE NVARCHAR(128);
DECLARE @RECIPIENTS NVARCHAR(1000);
DECLARE @IMPORTANCE NVARCHAR(6);
DECLARE @LINEBREAK NVARCHAR(2);
DECLARE @FILECONTENT NVARCHAR(MAX);
 
 
select @FILECONTENT = bulkcolumn from openrowset(bulk 'D:\script_output\script_output.txt', SINGLE_BLOB) as content;
--SET EMAIL PARAMETERS
SET @LINEBREAK = NCHAR(13) + NCHAR(10)
--SET @DESCRIPTION = 'Team, Please check the replication status';
SET @MAILBODY = 'Team, Please check the replication status,' + @LINEBREAK + @FILECONTENT;
SET @MAILSUBJECT = 'Replication Alert';
SET @MAILPROFILE = N'NAKORE04_DB_Mail';
SET @RECIPIENTS = 'karthickvelappan@discover.com';
SET @IMPORTANCE = 'High';
 
--SEND EMAIL
EXEC msdb.dbo.sp_send_dbmail
    @profile_name = @MAILPROFILE,
    @recipients = @RECIPIENTS,
    @subject = @MAILSUBJECT,
    @body = @MAILBODY,
    @body_format = 'TEXT',
    @importance = @IMPORTANCE
GO"
#TRUNCATE THE TEXT FILE
$FILEPATH = "D:\script_output\script_output.txt"
set-Content -Path $FILEPATH -Value ""
 
