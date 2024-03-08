#Change Baseline CSV path
$base = Import-Csv "C:\Users\\Downloads\results\base.csv" | Where-Object { $_.label -Like "ClaimsFNOL*" } | Select-Object label,@{n='elapsed';e={$_.elapsed/1000}}
#Change Current CSV Path
$current = Import-Csv "C:\Users\\Downloads\results\current.csv" | Where-Object { $_.label -Like "ClaimsFNOL*" } | Select-Object label,@{n='elapsed';e={$_.elapsed/1000}}
$results = @()
foreach($record in $current){
    $basevalue = ($base | Where-Object { $_.label -eq $record.label }).'elapsed'[0]
    $diffValue = $basevalue - $($record.'elapsed')
    $divMark = (($basevalue/100)*5)+ $basevalue
    If($($record.'elapsed') -ge $divMark){
        $DeviationThreshold = "Reached"
    }
    Else{
        $DeviationThreshold = "Below Limit"
    }
    $obj = [PSCustomObject]@{
        Label = $record.label
        Base_Lapsed = $basevalue
        'current_Lapsed' = $($record.'elapsed')
       'Difference(Secs)' = $diffValue

    }
$results += $obj
}

Write-Output ($results | ft)

## Only following 5 variables are required to send mail
$myorg = “my-ado-org”
$myproj = “my-ado-project”
$sendmailto = “devops.user1@xyz.com,devops.user2@xyz.com” ## comma separated email ids of receivers
$mysubject = “my custom subject of the mail” ## Subject of the email
$mailbody = “my custom mail body details” ## mail body
#########################
## Get tfsids of users whom to send mail
$mailusers = “$sendmailto”
$mymailusers = $mailusers -split “,”
$pat = “Bearer $env:System_AccessToken”
$myurl =”https://dev.azure.com/${myorg}/_apis/projects/${myproj}/teams?api-version=5.1"
$data = Invoke-RestMethod -Uri “$myurl” -Headers @{Authorization = $pat}
$myteams = $data.value.id
##Get list of members in all teams
$myusersarray = @()
foreach($myteam in $myteams) {
$usrurl = “https://dev.azure.com/${myorg}/_apis/projects/${myproj}/teams/"+$myteam+"/members?api-version=5.1"
$userdata = Invoke-RestMethod -Uri “$usrurl” -Headers @{Authorization = $pat}
$myusers = $userdata.value
foreach($myuser in $myusers) {
$myuserid = $myuser.identity.id
$myusermail = $myuser.identity.uniqueName
$myuserrecord = “$myuserid”+”:”+”$myusermail”
$myusersarray += $myuserrecord
}
}
## filter unique users
$myfinalusersaray = $myusersarray | sort -Unique
## create final hash of emails and tfsids
$myusershash = @{}
for ($i = 0; $i -lt $myfinalusersaray.count; $i++)
{
$myusershash[$myfinalusersaray[$i].split(“:”)[1]] = $myfinalusersaray[$i].split(“:”)[0]
}
##
## create list of tfsid of mailers
foreach($mymail in $mymailusers) {
$myto = $myto +’”’+$myusershash[$mymail]+’”,’
}
##send mail
$uri = “https://${myorg}.vsrm.visualstudio.com/${myproj}/_apis/Release/sendmail/$(RELEASE.RELEASEID)?api-version=3.2-preview.1"
$requestBody =
@”
{
“senderType”:1,
“to”:{“tfsIds”:[$myto]},
“body”:”${mailbody}”,
“subject”:”${mysubject}”
}
“@
Try {
Invoke-RestMethod -Uri $uri -Body $requestBody -Method POST -Headers @{Authorization = $pat} -ContentType “application/json”
}
Catch {
$_.Exception
}

#Convert Table data To Html Format 
          #$HTMLTable=ConvertTo-Html -InputObject ($result) -Fragment
          $Body = "<h1>Test Data</h1><ul>"  # Start building the HTML body
          $Body += "<table border='1' cellspacing='0' cellpadding='4'><tr><th>TransactionName</th><th>BaseLine AVG PageTime</th><th>Current AVG PageTime</th><th>Difference (Secs)</th></tr>"
          # Loop through the array and add each element to the HTML body
          foreach ($item in $results) {
              $Body += "<tr>"
              $Body += "<td>$($item.Label)</td>"
              $Body += "<td>$($item.Base_Lapsed)</td>"
              $Body += "<td>$($item.current_Lapsed)</td>"
              $Body += "<td>$($item.Difference_Secs)</td>"
              $Body += "</tr>"
          }

          $Body += "</table>"  # Close the table tag
         # $ToList="sreddy@hagerty.com","mlal@Hagerty.com","mande@hagerty.com","achaudhary@hagerty.com"
          [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
          # Define email parameters
          $From = "ODCH_SMTP_svc@hagerty.com"
          #$To = $ToList
          $To = "sreddy@hagerty.com"
          $Subject = "Pipeline Run Notification"
          #$Body = "<h2> Table Data</h2>$HTMLTable"

          # Specify SMTP server settings
          $SMTPServer = "smtp.office365.com"
          $SMTPPort = 587 # or the appropriate port for your SMTP server
          $Username = "ODCH_SMTP_svc@hagerty.com"
          $Password = "$(smtppwd)"

          # Create an SMTP client object
          $SMTPClient = New-Object System.Net.Mail.SmtpClient($SMTPServer, $SMTPPort)
          $SMTPClient.EnableSsl = $true  # Enable SSL/TLS encryption

          # Set SMTP credentials
          $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($Username, $Password)

          # Create Mail body object
          $Message = New-Object System.Net.Mail.MailMessage($From, $To, $Subject, $Body)
          $Message.IsBodyHtml = $true # setting boday content as a html

          # Send the email
          $SMTPClient.Send($Message)



################UPDATED############
#Change Baseline CSV path
$base = Import-Csv "C:\Users\\Downloads\results\base.csv" | Where-Object { $_.label -Like "ClaimsFNOL*" } | Select-Object label,@{n='elapsed';e={$_.elapsed/1000}}
#Change Current CSV Path
$current = Import-Csv "C:\Users\\Downloads\results\current.csv" | Where-Object { $_.label -Like "ClaimsFNOL*" } | Select-Object label,@{n='elapsed';e={$_.elapsed/1000}}
$results = @()
[Bool]$flag = $false
foreach($record in $current){
    $basevalue = ($base | Where-Object { $_.label -eq $record.label }).'elapsed'[0]
    $diffValue = $basevalue - $($record.'elapsed')
    $divMark = (($basevalue/100)*5)+ $basevalue
    If($($record.'elapsed') -ge $divMark){
        $flag = $true
    }
    $obj = [PSCustomObject]@{
        Label = $record.label
        Base_Lapsed = $basevalue
        'current_Lapsed' = $($record.'elapsed')
       'Difference(Secs)' = $diffValue

    }
$results += $obj
}

#Convert Table data To Html Format 
          #$HTMLTable=ConvertTo-Html -InputObject ($result) -Fragment
          $Body = "<h1>Test Data</h1><ul>"  # Start building the HTML body
          $Body += "<table border='1' cellspacing='0' cellpadding='4'><tr><th>TransactionName</th><th>BaseLine AVG PageTime</th><th>Current AVG PageTime</th><th>Difference (Secs)</th></tr>"
          # Loop through the array and add each element to the HTML body
          foreach ($item in $results) {
              $Body += "<tr>"
              $Body += "<td>$($item.Label)</td>"
              $Body += "<td>$($item.Base_Lapsed)</td>"
              $Body += "<td>$($item.current_Lapsed)</td>"
              $Body += "<td>$($item.Difference_Secs)</td>"
              $Body += "</tr>"
          }

          $Body += "</table>"  # Close the table tag

         $ToList=@("sreddy@hagerty.com","mlal@Hagerty.com","mande@hagerty.com","achaudhary@hagerty.com")
          [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
          # Define email parameters
          $From = "ODCH_SMTP_svc@hagerty.com"
          $To = $ToList
          #$To = "sreddy@hagerty.com"
          $Subject = "Pipeline Run Notification"
          #$Body = "<h2> Table Data</h2>$HTMLTable"

          # Specify SMTP server settings
          $SMTPServer = "smtp.office365.com"
          $SMTPPort = 587 # or the appropriate port for your SMTP server
          $Username = "ODCH_SMTP_svc@hagerty.com"
          $Password = "$(smtppwd)"

          # Create an SMTP client object
          $SMTPClient = New-Object System.Net.Mail.SmtpClient($SMTPServer, $SMTPPort)
          $SMTPClient.EnableSsl = $true  # Enable SSL/TLS encryption

          # Set SMTP credentials
          $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($Username, $Password)

          # Create Mail body object
          $Message = New-Object System.Net.Mail.MailMessage($From, $To, $Subject, $Body)
          $Message.IsBodyHtml = $true # setting boday content as a html

          if($flag){
            # Send the email
            $SMTPClient.Send($Message)
          }
          
