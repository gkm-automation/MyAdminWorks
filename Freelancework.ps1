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
