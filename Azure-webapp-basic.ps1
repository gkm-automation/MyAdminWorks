#Sample Azure Web App Creation

$gitrepo = "https://github.com/gkm-automation/app-service-web-dotnet-get-started.git"
$webappname = "$(get-random)"
$location = "South India"
$rg = 'myresourcegroup'

#Create Resource Group

New-AzureRmResourceGroup -Name $rg  -Location $location

#Create App Service Plan
New-AzureRmAppServicePlan -Name $webappname -Location $location -ResourceGroupName $rg -Tier Free

#create web app
New-AzureRmWebApp -Name $webappname -ResourceGroupName $rg -Location $location -AppServicePlan $webappname

#Upgrade plan
Set-AzureRmAppServicePlan -Name $webappname  -ResourceGroupName $rg -Tier Standard

#create deployment slot
New-AzureRmWebAppSlot -Name $webappname -ResourceGroupName $rg -Slot staging

# Configure GitHub deployment to the staging slot from your GitHub repo and deploy once.
$PropertiesObject = @{
    repoUrl = "$gitrepo";
    branch = "master";
}

Set-AzureRmResource -PropertyObject $PropertiesObject -ResourceGroupName $rg `
-ResourceType Microsoft.Web/sites/slots/sourcecontrols `
-ResourceName $webappname/staging/web -ApiVersion 2015-08-01 -Force

#get the web app
Get-AzureRmWebApp -Name $webappname

#register Github toekn into powersh
$PropertiesObject = @{
    token = "4d3d319987fca5d632d7a39839673ea05831d773";
}
Set-AzureRmResource -PropertyObject $PropertiesObject -ResourceId /providers/Microsoft.Web/sourcecontrols/GitHub -ApiVersion 2015-08-01 -Force

#switch slot
Switch-AzureRmWebAppSlot -Name $webappname -SourceSlotName staging  -DestinationSlotName production -ResourceGroupName $rg