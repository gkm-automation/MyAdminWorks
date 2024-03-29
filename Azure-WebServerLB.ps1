﻿<#
Create an Azure load balancer
Create a load balancer health probe
Create load balancer traffic rules
Use the Custom Script Extension to create a basic IIS site
Create virtual machines and attach to a load balancer
View a load balancer in action
Add and remove VMs from a load balancer
#>



$rg = 'myrg'
$location ="East US"
#credentials for VM Creation
$crd = $(Get-Credential)

#create RG
New-AzureRmResourceGroup -Name $rg -Location $location

#create Public IP
$publicip = New-AzureRmPublicIpAddress -Name "mypublicip" -ResourceGroupName $rg -Location $location -AllocationMethod Static

#Create LB frontpool

$frontpool = New-AzureRmLoadBalancerFrontendIpConfig -Name "lbfront" -PublicIpAddress $publicip


#Create LBback end
$backpool = New-AzureRmLoadBalancerBackendAddressPoolConfig -Name "lbbackend" 

#create LB
New-AzureRmLoadBalancer -Name "mylb" -ResourceGroupName $rg -Location $location -FrontendIpConfiguration $frontpool -BackendAddressPool $backpool 

#get LB details
$lb = Get-AzureRmLoadBalancer -Name "mylb" -ResourceGroupName $rg

#create health probe
Add-AzureRmLoadBalancerProbeConfig -Name "lbhealthprobe" -LoadBalancer $lb -Protocol Tcp -Port 80 -ProbeCount 2 -IntervalInSeconds 15 

#Associate health probe with LB
Set-AzLoadBalancer -LoadBalancer $lb

#Get Probe
$probe = Get-AzureRmLoadBalancerProbeConfig -LoadBalancer $lb -Name "lbhealthprobe"

#Create LB Rule
Add-AzureRmLoadBalancerRuleConfig -Name "LBrule" -LoadBalancer $lb -Protocol Tcp `
                          -FrontendIpConfiguration $lb.FrontendIpConfigurations[0] `
                          -BackendAddressPool $lb.BackendAddressPools[0] `
                          -FrontendPort 80 `
                          -BackendPort 80 `
                          -Probe $probe

Set-AzureRmLoadBalancer -LoadBalancer $lb

# Create a virtual network with a front-end subnet
$subnetconfig = New-AzureRmVirtualNetworkSubnetConfig -Name "VMSub" -AddressPrefix "10.0.1.0/24"
$vnet = New-AzureRmVirtualNetwork -Name "vnet" -ResourceGroupName $rg -Location $location -AddressPrefix 10.0.0.0/16 -Subnet $subnetconfig

 # Create an NSG rule to allow HTTP traffic in from the Internet to the front-end subnet.
 
 $nsgrule1 = New-AzureRmNetworkSecurityRuleConfig -Name "AllowHTTP" -Description "AllowHTTP" -Protocol Tcp -SourcePortRange * -DestinationPortRange 80 -SourceAddressPrefix * -DestinationAddressPrefix * -Priority 101 -Direction Inbound -Access Allow
 $nsgrule2 = New-AzureRmNetworkSecurityRuleConfig -Name "AllowRDP" -Description "AllowRDP" -Protocol Tcp -SourcePortRange * -DestinationPortRange 3389 -SourceAddressPrefix * -DestinationAddressPrefix * -Priority 101 -Direction Inbound -Access Allow

$nsg = New-AzureRmNetworkSecurityGroup -Name "nsg1" -ResourceGroupName $rg -Location $location -SecurityRules $nsgrule1,$nsgrule2 -Force

Set-AzureRmVirtualNetworkSubnetConfig -Name "VMSub" -VirtualNetwork $vnet -NetworkSecurityGroup $nsg -AddressPrefix 10.0.1.0/24   
                                    
#create Availability Set
$as = New-AzureRmAvailabilitySet -Name "AS1" -ResourceGroupName $rg -Location $location -Sku aligned -PlatformUpdateDomainCount 5 -PlatformFaultDomainCount 2


for ($i=1; $i -le 3; $i++)
{
   New-AzNetworkInterface `
     -ResourceGroupName $rg `
     -Name myVM$i `
     -Location $location `
     -Subnet $vnet.Subnets[0] `
     -LoadBalancerBackendAddressPool $lb.BackendAddressPools[0]
}

#Create Web Server VM
for ($i=1; $i -le 3; $i++)
{
    New-AzVm `
        -ResourceGroupName $rg `
        -Name "myVM$i" `
        -Location $location `
        -VirtualNetworkName "vnet" `
        -SubnetName "VMSub" `
        -SecurityGroupName "nsg1" `
        -OpenPorts 80 `
        -AvailabilitySetName "AS1" `
        -Credential $crd `
        -AsJob
}

#Install IIS with Custom Script Extension

for ($i=1; $i -le 3; $i++)
{
   Set-AzVMExtension `
     -ResourceGroupName $rg `
     -ExtensionName "IIS" `
     -VMName myVM$i `
     -Publisher Microsoft.Compute `
     -ExtensionType CustomScriptExtension `
     -TypeHandlerVersion 1.8 `
     -SettingString '{"commandToExecute":"powershell Add-WindowsFeature Web-Server; powershell Add-Content -Path \"C:\\inetpub\\wwwroot\\Default.htm\" -Value $($env:computername)"}' `
     -Location $location
}