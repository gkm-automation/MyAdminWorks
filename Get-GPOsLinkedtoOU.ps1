<#
.Synopsis
   The purpose of this script is to get Specific GPO status from Linked OU's from entire forest.
.DESCRIPTION
   This script will get status of specific GPO agaist Requested OU within Domain.Ouput will be save as CSV under working directory.
   You should have proper domain Privilege to run this script.
.EXAMPLE
   ./Get-GPOsLinkedtoOU.ps1 -OUNAME Servers -GPONAME 'Test_GPO'
        

.OUTPUTS
    Name    DistinguishedName                           DisplayName Enforced LinkEnabled BlockInheritance WMIFilter
    ----    -----------------                           ----------- -------- ----------- ---------------- ---------
    Servers OU=Servers,OU=OU1,DC=child,DC=test,DC=local                                             False
    Servers OU=Servers,OU=OU2,DC=child,DC=test,DC=local                                             False
    Servers OU=Servers,OU=OU1,DC=test,DC=local                                                      False
    Servers OU=Servers,OU=OU2,DC=test,DC=local                                                       True
    Servers OU=Servers,OU=OU3,DC=test,DC=local          Test_GPO    False    True                   False WMI_Filter
#>


[CmdletBinding()]
param(
[Parameter(Mandatory=$true)]
[String]$OUNAME,
[Parameter(Mandatory=$true)]
[String]$GPONAME
)


Import-Module GroupPolicy
Import-Module ActiveDirectory

# Get all the domains within forest
$domains = (Get-ADForest).domains

#Report arrary declaration
$report = @()


Foreach($domain in $domains){

        # Grab a list of all GPOs
        $GPOs = Get-GPO -All -Domain $domain | Select-Object ID, Path, DisplayName, GPOStatus, WMIFilter, CreationTime, ModificationTime, User, Computer

        $findgpo = $GPOs | ? {$_.DisplayName -eq $GPONAME}
        if(!$findgpo) { Write-Warning "No such GPO found in $domain";  }
        

        # Create a hash table for fast GPO lookups later in the report.
        # Hash table key is the policy path which will match the gPLink attribute later.
        # Hash table value is the GPO object with properties for reporting.
        $GPOsHash = @{}
        ForEach ($GPO in $GPOs) {
            $GPOsHash.Add($GPO.Path,$GPO)
        }
        
       $OUs = @(Get-ADOrganizationalUnit -Filter * -Server $domain -Properties * | ? {$_.Name -eq "$OUName"})
       #Return if no OUs found with given name
       if(!$OUs) { Write-Warning "No such OU found in $domain"   }

          foreach($OU in $OUs) {
	        
                    if($OU.gPlink){
                            if($OU.gPLink.length -gt 1) {
                                $links = @($OU.gPLink -split {$_ -eq '[' -or $_ -eq ']'} | Where-Object {$_})
                                $GPOmatch = 0                          
                                For( $i = $links.count - 1 ; $i -ge 0 ; $i-- ) {
                                $GPOData = $links[$i] -split {$_ -eq '/' -or $_ -eq ';'}
                                        
                                        if($GPOsHash[$GPOData[2]].DisplayName -match $GPONAME){
                                                $GPOmatch ++               
                                                $report += New-Object -TypeName PSCustomObject -Property @{
                                                
                                                    Name              = $OU.Name;
                                                    DistinguishedName = $OU.distinguishedName;
                                                    Config            = $GPOData[3];
                                                    LinkEnabled       = [bool](!([int]$GPOData[3] -band 1));
                                                    Enforced          = [bool]([int]$GPOData[3] -band 2);
                                                    BlockInheritance  = [bool]($OU.gPOptions -band 1)
                                                    Path              = $GPOData[2];
                                                    GPODisplayName       = $GPOsHash[$GPOData[2]].DisplayName;
                                                    GPOStatus         = $GPOsHash[$GPOData[2]].GPOStatus;
                                                    WMIFilter         = $GPOsHash[$GPOData[2]].WMIFilter.Name;
                                                    CreationTime      = $GPOsHash[$GPOData[2]].CreationTime;
                                                    ModificationTime  = $GPOsHash[$GPOData[2]].ModificationTime
                                                }
                                          }
                                 
                                  }
                                  if($GPOmatch -eq 0){
                                        $report += New-Object -TypeName PSCustomObject -Property @{
                                            Name              = $OU.Name;
                                            DistinguishedName = $OU.distinguishedName;
                                            BlockInheritance  = [bool]($OU.gPOptions -band 1)
                                        }
                                  }
                             
                        
                         }
                         else
                         { 
                            $report += New-Object -TypeName PSCustomObject -Property @{
                                Name              = $OU.Name;
                                DistinguishedName = $OU.distinguishedName;
                                BlockInheritance  = [bool]($OU.gPOptions -band 1)
                             }
                         }
                     }
                    
                     else{ 
                                # No gPLink at this SOM
                                $report += New-Object -TypeName PSCustomObject -Property @{
                                Name              = $OU.Name;
                                DistinguishedName = $OU.distinguishedName;
                                BlockInheritance  = [bool]($OU.gPOptions -band 1)
                                }
                     } 

            }

            
}#LoopEnd

#Format Output
$report | select Name,DistinguishedName,GPODisplayName,Enforced,LinkEnabled,BlockInheritance,WMIFilter | Export-Csv ".\GPOReport_$(Get-Date -Format 'MMddyyyyhhmm').csv" -NoTypeInformation