
function Check-BiosPwd {

param([parameter(Mandatory=$True)]
[String]$computerName=$env:COMPUTERNAME
)
$output = @()
#check Remote machine connectivity

if($bios=Get-WmiObject -Class Win32_BIOS -ComputerName $computerName){

    Switch -Regex ($bios.Manufacturer){

        ^Dell { 
                $passwords = Get-CimInstance -Namespace root\dcim\sysman -classname dcim_biospassword 
                $passwords | foreach-Object {
                $attrval = $_.AttributeName
        
                if ($_.IsSet -match "True") {
                $output += "$attrval is set on $computerName."
        
                }
                elseif ($_.IsSet -match "False") {
                $output += "$attrval is not set on $computerName."
        
                }
        
        
               }
               }
        Default  { "Didn't match anything..."  }

    }#switch
}#if
else{
    Write-Host "WMI Namespace of $computerName inaccessible..!" -ForegroundColor Yellow
    } 
    

    $output 

}