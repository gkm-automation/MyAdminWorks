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
       'Difference(Secs)' = $diffValue
       'DeviationThreshold' = $DeviationThreshold

    }
$results += $obj
}

Write-Output ($results | ft)
