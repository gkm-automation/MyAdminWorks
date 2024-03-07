#Change Baseline CSV path
$base = Import-Csv "C:\Users\kg212483\Downloads\results\base.csv" | Where-Object { $_.label -Like "ClaimsFNOL*" } | Select-Object label,@{n='elapsed';e={$_.elapsed/1000}}
#Change Current CSV Path
$current = Import-Csv "C:\Users\kg212483\Downloads\results\current.csv" | Where-Object { $_.label -Like "ClaimsFNOL*" } | Select-Object label,@{n='elapsed';e={$_.elapsed/1000}}
$results = @()
foreach($record in $current){
    $basevalue = ($base | Where-Object { $_.label -eq $record.label }).'elapsed'[0]
    $diffValue = $basevalue - $($record.'elapsed')
    $obj = [PSCustomObject]@{
        Label = $record.label
        'Difference(Secs)' = $diffValue
    }
$results += $obj
}

Write-Output ($results | ft)
