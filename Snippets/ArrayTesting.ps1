cls

$array1 = "1","2","3","5","6"
$array2 = "2","3","4","5","6"

For ($i = 0; $i -le $array1.Length; $i++)
{
    if (($array1[$i] -match $array2[$i]) -and ($array1[$i] -ne $null -or $array2[$i] -ne $null))
    {
        Write-host "At: $i there is a match " $array1[$i] " and " $array2[$i] -BackgroundColor Black -ForegroundColor White
    }
    elseif ($array1[$i] -eq $null -or $array2[$i] -eq $null) {
        Write-Host "Done."
        break
    }
    Else
    {
        Write-Host "No Match At: $i between " $array1[$i] " and " $array2[$i] -BackgroundColor Black -ForegroundColor Red
    }
}