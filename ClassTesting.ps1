$file = "C:\Temp\users.csv"

Class InTime
{
    [String]$firstname
    [String]$lastname
    [datetime]$time
}

$user = New-Object InTime

import-csv $file -Encoding Default  | select -Skip 1 |ForEach {
    
    
    $user.firstname = $_.firstname
    $user.lastname = $_.lastname
    $user.time = $_.time
    
}


$user | Out-GridView