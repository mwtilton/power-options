$file = "C:\Temp\qo.txt"
$regfile = "C:\Temp\qo.txt"

$getFile = (get-content -Path $file).Split(",")


Write-Host $getFile[0] $getRegfile[1]