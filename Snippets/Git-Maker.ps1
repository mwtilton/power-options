Get-Content .\Shelly\PsPaths.ps1

$date = Get-Date
Write-warning "This may take some time to complete."

$pat1 = "^(.*?\beigrp)\b.*?(\bTEST\b)"

$OriginalFileLocation = "C:\Users\LeviWard\Documents\IT\vsCode\Shelly\PsPaths.ps1"

$ARCHIVE = "C:\Users\LeviWard\Documents\GitHub\power-options"
New-Item -Path $ARCHIVE -ItemType directory -ErrorAction SilentlyContinue

$GetBasenameCFG = Get-ChildItem $OriginalFileLocation -Recurse | Where-Object {!$_.PSIsContainer} | ForEach-Object {$_.BaseName}

if($GetContentFile -match $pat1)
{
    Write-Host "Found PS1: $file which matches $pattern"
    #Gets the contents of the initial file and replaces the EIGRP value currently this will be replaced
    #and set to the value of 10
    #You can change this number to any value just make sure to leave in the ${1} inside the "" as that is set to the regex grouping{1} number
    
    (Get-Content $OriginalFileLocation\$file.cfg) -replace $pattern, "'${1} ' $ReplacementValue[0]"  -join "`r`n" | Set-Content $ARCHIVE\$file.cfg -Force

} 
Else {Write-Host "$File does not match regex $pattern."}