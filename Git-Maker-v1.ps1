$date = Get-Date

#REGEX pattern
$pat1 = "(x\:\\)IT\\NOC"

$OriginalFileLocation = "C:\Users\LeviWard\Documents\IT\vsCode\Shelly\SM"

$ARCHIVE = "C:\Users\LeviWard\Documents\GitHub\power-options"
#$ARCHIVE = "$OriginalFileLocation\Archive"
New-Item -Path $ARCHIVE -ItemType directory -ErrorAction SilentlyContinue

$GetBasenameCFG = Get-ChildItem $OriginalFileLocation -Filter *.ps1 | Where-Object {!$_.PSIsContainer} | ForEach-Object {$_.BaseName}

foreach ($file in $GetBasenameCFG)
{   
    $GetContentFile = Get-Content $OriginalFileLocation\$file.ps1
    if($GetContentFile -match $pat1)
    {
        Write-Host "Found PS1: $file which matches $pat1"
        ($GetContentFile) -replace $pat1, "${1}c:\Path\To"  -join "`r`n" | Set-Content $ARCHIVE\$file.ps1 -Force

    } 
    Else {
        Write-Host "$File does not match regex $pat1."
    }
}

