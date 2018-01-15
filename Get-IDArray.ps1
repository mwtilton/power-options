#Works in PS v2+
#Clear host screen
Clear-Host

<#
General Information

Should you need to update or change the REGEX values you are looking for 
I would suggest starting with the website below, as I believe you mentioned you 
may need to update locations at a future date

REGEX website:
https://regex101.com/
#>

$date = Get-Date
Write-warning "This may take some time to complete."

#REGEX pattern for eigrp TEST
#Currently this does not affect the "SOMEKEY" value
$pat1 = "^(.*?\beigrp)\b.*?(\bTEST\b)"

#REGEX pattern for bgp [number]
#This is set to select only the first number in the sequence
$pat2 = "^(.*?\bbgp)\b.*?\s[0-9]"

#REGEX for bgp [number] (after remote-as)
#optional remove if the remote-as value needs to remain the same
#Or update this value to future options such as the example you gave East DC to South DC
$pat3 = "^(.*?\bremote-as)\b.*?\s[0-9]"

#Sets the location for the files to be pulled from
$OriginalFileLocation = "C:\Temp"

#Set the archive location
#This will update all your output file locations
#Process will error out if this folder is not created first or the folder does not exist
$ARCHIVE = "C:\Temp\Modified"
#You can remove or comment out this "New-item" cmdlet if you know the archiving folder exists
New-Item -Path $ARCHIVE -ItemType directory -ErrorAction SilentlyContinue

$GetBasenameCFG = Get-ChildItem $OriginalFileLocation | Where-Object {!$_.PSIsContainer} | ForEach-Object {$_.BaseName}

#Go through each file in the folder and does something 
ForEach ($file in $GetBasenameCFG) {
    #This just writes the filename to the host.
    
    $array1 = "$pat1", "$pat2", "$pat3"
    $ReplaceArray = '${1} 20x6','${1} 20x7','${1} 20x8'
    $GetContentFile = Get-Content $OriginalFileLocation\$file.cfg
        
    For ($i = 0; $i -le $array1.Length; $i++)
    {

        if (($GetContentFile -match $array1[$i]) -and ($array1[$i] -ne $null))
        {
            Write-host "$file At: $i there is a match " $array1[$i] -BackgroundColor Black -ForegroundColor White
            ($GetBasenameCFG) | foreach{$_ -replace $array1[$i], $replaceArray[$i]} | Set-Content $file.PSPath
            
            
        }
        elseif (($GetContentFile -notmatch $array1[$i]) -and ($array1[$i] -ne $null)){
            Write-Host "$File does not match regex " $array1[$i] -BackgroundColor Black -ForegroundColor Red
        }
        elseif ($array1[$i] -eq $null) {
            Write-Host "$file Done."
            BREAK   
        }

    }

}

#This opens a single file from the "Modified" file location
#this was only for testing purposes you can remove if needed
explorer $ARCHIVE

#Length of time it takes to run script
$date2 = Get-Date
$Time = New-TimeSpan -Start $date -End $date2
Write-Host "This took $time to run."