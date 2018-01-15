#Works in PS v2+
#
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
New-Item -Path $ARCHIVE -ItemType directory -ErrorAction SilentlyContinue

$GetBasenameCFG = Get-ChildItem C:\temp -Filter "*.cfg" | Where-Object {!$_.PSIsContainer} | ForEach-Object {$_.BaseName}

#Go through each file in the folder and does something
ForEach ($file in $GetBasenameCFG) {
    #This just writes the filename to the host.
    Write-Host "Found CFG: $file"
    
    #Gets the contents of the initial file
    $GetContentFile = Get-Content $OriginalFileLocation\$file.cfg
    
    #Checks to see if it matches the first REGEX
    $pattern = $pat1
    If ($GetContentFile -match $pattern)
    {
        Write-Host "$File updated." -BackgroundColor Black -ForegroundColor White
        #You can change this number to any character (letters or numbers) just make sure to leave in the ${1} inside the "" as that is set to the regex grouping{1} number
        ($GetContentFile) -replace $pattern, '${1} 10' -join "`r`n" | Set-Content $ARCHIVE\$file.cfg -Force
    }
    #Checks file if it matches other REGEX
    #You will need to change this to the other patterns from the one you are checking for
    #
    #i.e.
    #
    #If you wanted to check for pat2
    #Then change these variables after -match too pat1 and pat3
    elseif (($GetContentFile -match $pat2) -and ($GetContentFile -match $pat3)) {
        Write-Warning "$File does match other possible regex"
    }
    #Files that dont match any REGEX at all
    else{
        Write-Host "$File does not match regex " $pattern -BackgroundColor Black -ForegroundColor Red
    }

  
}

#This opens a single file from the "Modified" file location
#this was only for testing purposes you can remove if needed
#ii C:\temp\Modified\sampleConfig1.cfg

#Length of time it takes to run script
$date2 = Get-Date
$Time = New-TimeSpan -Start $date -End $date2
Write-Host "This took $time to run."