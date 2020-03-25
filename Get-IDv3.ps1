#Works in PS v2+
#
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



$GetBasenameCFG = Get-ChildItem $OriginalFileLocation -Recurse | Where-Object {!$_.PSIsContainer} | ForEach-Object {$_.BaseName}

#Go through each file in the folder and does something
ForEach ($file in $GetBasenameCFG) {

        if($GetContentFile -match $pattern)
        {
            Write-Host "Found CFG: $file which matches $pattern"
            #Gets the contents of the initial file and replaces the EIGRP value currently this will be replaced
            #and set to the value of 10
            #You can change this number to any value just make sure to leave in the ${1} inside the "" as that is set to the regex grouping{1} number
            
            (Get-Content $OriginalFileLocation\$file.cfg) -replace $pattern, "'${1} ' $ReplacementValue[0]"  -join "`r`n" | Set-Content $ARCHIVE\$file.cfg -Force

        } 
        Else {Write-Host "$File does not match regex $pattern."}
       
    
}

#This opens a single file from the "Modified" file location
#this was only for testing purposes you can remove if needed
ii $ARCHIVE\sampleConfig2.cfg
ii $ARCHIVE\sampleConfig3.cfg

#Length of time it takes to run script
$date2 = Get-Date
$Time = New-TimeSpan -Start $date -End $date2
Write-Host "This took $time to run."