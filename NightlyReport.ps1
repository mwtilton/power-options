$date = Get-Date
#Incoming
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$incoming = [Microsoft.VisualBasic.Interaction]::InputBox("Incoming","Incoming")

#Outgoing
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$outgoing = [Microsoft.VisualBasic.Interaction]::InputBox("Outgoing","Outgoing")

#Time
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$passtime = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the time","Time")

$year = $date.ToString("yyyy")
$dateout1 = $date.ToString("MM-MMM-yyyy").ToUpper()
$dateout2 = $date.ToString("dd-MMM-yyyy").ToUpper()

$dateout3 = $date.ToString("MM - MMMMM")
$dateout4 = $date.ToString("MM MMMM yyyy")
$dateout5 = $date.ToString("MM - MMM").ToUpper()
$dateout6 = $date.ToString("yyyy-MM-dd")

Write-host "$year" -BackgroundColor Black -ForegroundColor White
Write-host "$dateout1" -BackgroundColor Black -ForegroundColor White
Write-host "$dateout2" -BackgroundColor Black -ForegroundColor White

$facmonARCHIVE = "X:\path\to\dated\folder1\$year\$dateout4"
$upsmonARCHIVE = "X:\path\to\dated\folder2\$year\$dateout5"
$passdownARCHIVE = "X:\path\to\dated\folder3\$year\$dateout3"

Write-Host "$dateout1 facmon Folder creation" -ForegroundColor Magenta
New-Item -Path $facmonARCHIVE -ItemType directory -ErrorAction SilentlyContinue
Write-Host "$dateout1 upsmon Folder creation" -ForegroundColor Magenta
New-Item -Path $upsmonARCHIVE -ItemType directory -ErrorAction SilentlyContinue
Write-Host "$dateout1 passdown Folder creation" -ForegroundColor Magenta
New-Item -Path $passdownARCHIVE -ItemType directory -ErrorAction SilentlyContinue

$getFacmon = Get-ChildItem -Path X:\path\to\scanned\doc*.* -Filter *.pdf | % { $_.basename, $_.LastWriteTime -gt (Get-Date).AddHours(1) } | select -last 1 -Skip 2
$getUpsmon = Get-ChildItem -Path X:\path\to\scanned\doc*.* -Filter *.pdf | % { $_.basename, $_.LastWriteTime -gt (Get-Date).AddHours(1) } | select -last 1 -skip 1
$getPassdown = Get-ChildItem -Path X:\path\to\scanned\doc*.* -Filter *.pdf | % { $_.basename, $_.LastWriteTime -gt (Get-Date).AddHours(1) } | select -last 1

Write-Host "Set destinations." -ForegroundColor Magenta
$facmonDEST = "X:\path\to\archive\dated\folder1\$year\$dateout4\SMFM-$dateout2.pdf"
$upsmonDEST = "X:\path\to\archive\dated\folder2\$year\$dateout5\PWR-$dateout2.pdf"
$passdownDEST = "X:\path\to\archive\dated\folder3\$year\$dateout3\SMPD-$dateout6 $passtime.pdf"

Copy-Item -Path X:\Scans\$getFacmon.pdf -Destination $facmonDEST -Recurse
Copy-Item -Path X:\Scans\$getUpsmon.pdf -Destination $upsmonDEST -Recurse
Copy-Item -Path X:\Scans\$getPassdown.pdf -Destination $passdownDEST -Recurse


#Create outlook Object: Facmon
$Outlook = New-Object -comObject  Outlook.Application
$Mail = $Outlook.CreateItem(0)
$Mail.Recipients.Add("email") 

#Add the text part I want to display first
$Mail.Subject = "Facmon Reports $dateout2"
$Mail.Body = "Hey Client,

Here’s the report for $dateout2.

From,
Your name
Title"
$Mail.Attachments.add($facmonDEST)

#Create outlook Object: Facmon
$Outlook = New-Object -comObject  Outlook.Application
$Mail1 = $Outlook.CreateItem(0)
$Mail1.Recipients.Add("email") 

#Add the text part I want to display first
$Mail1.Subject = "UPS Mon Reports $dateout2"
$Mail1.Body = "Hey Client,

Here’s the Report for $dateout2.

From,
Your Name
Title"
$Mail1.Attachments.add($upsmonDEST)

#Create outlook Object: Facmon
$Outlook = New-Object -comObject  Outlook.Application
$Mail2 = $Outlook.CreateItem(0)
$Mail2.Recipients.Add("email") 

#Add the text part I want to display first
$Mail2.Subject = "Report for $dateout2 $passtime"
$Mail2.Body = "Hey Cleint,

Here’s the report for $dateout2. Outgoing is $outgoing and incoming is $incoming.

From,
Your name
Title"
$Mail2.Attachments.add($passdownDEST)

Write-Warning "Sending emails."
$mail.Send()
$mail1.Send()
$mail2.Send()

$date2 = Get-Date
$Time = New-TimeSpan -Start $date -End $date2
Write-Host "This took $time to run."