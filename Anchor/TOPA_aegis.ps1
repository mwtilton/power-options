#Set adddays
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$datevalue = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Add-days")
write-Host "Adddays is set to: $datevalue"  -ForegroundColor Green

$date1 = (Get-Date).AddDays($datevalue)
$dateout1 = $date1.ToString("yyyyMMdd")
Write-host "$dateout1" -ForegroundColor Green

$date2 = (Get-Date).AddDays($datevalue)
$dateout2 = $date2.ToString("MMdd")
Write-host "$dateout2" -ForegroundColor Green

$date3 = (Get-Date).AddDays($datevalue)
$dateout3 = $date3.ToString("yyyy")
Write-host "$dateout3" -ForegroundColor Green

$agcARCHIVE = "\\SRV\Docfiles\Applications\STAT\Daily\3_Aegis_Claims\$dateout3 AGC"
$agsARCHIVE = "\\SRV\Docfiles\Applications\STAT\Daily\3_Aegis_DailyStat\$dateout3 AGS"
$tpaARCHIVE = "\\SRV\Docfiles\Applications\STAT\Daily\3_Topa_Claims\$dateout3 TPA"
$tpARCHIVE = "\\SRV\Docfiles\Applications\STAT\Daily\3_Topa_DailyStat\$dateout3 TP"

Write-Host "$dateout3 AGC Folder creation" -ForegroundColor Magenta
New-Item -Path $agcARCHIVE -ItemType directory -ErrorAction SilentlyContinue
sleep 1
Write-Host "$dateout3 AGS Folder creation" -ForegroundColor Magenta
New-Item -Path $agsARCHIVE -ItemType directory -ErrorAction SilentlyContinue
sleep 1
Write-Host "$dateout3 TPA Folder creation" -ForegroundColor Magenta
New-Item -Path $tpaARCHIVE -ItemType directory -ErrorAction SilentlyContinue
sleep 1
Write-Host "$dateout3 Tp Folder creation" -ForegroundColor Magenta
New-Item -Path $tpARCHIVE -ItemType directory -ErrorAction SilentlyContinue
sleep 1

Write-Host "Get child-items." -ForegroundColor Magenta
$getagc = Get-ChildItem "\\ANC-SRV2\Docfiles\Applications\STAT\Daily\$dateout1.AGC" -ErrorAction SilentlyContinue
Write-Host "Filename: $getagc" -ForegroundColor Magenta

$getags = Get-ChildItem "\\ANC-SRV2\Docfiles\Applications\STAT\Daily\$dateout1.AGS" -ErrorAction SilentlyContinue
Write-Host "Filename: $getags" -ForegroundColor Magenta

$getTPA = Get-ChildItem "\\ANC-SRV2\Docfiles\Applications\STAT\Daily\$dateout1.TPA" -ErrorAction SilentlyContinue
Write-Host "Filename: $getTPA" -ForegroundColor Magenta

$getTP = Get-ChildItem "\\ANC-SRV2\Docfiles\Applications\STAT\Daily\Tp$dateout2" -Attributes archive | % { $_.FullName }
$getTPzipname = Get-ChildItem "\\ANC-SRV2\Docfiles\Applications\STAT\Daily\Tp$dateout2*" -Attributes archive  | % { $_.basename } 
Write-Host "Filename: $getTPzipname" -ForegroundColor Magenta

Write-Host "Set destinations." -ForegroundColor Magenta
$destAGC = "\\ANC-SRV2\Docfiles\Applications\STAT\Daily\3_Aegis_Claims\$dateout3 AGC\$dateout1.zip"
$destAGS = "\\ANC-SRV2\Docfiles\Applications\STAT\Daily\3_Aegis_DailyStat\$dateout3 AGS\$dateout1.zip"
$destTPA = "\\ANC-SRV2\Docfiles\Applications\STAT\Daily\3_Topa_Claims\$dateout3 TPA\$dateout1.zip"
$destTP = "\\ANC-SRV2\Docfiles\Applications\STAT\Daily\3_Topa_DailyStat\$dateout3 TP\$getTPzipname.zip"

#Compression AGC
Write-Host "Compressing AGC" -ForegroundColor Magenta
Compress-Archive -path $getAGC -DestinationPath $destAGC

#Compression AGS
Write-Host "Compressing AGS" -ForegroundColor Magenta
Compress-Archive -path $getags -DestinationPath $destAGS

#Compression TPA
Write-Host "Compressing TPA" -ForegroundColor Magenta
Compress-Archive -path $getTPA -DestinationPath $destTPA 

#Compression TP
Write-Host "Compressing TP" -ForegroundColor Magenta
Compress-Archive -path $getTP -DestinationPath $destTP

#Emails and smtp server
$fromemail = "mtilton@anchorgeneral.com"
$smtp = "anchorgeneral-com.mail.protection.outlook.com"

#anchorclaims@aegisfirst.com
Write-Host "Aegis Claims Email" -ForegroundColor Yellow
#Send-MailMessage -To anchorclaims@aegisfirst.com -From $fromemail -Subject "Aegis Claims" -SmtpServer $smtp -Body "Aegis Claims File $dateout1" -Attachments $destAGC

#DMcpherson@aegisfirst.com
Write-Host "Aegis Daily Stat Email" -ForegroundColor Yellow
#Send-MailMessage -To DMcpherson@aegisfirst.com -From $fromemail -Subject "Aegis Daily Stat" -SmtpServer $smtp -Body "Aegis Daily Stat $dateout1" -Attachments $destAGS

#Upload <upload@topa-ins.com>
Write-Host "Topa Claims/Daily Stat Email" -ForegroundColor Yellow
#Send-MailMessage -To helpdesk@anchorgeneral.com -From helpdesk@anchorgeneral.com -Subject "Topa Claims/Daily Stat" -SmtpServer $smtp -Body "Topa Claims/Daily Stat $dateout1" -Attachments $destTPA, $destTP
