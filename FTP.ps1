Import-Module PSFTP 
Set-FTPConnection -Credentials anonymous -Server ftp://10.31.2.143 -Session MyTestSession -UsePassive 
$Session = Get-FTPConnection -Session MyTestSession 

Get-FTPChildItem -Session $Session -Path /pub -Recurse