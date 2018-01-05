#Enter in Log ID information
$id = ''

#Enter in Log Name. You can use the Asterisk(*) symbol for wildcards
$Logname = "Application"
$event = Get-EventLog -LogName $Logname -InstanceId $id -Newest 1

#Check Event log for error
if ($event.EntryType -eq "Error")
{
    #region Variables and Arguments
    $date = Get-Date -Format MM/dd/yy
    $users = "Josh@Justic.net" # List of users to email your report to (separate by comma)
    $fromemail = "USERNAME@gmail.com"
    $SMTPServer = "smtp.gmail.com"
    $SMTPPort = "587"
    $SMTPUser = "USERNAME@gmail.com"
    $SMTPPassword = "PASSWORD"
    $ComputerName = gc env:computername
    $EmailSubject = "COMPUTERNAME - New Event Log [Application] $date"
    $MailSubject = $MailSubject -replace('COMPUTERNAME', $ComputerName)
    $Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $SMTPUser, $($SMTPPassword | ConvertTo-SecureString -AsPlainText -Force) 
    $EnableSSL = $true
    $ListOfAttachments = @()
    $Report = @()
    $CurrentTime = Get-Date
    $PCName = $env:COMPUTERNAME
    $EmailBody = $event | ConvertToHtml > elog.htm
    $getHTML = Get-Content "elog.htm"
    #sending email
    send-mailmessage -from $fromemail -to $users -subject $EmailSubject -BodyAsHTML -body $getHTML -priority Normal -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $Credentials
    Remove-Item elog.htm
}
else
{
    write-host "No error found"
    write-host "Here is the log entry that was inspected: $event"
}