#Set customer ID
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$custID = [Microsoft.VisualBasic.Interaction]::InputBox("Enter customer ID","Customer ID")

###########################################################################
#Calendar

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object Windows.Forms.Form 

$form.Text = "Select a Date" 
$form.Size = New-Object Drawing.Size @(243,230) 
$form.StartPosition = "CenterScreen"

$calendar = New-Object System.Windows.Forms.MonthCalendar 
$calendar.ShowTodayCircle = $False
$calendar.MaxSelectionCount = 90
$form.Controls.Add($calendar) 

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(38,165)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(113,165)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$form.Topmost = $True

$result = $form.ShowDialog() 

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $Date1 = $Calendar.SelectionStart
    $DateFormat1 = $Date1.ToString("yyy-MM-dd")
    $Date2 = $Calendar.SelectionEnd
    $DateFormat2 = $Date2.ToString("yyy-MM-dd")
    
    Write-Host "Date selected: $DateFormat1 thru $DateFormat2"
}

Write-Host "Filename is set to: $custID-Report $DateFormat1 thru $DateFormat2.csv"  -ForegroundColor Green
$outTxt = Out-File "C:\path\to\report\output\$custID - $DateFormat1 thru $DateFormat2.txt"
$getOutTxt = Get-ChildItem $outTxt -Attributes archive  | % { $_.basename }

######################
#merger
<#
Start of merge process. Must have a customer ID and date in order to perform merge.
#>

$outMergedBioReport = "C:\path\to\report\output\$getOutTxt.csv"
$TPoutMergedBioReport = Test-Path $outMergedBioReport
if($TPoutMergedBioReport -ne $true)
{
    Write-Host "Running Biometric Report Merge Script" -ForegroundColor Magenta
    Write-Warning "This may take some time to complete."
    
    $getBioReportCSVbasename = Get-ChildItem "C:\path\to\report\*.csv" -Attributes archive  | % { $_.basename }
    foreach($file in $getBioReportCSVbasename)
    {

    #############################################
    #CSV Merger
        function Merge-CSVFiles { 
        [cmdletbinding()] 
        param( 
            [string[]]$csvFiles, 
            [string]$outMergedFile
        ) 
        $Output = @(); 
        foreach($CSV in $csvFiles) { 
            if(Test-Path $CSV) { 
                Write-Warning "Starting merge process for: $file.csv" 
                $FileName = [System.IO.Path]::GetFileName($CSV) 
                $temp = Import-CSV -Path $CSV | select *, @{Expression={$FileName};Label="FileName"} 
                $Output += $temp 
                } 
                else 
                { 
                Write-Warning "$CSV : No such file found" 
                } 
 
            }
        $Output | select "Date/Time", "Zones / Devices / Slave Servers","Event Subtype","User","Workstation" | Export-Csv -Path $outMergedBioReport -nti
        } 
         
    }
        Merge-CSVFiles -csvFiles "C:\path\to\report\$file.csv"
        
        Write-Warning "Removing old $getOutTxt.txt"
        Remove-Item "C:\path\to\report\output\*.txt"
}
Else
{
    Write-Warning "Else Warning."
}


#######################
#.xlsx to .pdf
<#
$path = $outMergedBioReport
$xlFixedFormat = "Microsoft.Office.Interop.Excel.xlFixedFormatType" -as [type] 
$excelFiles = Get-ChildItem -Path $path -include *.xls, *.xlsx -recurse 
$objExcel = New-Object -ComObject excel.application 
$objExcel.visible = $false 
foreach($wb in $excelFiles) 
{ 
 $filepath = Join-Path -Path $path -ChildPath ($wb.BaseName + ".pdf") 
 $workbook = $objExcel.workbooks.open($wb.fullname, 3) 
 $workbook.Saved = $true 
"saving $filepath" 
 $workbook.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $filepath) 
 $objExcel.Workbooks.close() 
} 
$objExcel.Quit()
#>


####################### 
#SQL
<# 
.SYNOPSIS 
Runs a T-SQL script. 
.DESCRIPTION 
Runs a T-SQL script. Invoke-Sqlcmd2 only returns message output, such as the output of PRINT statements when -verbose parameter is specified 
.INPUTS 
None 
    You cannot pipe objects to Invoke-Sqlcmd2 
.OUTPUTS 
   System.Data.DataTable 
.EXAMPLE 
Invoke-Sqlcmd2 -ServerInstance "MyComputer\MyInstance" -Query "SELECT login_time AS 'StartTime' FROM sysprocesses WHERE spid = 1" 
This example connects to a named instance of the Database Engine on a computer and runs a basic T-SQL query. 
StartTime 
----------- 
2010-08-12 21:21:03.593 
.EXAMPLE 
Invoke-Sqlcmd2 -ServerInstance "MyComputer\MyInstance" -InputFile "C:\MyFolder\tsqlscript.sql" | Out-File -filePath "C:\MyFolder\tsqlscript.rpt" 
This example reads a file containing T-SQL statements, runs the file, and writes the output to another file. 
.EXAMPLE 
Invoke-Sqlcmd2  -ServerInstance "MyComputer\MyInstance" -Query "PRINT 'hello world'" -Verbose 
This example uses the PowerShell -Verbose parameter to return the message output of the PRINT command. 
VERBOSE: hello world 
.NOTES 
Version History 
v1.0   - Chad Miller - Initial release 
v1.1   - Chad Miller - Fixed Issue with connection closing 
v1.2   - Chad Miller - Added inputfile, SQL auth support, connectiontimeout and output message handling. Updated help documentation 
v1.3   - Chad Miller - Added As parameter to control DataSet, DataTable or array of DataRow Output type 
 
function Invoke-Sqlcmd2 
{ 
    [CmdletBinding()] 
    param( 
    [Parameter(Position=0, Mandatory=$true)] [string]$ServerInstance, 
    [Parameter(Position=1, Mandatory=$false)] [string]$Database, 
    [Parameter(Position=2, Mandatory=$false)] [string]$Query, 
    [Parameter(Position=3, Mandatory=$false)] [string]$Username, 
    [Parameter(Position=4, Mandatory=$false)] [string]$Password, 
    [Parameter(Position=5, Mandatory=$false)] [Int32]$QueryTimeout=600, 
    [Parameter(Position=6, Mandatory=$false)] [Int32]$ConnectionTimeout=15, 
    [Parameter(Position=7, Mandatory=$false)] [ValidateScript({test-path $_})] [string]$InputFile, 
    [Parameter(Position=8, Mandatory=$false)] [ValidateSet("DataSet", "DataTable", "DataRow")] [string]$As="DataRow" 
    ) 
 
    if ($InputFile) 
    { 
        $filePath = $(resolve-path $InputFile).path 
        $Query =  [System.IO.File]::ReadAllText("$filePath") 
    } 
 
    $conn=new-object System.Data.SqlClient.SQLConnection 
      
    if ($Username) 
    { $ConnectionString = "Server={0};Database={1};User ID={2};Password={3};Trusted_Connection=False;Connect Timeout={4}" -f $ServerInstance,$Database,$Username,$Password,$ConnectionTimeout } 
    else 
    { $ConnectionString = "Server={0};Database={1};Integrated Security=True;Connect Timeout={2}" -f $ServerInstance,$Database,$ConnectionTimeout } 
 
    $conn.ConnectionString=$ConnectionString 
     
    #Following EventHandler is used for PRINT and RAISERROR T-SQL statements. Executed when -Verbose parameter specified by caller 
    if ($PSBoundParameters.Verbose) 
    { 
        $conn.FireInfoMessageEventOnUserErrors=$true 
        $handler = [System.Data.SqlClient.SqlInfoMessageEventHandler] {Write-Verbose "$($_)"} 
        $conn.add_InfoMessage($handler) 
    } 
     
    $conn.Open() 
    $cmd=new-object system.Data.SqlClient.SqlCommand($Query,$conn) 
    $cmd.CommandTimeout=$QueryTimeout 
    $ds=New-Object system.Data.DataSet 
    $da=New-Object system.Data.SqlClient.SqlDataAdapter($cmd) 
    [void]$da.fill($ds) 
    $conn.Close() 
    switch ($As) 
    { 
        'DataSet'   { Write-Output ($ds) } 
        'DataTable' { Write-Output ($ds.Tables) } 
        'DataRow'   { Write-Output ($ds.Tables[0]) } 
    } 
 
} 

Invoke-Sqlcmd2

#>


#Fatal Error Handling
Try{
    If($_.Exception.ToString().Contains("something")){
        Write-Host " already exists. Skipping!" -ForegroundColor DarkGreen
    }
    Else{

        Write-host $_.Exception -ForegroundColor Yellow
    }
}
Catch{
    $_ | fl * -force
    $_.InvocationInfo.BoundParameters | fl * -force
    $_.Exception
}

#one line error thrower
if ($?) {throw}

#runAs Admin stuffs
#R equires -RunAsAdministrator

#DSC
Get-DscResource * | Select -ExpandProperty Properties | ft -AutoSize
Get-DscResource * -Syntax


#Igonore SSL
add-type @"
    using System.Net;
    using System.Security.Cryptography.X509Certificates;
    public class TrustAllCertsPolicy : ICertificatePolicy {
        public bool CheckValidationResult(
            ServicePoint srvPoint, X509Certificate certificate,
            WebRequest request, int certificateProblem) {
            return true;
        }
    }
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy


<#

    https://techcommunity.microsoft.com/t5/ITOps-Talk-Blog/PowerShell-Basics-Finding-Your-Way-in-the-PowerShell-Console/ba-p/300935
    PowerShell Basics: Finding Your Way in the PowerShell Console

#>

#Get-Command
get-command -CommandType Function
get-command -CommandType Function -name Get-*

#Firewall
get-command -name *firewall*
get-command -name *netfirewallrule

#Get-Help
get-help set-netfirewallrule
get-help set-netfirewallrule -examples
help Set-NetFirewallRule -Full
Help set-netfirewallrule -Parameter RemoteAddress

#Get-Member
get-service | Get-Member
(Get-Service | Get-Member | Where-Object -Property Membertype -EQ Property).count
Get-Service | Get-Member | Where-Object -Property Membertype -EQ Property
get-service | format-table -Property Name,Status,ServicesDependedOn

Get-Variable

#Prompt Information
(Get-Command prompt).definition
(Get-Command Prompt).ScriptBlock

#RawUI settings
$(Get-Host).UI.RawUI

#Check if you are in a nested powershell session
$NestedPromptLevel

#get running history of commands used
Get-History

#Making a new profile
if (!(Test-Path -Path $profile)) {New-Item -ItemType File -Path $profile -Force}

#profile information
$PROFILE | select *

#Git
git checkout --track origin/<branch_name>

# Connection to mysql database from dataset
import dataset
from CreatePointsDatabase import create_points_database

db = dataset.connect('mssql+pymssql://mssql+pymssql://user:password@server:port/database')


# Sending Email
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

# Moving from one repo to another
1. Check out the existing repository from Bitbucket:
$ git clone https://USER@bitbucket.org/USER/PROJECT.git

2. Add the new Github repository as upstream remote of the repository checked out from Bitbucket:

$ cd PROJECT
$ git remote add upstream https://github.com:USER/PROJECT.git

3. Checkout and track any extra branches you want to push to the new repo
$ git checkout --track origin/dev

4. Push all branches (below: just master) and tags to the Github repository:

$ git push upstream master
$ git push --tags upstream

# PS ModulePath
Import-Module '$($env:PSModulePath).Split(;)[1]\UCSD' -Force -ErrorAction Stop;

# .lnk file to run ps script
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -executinpolicy unrestricted -file "F:\ipconfig.ps1"


<#
Set-ExecutionPolicy RemoteSigned
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking
Start-transcript
get-mailbox -identity |FL
connect-msolservice
install-module msonline
";

        //holding
        public static string PSholding = @"
Get-VM | where-object {$_.PowerState -eq ""PoweredOn"" -or $_.PowerState -eq ""Suspended"" -and $_.Name -notlike ""*vCenter*""}| select Name
#
# DO NOT UNCOMMENT. HIGHLY RADIOACTIVE ISOTOPES!!!
#
#Get-VM | where-object {$_.PowerState -eq ""PoweredOn"" -or $_.PowerState -eq ""Suspended"" -and $_.Name-notlike ""*vCenter*""} | Stop-VM -Confirm:$false ##|shutdown-vmguest -Confirm:$false
#
	
function Get-FolderByPath{
  <# .SYNOPSIS Retrieve folders by giving a path .DESCRIPTION The function will retrieve a folder by it's path. The path can contain any type of leave (folder or datacenter). .NOTES Author: Luc Dekens .PARAMETER Path The path to the folder. This is a required parameter. .PARAMETER Path The path to the folder. This is a required parameter. .PARAMETER Separator The character that is used to separate the leaves in the path. The default is '/' .EXAMPLE PS> Get-FolderByPath -Path ""Folder1/Datacenter/Folder2""
.EXAMPLE
  PS> Get-FolderByPath -Path ""Folder1>Folder2"" -Separator '>'
#>
 
  param(
  [CmdletBinding()]
  [parameter(Mandatory = $true)]
  [System.String[]]${Path
    },
  [char]${Separator
} = '/'
  )
 
  process{
    if((Get-PowerCLIConfiguration).DefaultVIServerMode -eq ""Multiple""){
      $vcs = $defaultVIServers
    }
    else{
      $vcs = $defaultVIServers[0]
    }
 
    foreach($vc in $vcs){
      foreach($strPath in $Path){
        $root = Get-Folder -Name Datacenters -Server $vc
        $strPath.Split($Separator) | %{
          $root = Get-Inventory -Name $_ -Location $root -Server $vc -NoRecursion
          if((Get-Inventory -Location $root -NoRecursion | Select -ExpandProperty Name) -contains ""vm""){
            $root = Get-Inventory -Name ""vm"" -Location $root -Server $vc -NoRecursion
          }
        }
        $root | where {$_ -is [VMware.VimAutomation.ViCore.Impl.V1.Inventory.FolderImpl]}|%{
          Get-Folder -Name $_.Name -Location $root.Parent -NoRecursion -Server $vc
        }
      }
    }
  }
}


 Get-vCD-VM-Detail
PowerShell
# Get a list of vCloud Orgs vDCs
$ListOrgVDC = Get-OrgVdc | Sort-Object -Property Name

# Write the list and prompt for input
write-host """"
write-host ""$ListOrgVDC""
$OrgVDC = read-host ""Please Enter the Name of the OrgVDC""
$vms = Get-OrgVdc -Name $OrgVDC | get-civm
$objects = @()

# Jakes Smarts
foreach($vm in $vms)
{
 $hardware = $vm.ExtensionData.GetVirtualHardwareSection()
 $diskMB = (($hardware.Item | where {$_.resourcetype.value -eq ""17""}) | %{$_.hostresource[0].anyattr[0].""#text""} | Measure-Object -Sum).sum
 $row = New-Object PSObject -Property @{""vapp"" = $vm.vapp; ""name""=$vm.Name;""cpuCount""=$vm.CpuCount;""memoryGB""=$vm.MemoryGB;""storageGB""=($diskMB/1024)}
 $objects += $row
}


# Use select object to get the column order right. Sort by vApp. Force table formatting and auto-width.
$objects | select-Object name, vapp, cpuCount, memoryGB, storageGB | Sort-Object -Property vapp | Format-Table -AutoSize

# Also Export results to CVS for further processing
$objects | Export-Csv ""$OrgVDC.csv"" -NoTypeInformation -UseCulture

	
# Get a list of vCloud Orgs vDCs
$ListOrgVDC = Get-OrgVdc | Sort-Object -Property Name

# Write the list and prompt for input
write-host """"
write-host ""$ListOrgVDC""
$OrgVDC = read-host ""Please Enter the Name of the OrgVDC""
$vms = Get-OrgVdc -Name $OrgVDC | get-civm
$objects = @()
 
# Jakes Smarts
foreach($vm in $vms)
{
 $hardware = $vm.ExtensionData.GetVirtualHardwareSection()
 $diskMB = (($hardware.Item | where {$_.resourcetype.value -eq ""17""}) | %{$_.hostresource[0].anyattr[0].""#text""} | Measure-Object -Sum).sum
 $row = New-Object PSObject -Property @{""vapp"" = $vm.vapp; ""name""=$vm.Name;""cpuCount""=$vm.CpuCount;""memoryGB""=$vm.MemoryGB;""storageGB""=($diskMB/1024)}
 $objects += $row
}
 
 
# Use select object to get the column order right. Sort by vApp. Force table formatting and auto-width.
$objects | select-Object name, vapp, cpuCount, memoryGB, storageGB | Sort-Object -Property vapp | Format-Table -AutoSize
 
# Also Export results to CVS for further processing
$objects | Export-Csv ""$OrgVDC.csv"" -NoTypeInformation -UseCulture

Get-vORG-VM-Detail
PowerShell
# Get a list of vCloud Orgs
$ListOrg = Get-Org | Sort-Object -Property Name
# Write the list and prompt for input
write-host ""$ListOrg""
write-host """"
$Org = read-host ""Please Enter the Name of the Org""
$vms = Get-Org -Name $Org | get-civm
$objects = @()

# Jakes Smarts
foreach($vm in $vms)
{
 $hardware = $vm.ExtensionData.GetVirtualHardwareSection()
 $diskMB = (($hardware.Item | where {$_.resourcetype.value -eq ""17""}) | %{$_.hostresource[0].anyattr[0].""#text""} | Measure-Object -Sum).sum
 $row = New-Object PSObject -Property @{""vapp"" = $vm.vapp; ""name""=$vm.Name;""cpuCount""=$vm.CpuCount;""memoryGB""=$vm.MemoryGB;""storageGB""=($diskMB/1024)}
 $objects += $row
}

# Use select object to get the column order right. Sort by vApp. Force table formatting and auto-width.
$objects | select-Object name, vapp, cpuCount, memoryGB, storageGB | Sort-Object -Property vapp | Format-Table -AutoSize

# Also Export results to CVS for further processing
$objects | Export-Csv ""$Org.csv"" -NoTypeInformation -UseCulture


	
# Get a list of vCloud Orgs
$ListOrg = Get-Org | Sort-Object -Property Name
# Write the list and prompt for input
write-host ""$ListOrg""
write-host """"
$Org = read-host ""Please Enter the Name of the Org""
$vms = Get-Org -Name $Org | get-civm
$objects = @()
 
# Jakes Smarts
foreach($vm in $vms)
{
 $hardware = $vm.ExtensionData.GetVirtualHardwareSection()
 $diskMB = (($hardware.Item | where {$_.resourcetype.value -eq ""17""}) | %{$_.hostresource[0].anyattr[0].""#text""} | Measure-Object -Sum).sum
 $row = New-Object PSObject -Property @{""vapp"" = $vm.vapp; ""name""=$vm.Name;""cpuCount""=$vm.CpuCount;""memoryGB""=$vm.MemoryGB;""storageGB""=($diskMB/1024)}
 $objects += $row
}
 
# Use select object to get the column order right. Sort by vApp. Force table formatting and auto-width.
$objects | select-Object name, vapp, cpuCount, memoryGB, storageGB | Sort-Object -Property vapp | Format-Table -AutoSize
 
# Also Export results to CVS for further processing
$objects | Export-Csv ""$Org.csv"" -NoTypeInformation -UseCulture
#>