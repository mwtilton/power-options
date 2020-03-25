[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

$objForm = New-Object Windows.Forms.Form 

$objForm.Text = "Select a Date" 
$objForm.Size = New-Object Drawing.Size @(500,210) 
$objForm.StartPosition = "CenterScreen"
$objForm.Topmost = $True
$objForm.KeyPreview = $True

$objForm.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") 
        {
            $dtmDate = $objCalendar.SelectionStart
            $objForm.Close()
        }
    })

$objForm.Add_KeyDown({
    if ($_.KeyCode -eq "Escape") 
        {
            $dtmDate = "Nothing selected"
            $objForm.Close()
        }
    })

$objCalendar = New-Object System.Windows.Forms.MonthCalendar 
$objCalendar.ShowTodayCircle = $true
$objCalendar.MaxSelectionCount = 1
$objForm.Controls.Add($objCalendar) 

$objForm.Add_Shown({$objForm.Activate()})  
[void] $objForm.ShowDialog() 

if ($dtmDate)
    {
        Write-Host "Date selected: $dtmDate"
    }



#Set adddays
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$custID = [Microsoft.VisualBasic.Interaction]::InputBox("Enter customer ID","Customer ID")
$datevalue1 = [Microsoft.VisualBasic.Interaction]::InputBox("Enter start date","Start Date")
$datevalue2 = [Microsoft.VisualBasic.Interaction]::InputBox("Enter end date","End Date")
write-Host "Start is set to: $datevalue1"  -ForegroundColor Green

$date1 = (Get-Date).Date($datevalue1)
$dateout1 = $date1.ToString("yyyy-MM-dd")
Write-host "$dateout1" -ForegroundColor Green

$date2 = (Get-Date).Date($datevalue2)
$dateout2 = $date2.ToString("yyyy-MM-dd")
Write-host "$dateout2" -ForegroundColor Green


$outMergedBioReport = "C:\Users\LeviWard\Documents\IT\Powershell\$custID-Biometric-Report $dateout1 thru $dateout2.csv"
$TPoutMergedBioReport = Test-Path $outMergedBioReport
if($TPoutMergedBioReport -ne $true)
{
    Write-Host "Running Biometric Report Merge Script" -ForegroundColor Magenta
    Write-Warning "This may take up to an hour to complete."



    $getBioReportCSVbasename = Get-ChildItem "C:\Users\LeviWard\Documents\IT\Powershell\*.csv" -Attributes archive  | % { $_.basename }
    foreach($file in $getBioReportCSVbasename)
    {

    #############################################
    #CSV Merger
        function Merge-CSVFiles { 
        [cmdletbinding()] 
        param( 
            [string[]]$csvFiles, 
            [string]$outMergedBioReport
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
        $Output | select name, activity | Export-Csv -Path $outMergedBioReport  -nti
        } 
        Merge-CSVFiles -csvFiles "C:\Users\LeviWard\Documents\IT\Powershell\$file.csv"
    }
}
Else
{
    Write-Warning "Old merge file existed, please rerun script to create a new merge file."
    Remove-Item -Path $outMergedBioReport
}




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



####################### 
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