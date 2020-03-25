$user = "LeviWard"

cd C:\Users\$user\Documents\report\output\

Remove-Item "C:\Users\$user\Documents\report\output\*" -Recurse -Force -ErrorAction SilentlyContinue
Write-Warning "Previous files existed. ALL Old files deleted."

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
    
    Write-Host "Date selected: $DateFormat1 thru $DateFormat2" -ForegroundColor White -BackgroundColor Black
}

Write-Host "Filename is set."  -ForegroundColor White -BackgroundColor Black

$outMergedBioReport = "C:\Users\$user\Documents\report\output\out.csv"
$TPoutMergedBioReport = Test-Path $outMergedBioReport
if($TPoutMergedBioReport -ne $true)
{
    

    Remove-Item "C:\Users\$user\Documents\report\output\*" -Recurse -Force
    
    
    Write-Warning "This may take some time to complete."
    

    ######################
    #merger
    <#
    Start of merge process. Must have a customer ID and date in order to perform merge.

    #$custID-DCAccess$DateFormat1 thru $DateFormat2
    #>
    Write-Host "Running .CSV file merge process." -ForegroundColor White -BackgroundColor Black
    $getBioReportCSVbasename = Get-ChildItem "C:\Users\$user\Documents\report\*.csv" | % { $_.basename }
    
    foreach($file in $getBioReportCSVbasename)
    {

        Write-Host "$file" -BackgroundColor Black -ForegroundColor White
        #"Date/Time", "Zones / Devices / Slave Servers","Event Subtype","User","Workstation"
        Import-Csv "C:\Users\$user\Documents\report\$file.csv" | select "Date/Time", "Zones / Devices / Slave Servers","Event Subtype","User","Workstation"| Sort-Object "Date/Time" | Export-Csv -Path $outMergedBioReport -Append -NoTypeInformation -Force 
        
   
    } 
    
    Function Release-Ref ($ref) 
    {
        #WTF is true?
        ([System.Runtime.InteropServices.Marshal]::ReleaseComObject(
        [System.__ComObject]$ref) -gt 0)
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    Function ConvertCSV-ToExcel
    {
    <#   
        .SYNOPSIS  
            Converts one or more CSV files into an excel file.
     
        .DESCRIPTION  
            Converts one or more CSV files into an excel file. Each CSV file is imported into its own worksheet with the name of the
            file being the name of the worksheet.
       
        .PARAMETER inputfile
            Name of the CSV file being converted
  
        .PARAMETER output
            Name of the converted excel file
       
        .EXAMPLE  
        Get-ChildItem *.csv | ConvertCSV-ToExcel -output 'report.xlsx'
  
        .EXAMPLE  
        ConvertCSV-ToExcel -inputfile 'file.csv' -output 'report.xlsx'
    
        .EXAMPLE      
        ConvertCSV-ToExcel -inputfile @("test1.csv","test2.csv") -output 'report.xlsx'
    
        #>
     
        #Requires -version 2.0  
        [CmdletBinding(
            SupportsShouldProcess = $True,
            ConfirmImpact = 'low',
	        DefaultParameterSetName = 'file'
            )]
        Param (    
            [Parameter(
                ValueFromPipeline=$True,
                Position=0,
                Mandatory=$True,
                HelpMessage="Name of CSV/s to import")]
            [ValidateNotNullOrEmpty()]
            [array]$inputfile,
            [Parameter(
                ValueFromPipeline=$False,
                Position=1,
                Mandatory=$True,
                HelpMessage="Name of excel file output")]
            [ValidateNotNullOrEmpty()]
            [string]$output    
            )

    Begin {     
        #Configure regular expression to match full path of each file
        [regex]$regex = "^\w\:\\"
    
        #Find the number of CSVs being imported
        $count = ($inputfile.count -1)
   
        #Create Excel Com Object
        $excel = new-object -com excel.application
    
        #Disable alerts
        $excel.DisplayAlerts = $false

        #Show Excel application
        $excel.Visible = $false

        #Add workbook
        $workbook = $excel.workbooks.Add()

        #Remove other worksheets
        #$workbook.worksheets.Item(2).delete()
        #After the first worksheet is removed,the next one takes its place
        #$workbook.worksheets.Item(2).delete()   

        #Define initial worksheet number
        $i = 1
        }

    Process {
        ForEach ($input in $inputfile) {
            #If more than one file, create another worksheet for each file
            If ($i -gt 1) {
                $workbook.worksheets.Add("sheet1") | Out-Null
            }
            #Use the first worksheet in the workbook (also the newest created worksheet is always 1)
            $worksheet = $workbook.worksheets.Item(1)
            #Add name of CSV as worksheet name
            $worksheet.name = "$((GCI $input).basename)"

            #Open the CSV file in Excel, must be converted into complete path if not already done
            If ($regex.ismatch($input)) {
                $tempcsv = $excel.Workbooks.Open($input) 
            }
            ElseIf ($regex.ismatch("$($input.fullname)")) {
                $tempcsv = $excel.Workbooks.Open("$($input.fullname)") 
            }    
            Else {    
                $tempcsv = $excel.Workbooks.Open("$($pwd)\$input")      
            }
            $tempsheet = $tempcsv.Worksheets.Item(1)
            #Copy contents of the CSV file
            $tempSheet.UsedRange.Copy() | Out-Null
            #Paste contents of CSV into existing workbook
            $Worksheet.Paste()  
            
            #Close temp workbook
            $tempcsv.close()

            #Select all used cells
            $range = $worksheet.UsedRange

            ####################################
            #Format sheet


$worksheet.SetBackgroundPicture("C:\Users\LeviWard\Documents\IT\Powershell\scalematrix.png")
$objRange = $Worksheet.UsedRange
$objRange.Interior.ColorIndex = 16
$objRange.Font.Size = 14
$objRange.Font.Name = "Cambria"


#Create a Title for the first worksheet
$row = 1
$Column = 1

#Save the initial row so it can be used later to create a border
$initalRow = $row


#General Page Setup
$worksheet.Columns("A").NumberFormat = "MM-dd-yyyy"
$worksheet.Columns(1).ColumnWidth = 18
$worksheet.Columns(2).ColumnWidth = 15
$worksheet.Columns(3).ColumnWidth = 10
#user col
$worksheet.Columns(4).ColumnWidth = 16
#workstation col
$worksheet.Columns(5).ColumnWidth = 18
$worksheet.Columns.HorizontalAlignment = -4131

$worksheet.pageSetup.RightHeader = "Date: &D"
$worksheet.pageSetup.Color = 56

$worksheet.PageSetup.LeftHeaderPicture.FileName = "C:\Users\LeviWard\Documents\IT\Powershell\scalematrix.png"
$worksheet.PageSetup.LeftHeaderPicture.Height = 100 
$worksheet.PageSetup.LeftHeaderPicture.Width = 200


$worksheet.PageSetup.Zoom = $false
$worksheet.PageSetup.FitToPagesTall = 1
$worksheet.PageSetup.FitToPagesWide = 1
$worksheet.PageSetup.PrintArea = "A1:E50"




# Add the headers to the worksheet
$headers = "Access Date","Zone","Device","User","Workstation"
$headers | foreach {
    $worksheet.Cells.Item($row,$column)=$_
    $worksheet.Cells.Item($row,$column).Interior.ColorIndex =43
    $worksheet.Cells.Item($row,$column).Font.Bold = $True
    $worksheet.Cells.Item($row,$column).Font.Size = 14
    $column++
}

            $i++
            } 
        }        

    End {
        #Save spreadsheet
        $workbook.saveas("$pwd\$output")

        Write-Host "File $output saved." -ForegroundColor White -BackgroundColor Black

        #Close Excel
        $excel.quit()  

        #Release processes for Excel
        $a = Release-Ref($range)
        }
    }  
    
    Get-ChildItem "C:\Users\$user\Documents\report\output\*.csv" | ConvertCSV-ToExcel -output "$custID-DCAccess$DateFormat1-thru-$DateFormat2.xlsx"
    



    
    #######################################################
    #Excel to .pdf
    
    $path = "C:\Users\$user\Documents\report\output"
    $xlFixedFormat = "Microsoft.Office.Interop.Excel.xlFixedFormatType" -as [type] 
    $excelFiles = Get-ChildItem -Path $path -Include *.xlsx -recurse 
    $objExcel = New-Object -ComObject excel.application 
    $objExcel.visible = $false 
    foreach($wb in $excelFiles) 
        { 
            $filepath = Join-Path -Path $path -ChildPath ("$custID-DCAccess$DateFormat1-thru-$DateFormat2.pdf") 
            $workbook = $objExcel.workbooks.open($wb.fullname, 3) 
            $workbook.Saved = $true 
            Write-Host "Converting excel file to .PDF" -ForegroundColor White -BackgroundColor Black
            $workbook.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $filepath) 
            $objExcel.Workbooks.close() 
        } 
    $objExcel.Quit()
    ii $filepath
    Write-Host "Process completed." -ForegroundColor White -BackgroundColor Black
    }
Else
{
    Remove-Item "C:\Users\$user\Documents\report\output\*" -Recurse -Force
    Write-Warning "Filename already exists. Old file deleted. Please rerun script."
}




