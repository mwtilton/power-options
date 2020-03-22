#Start timer
cls

$ErrorActionPreference = 'SilentlyContinue'

$date = Get-Date
$year = $date.ToString("yyyy")

Write-host "{                                                                                       }" -BackgroundColor Black -ForegroundColor Green 
Write-host "                                                                                         " -BackgroundColor Black -ForegroundColor Green 
Write-host "                                [     ScaleMatrix    ]                                   " -BackgroundColor Black -ForegroundColor Green
Write-host "                                                                                         " -BackgroundColor Black -ForegroundColor Green
Write-host "{                                                                                       }" -BackgroundColor Black -ForegroundColor Green

#Set Dept
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$dept = [Microsoft.VisualBasic.Interaction]::InputBox("Select Department:
1 - Sales
2 - Engineering

","Department")


If($dept -eq 1)
{
    $Department = "Sales"
}
elseif($dept -eq 2)
{
    $Department = "Engineering"
}
Else
{
    Exit
}

#Date format for Week '04-JAN-2018'
$dateFormat1 = $date.ToString("dd-MMM-yyyy").ToUpper()

#Date format for Header '01/04/2017'
$dateFormat2 = $date.ToString("MM/dd/yyyy")

#Date format for Header '01/04/2017'
$dateFormat3 = $date.ToString("MM - MMM").ToUpper()

#Is excel running
Get-process -name excel -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

#File location for CSVs to be pulled from Digitus
$getExportFolder = "c:\Path\To\Biometric Audit Reports\Exports\*"

#Test the path to see if anything is in the folder
$tpExportsFolder = Test-Path $getExportFolder -Filter *.csv
If($tpExportsFolder -eq $True)
{
    #Header file
    $tpOutFiles2 = Test-Path $outChangeHeaders -ErrorAction SilentlyContinue
    if($tpOutFiles2 -eq $True)
    {
        Write-Host "The file: " -BackgroundColor Black -ForegroundColor Green -NoNewline
        write-host $outChangeHeaders -BackgroundColor Black -ForegroundColor White -NoNewline
        Write-Host " has been" -BackgroundColor Black -ForegroundColor Green -NoNewline
        Write-host " deleted." -BackgroundColor Black -ForegroundColor Red
        Remove-Item $outChangeHeaders
    }
    Else
    {
        
    }

    #Out First file
    $tpOutFiles4 = Test-Path $outMergedBioReportFirst -ErrorAction SilentlyContinue
    if($tpOutFiles4 -eq $True)
    {
        Write-Host "The file: " -BackgroundColor Black -ForegroundColor Green -NoNewline
        write-host $outMergedBioReportFirst -BackgroundColor Black -ForegroundColor White -NoNewline
        Write-Host " has been" -BackgroundColor Black -ForegroundColor Green -NoNewline
        Write-host " deleted." -BackgroundColor Black -ForegroundColor Red
        Remove-Item $outMergedBioReportFirst
    }
    Else
    {
        
    }
    #out Last file
    $tpOutFiles5 = Test-Path $outMergedBioReportLast -ErrorAction SilentlyContinue
    if($tpOutFiles5 -eq $True)
    {
        Write-Host "The file: " -BackgroundColor Black -ForegroundColor Green -NoNewline
        write-host $outMergedBioReportLast -BackgroundColor Black -ForegroundColor White -NoNewline
        Write-Host " has been" -BackgroundColor Black -ForegroundColor Green -NoNewline
        Write-host " deleted." -BackgroundColor Black -ForegroundColor Red
        Remove-Item $outMergedBioReportLast
    }
    Else
    {
        
    }


    #Set customer ID
    $custID = '00001'

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
    $calendar.MaxSelectionCount = 100
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
        $dates = $calendar.SelectionRange
        $datesArray = @()
        $dateNameArray = @()
        $dateHeaderArray = @()
        $rr = @()

        $n = 1

        $dateHeaderArray += ""
        
        

        for($d = $dates.Start; $d -le $dates.End; $d = $d.AddDays(1))
        {
            $datesArray += $d.ToString("MM/dd")
            $dateNameArray += $d.ToString("dddd" + "-MM-dd-yy-")
            $dateHeaderArray += $d.ToString("dddd" + " MM/dd/yy")
            
            $rr += $d.ToString("MM/dd/yyyy")
            $rr += $d.ToString("MM/dd/yyyy")
            

            $n++
        }

         
        Write-Host "Dates Selected: " -BackgroundColor Black -ForegroundColor Green -NoNewline
        Write-Host $datesArray -BackgroundColor Black -ForegroundColor White
        
        Write-Host "Day Names Selected: " -BackgroundColor Black -ForegroundColor Green -NoNewline
        Write-Host $dateNameArray -BackgroundColor Black -ForegroundColor White

    }
    Else{
        exit
    }




    #Get the folder name for the client
    $getClientFolder = Get-ChildItem "c:\Path\To\Biometric Audit Reports\$custID*"
    $getClientFolderName = Get-ChildItem "c:\Path\To\Biometric Audit Reports\$custID*" | % {$_.basename}
    
    #Main Folder
    Write-Host "Main Folder location set to: " -BackgroundColor Black -ForegroundColor Green -NoNewline
    Write-Host $getClientFolder -BackgroundColor Black -ForegroundColor White 

    #Header Date
    Write-Host "Header Date selected: " -BackgroundColor Black -ForegroundColor Green -NoNewline
    Write-Host "$DateFormat1" -BackgroundColor Black -ForegroundColor White 
    
    #Subject Date
    Write-Host "Subject Date selected: " -BackgroundColor Black -ForegroundColor Green -NoNewline
    Write-host "$DateFormat2" -BackgroundColor Black -ForegroundColor White
    
    #Creating Archive folder for client
    Write-Host "Creating Archive folders..." -BackgroundColor Black -ForegroundColor White -NoNewline
    
    Write-Host "CSVArchive" -BackgroundColor Black -ForegroundColor Magenta -NoNewline
    Write-Host "..." -BackgroundColor Black -ForegroundColor White -NoNewline
    $newClientCSVArchive = New-Item -Path $getClientFolder\CSVArchive\$year\$dateFormat3 -ItemType directory -ErrorAction SilentlyContinue
    
    Write-Host "ExcelArchive" -BackgroundColor Black -ForegroundColor Magenta -NoNewline
    Write-Host "..." -BackgroundColor Black -ForegroundColor White -NoNewline
    $newClientExcelArchive = New-Item -Path $getClientFolder\ExcelArchive\$year\$dateFormat3 -ItemType directory -ErrorAction SilentlyContinue

    Write-Host "PDFArchive" -BackgroundColor Black -ForegroundColor Magenta -NoNewline
    Write-Host "..." -BackgroundColor Black -ForegroundColor White -NoNewline
    $newClientPDFArchive = New-Item -Path $getClientFolder\PDFArchive\$year\$dateFormat3 -ItemType directory -ErrorAction SilentlyContinue



    $tpClientCSVArchiveFolder = Test-Path $getClientFolder\CSVArchive\$year\$dateFormat3
    If($tpClientCSVArchiveFolder -eq $True)
    {
        Write-Host "Done." -BackgroundColor Black -ForegroundColor White

        ######################
        #CSV to Excel merge Process
        <#
        Start of merge process. Must have a customer ID and date in order to perform merge.
        $custID-DCAccess$DateFormat1 thru $DateFormat2
        #>
        
        #Get names for all files in the Export folder    
        Write-Host "Running " -BackgroundColor Black -ForegroundColor White -NoNewline
        Write-host ".CSV" -BackgroundColor Black -ForegroundColor Magenta -NoNewline    
        Write-host " file merge process." -ForegroundColor White -BackgroundColor Black 
        $getBioReportCSVbasename = Get-ChildItem "c:\Path\To\Biometric Audit Reports\Exports\*.csv" | % { $_.basename }
    
        Write-Warning "This may take some time to complete."

        #File location for Output of merged .csv files into one CSV file
        $outMergedBioReportFirst = "$getClientFolder\CSVArchive\$year\outFirst.csv"
        $outMergedBioReportLast = "$getClientFolder\CSVArchive\$year\outLast.csv"
        $outChangeHeaders = "$getClientFolder\CSVArchive\$year\$dateFormat3\outHeaders.csv"
        
        foreach($file in $getBioReportCSVbasename)
        {
            #New Headers "Date", "Location", "User"
            $newHeaders = "Date", "Location", "EventType", "EventSubType", "Status", "LastName", "User"
            Get-Content "c:\Path\To\Biometric Audit Reports\Exports\$file.csv" -Encoding Default | select -Skip 1 | ConvertFrom-Csv -UseCulture -Header $newHeaders | Export-Csv -Path $outChangeHeaders -Append -NoTypeInformation -Force
            Write-Host "Replacing Headers for: " -BackgroundColor Black -ForegroundColor Green -NoNewline
            Write-Host "$file" -BackgroundColor Black -ForegroundColor White
            Write-Host "Moving File: " -BackgroundColor Black -ForegroundColor Green -NoNewline
            Write-Host "$file" -BackgroundColor Black -ForegroundColor White -NoNewline
            Write-Host " to " -BackgroundColor Black -ForegroundColor Green -NoNewline
            Write-Host "$getClientFolder\CSVArchive\$year\$dateFormat3" -BackgroundColor Black -ForegroundColor White
            #Move-Item -Path "c:\Path\To\Biometric Audit Reports\Exports\$file.csv" -Destination $getClientFolder\CSVArchive\$year\$dateFormat3 -ErrorAction SilentlyContinue
        }

        $nameArray = @()
        $lastNAmeArray = @()     

        Write-Host "Running " -BackgroundColor Black -ForegroundColor Green -NoNewline
        Write-Host "NameFinder.exe " -BackgroundColor Black -ForegroundColor cyan
        
        Import-Csv $outChangeHeaders | Group-Object "user"| select "Name" | foreach {   
            $nameArray += $_.name
             
        }
        

        Import-Csv $outChangeHeaders | Group-Object "LastName"| select "Name" | foreach {
            $lastNameArray += $_.name

        }

        $nameArray | Foreach {
            Write-Host $_ $lastNAmeArray[[array]::IndexOf($nameArray, $_)]
        }


        foreach ($user in $nameArray)
        {   
            Write-Host "$user" -BackgroundColor Black -ForegroundColor Yellow

            foreach ($d in $datesArray)
            {
                #Write-Host $user $d $datesArray[[array]::IndexOf($nameArray, $_)]
                #Import-Csv $outChangeHeaders | select "User", "Date", "Location","LastName" | where {$_.date -match "$d" -and $_.user -like "$user"} | select -First 1 | Format-List
                Import-Csv $outChangeHeaders | select "User", "Date", "Location","LastName" | where {$_.date -match "$d" -and $_.user -like "$user"} | select -First 1 | Export-Csv -Path $outMergedBioReportFirst -Append -NoTypeInformation -Force
            }
          
            
            foreach ($d in $datesArray)
            {
                #Write-Host $user $d
                #Import-Csv $outChangeHeaders | select "User", "Date", "Location","LastName" | where {$_.date -match "$d" -and $_.user -like "$user"} | select -Last 1 | Format-List
                Import-Csv $outChangeHeaders | select "User", "Date", "Location","LastName" | where {$_.date -match "$d" -and $_.user -like "$user"} | select -Last 1 | Export-Csv -Path $outMergedBioReportLast -Append -NoTypeInformation -Force
            }
        }

        ii $outMergedBioReportFirst
        
        <#
        Import-Csv $outChangeHeaders | Group-Object "user"| select "Name" | foreach {   
            
            Write-Host $_.name  -BackgroundColor Black -ForegroundColor Yellow -NoNewline
            
            $nameArray += $_.name
            Write-Host "..." -BackgroundColor Black -ForegroundColor White -NoNewline
            foreach ($d in $datesArray)
            {
                
                #Import-Csv $outChangeHeaders -Encoding Default | select "User", "Date", "Location","LastName" | where {$_.date -like "*$d*" -and $nameArray -contains $_.user} | select -First 1 | Format-List
                Import-Csv $outChangeHeaders -Encoding Default | select "User", "Date", "Location","LastName" | where {$_.date -like "*$d*" -and $nameArray -contains $_.user} | select -First 1 | Export-Csv -Path $outMergedBioReportFirst -Append -NoTypeInformation -Force
                #Import-Csv $outChangeHeaders -Encoding Default | select "User", "Date", "Location","LastName" | where {$_.date -like "*$d*" -and $nameArray -contains $_.user} | select -Last 1 | Format-List
                Import-Csv $outChangeHeaders -Encoding Default | select "User", "Date", "Location","LastName" | where {$_.date -like "*$d*" -and $nameArray -contains $_.user} | select -Last 1 | Export-Csv -Path $outMergedBioReportLast -Append -NoTypeInformation -Force
                
            }
        }
        #>

        Write-Host "Done." -BackgroundColor Black -ForegroundColor Green 

        


        Write-Warning "Finished Import export."

        
        Function Release-Ref ($ref) 
        {
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

            Begin 
            {     
                #Configure regular expression to match full path of each file
                [regex]$regex = "^\w\:\\"
    
                #Find the number of CSVs being imported
                $count = ($inputfile.count -1)
   
                #Create Excel Com Object
                $excel = new-object -com excel.application
    
                #Disable alerts
                $excel.DisplayAlerts = $false

                #Show Excel application
                $excel.Visible = $False

                #Add workbook
                $workbook = $excel.workbooks.Add()
                          
                $sheetNamesArray = @()
                $sheetNamesArray2 = @()

                #$ErrorActionPreference = 'SilentlyContinue'

                
                
                $dateNameArray | foreach {
                    
                    $sheetNamesArray += $_ + "IN"
                    $sheetNamesArray += $_ + "OUT"

                    $sheetNamesArray2 += $_ + "IN"
                    $sheetNamesArray2 += $_ + "OUT"
                    
                }
                $sheetNamesArray += "Export"
                          
                $sheetNamesArray | foreach {
                    
                    $workbook.worksheets.Add() | Out-Null  
                }
                
                $i = 1
                $sheetNamesArray | foreach {
                    
                    #Write-Host $_ -BackgroundColor Black -ForegroundColor Cyan
                    $excel.sheets.item($i).name = $_
                    $i++
                }

                #$ErrorActionPreference = 'Continue'
                
                $workbook.worksheets.item("Sheet1").Delete()
                $workbook.worksheets.item("Sheet2").Delete()
                $workbook.worksheets.item("Sheet3").Delete()
                #$workbook.worksheets.item("*Sheet*").Delete()
                

                #Define initial worksheet number
                $i = $sheetNamesArray.count + $inputfile.count + 1
            }

            Process 
            {

                    
                ForEach ($input in $inputfile) 
                {
                    #If more than one file, create another worksheet for each file
                    If ($i -gt 1) {
                        $workbook.worksheets.Add() | Out-Null
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
                    #   
                    $counter = $sheetNamesArray.count + $input.count
                    
                    #Write-Host $count -BackgroundColor Black -ForegroundColor Cyan
                                        
                    If($excel.sheets.item($counter).name -contains "Export")
                    {

                        Write-Host "Working on " -BackgroundColor Black -ForegroundColor White -NoNewline
                        Write-host $excel.sheets.item($counter).name -BackgroundColor Black -ForegroundColor Yellow -NoNewline
                        Write-host " sheet." -BackgroundColor Black -ForegroundColor White
                        $sheetToDest = $workbook.sheets.item("Export")
                        
                        
                        #Export sheet headers
                        $row = 1
                        $column = 1
                        
                        $emptyRowNum = 1
                        $emptyColNum = 2

                        
                        $dateHeaderArray | foreach {
                        

                            Write-Host "Merging: " -BackgroundColor Black -ForegroundColor Green -NoNewline
                            Write-Host "(" -BackgroundColor Black -ForegroundColor White -NoNewline
                            Write-host $row -BackgroundColor Black -ForegroundColor Cyan -NoNewline
                            Write-Host ", " -BackgroundColor Black -ForegroundColor White -NoNewline
                            Write-Host $column -BackgroundColor Black -ForegroundColor Cyan -NoNewline
                            Write-Host ") | (" -BackgroundColor Black -ForegroundColor White -NoNewline
                            Write-Host $emptyRowNum -BackgroundColor Black -ForegroundColor Cyan -NoNewline 
                            Write-host ", " -BackgroundColor Black -ForegroundColor White -NoNewline
                            Write-Host "$emptyColNum)"-BackgroundColor Black -ForegroundColor Cyan
                        

                            $sheetToDest.Cells.Item($row,$column) = $_
                            $sheetToDest.Cells.Item($emptyRowNum,$emptyColNum) = ""

                            $sheetToDest.Cells.Item($row,$column).Font.Bold = $True
                            $sheetToDest.Cells.Item($row,$column).Font.Color = 1
                            $sheetToDest.Cells.Item($row,$column).Interior.ColorIndex = 43
                            $sheetToDest.Cells.Item($row,$column).Font.Size = 12
                            
                            #merge cells
                            $mergeCells = $sheetToDest.Range($sheetToDest.Cells.Item($row,$column), $sheetToDest.Cells.Item($emptyRowNum,$emptyColNum))
                            $mergeCells.MergeCells = $True


                            $emptyColNum++
                            $emptyColNum++
                            $column++
                            $column++
                        }
                        Write-Host "Finished Cell Merge process." -BackgroundColor Black -ForegroundColor White
                        
                        Write-Host "Doing " -BackgroundColor Black -ForegroundColor White -NoNewline
                        Write-Host "Excel " -BackgroundColor Black -ForegroundColor DarkGreen -NoNewline
                        Write-Host "things." -BackgroundColor Black -ForegroundColor White

                        Write-Host "Creating Arrays" -BackgroundColor Black -ForegroundColor White
                        #out/in Col
                        $row = 2
                        $column = 1
                        
                        $exportHeaders = @()

                        $exportHeaders += "First Name"
                        $exportHeaders += "Last Name"
                        
                        $soloExportHeaders = @()
                        
                        $inOutArray = ($sheetNamesArray2.count)/2

                        
                        #Write-Host $inOutArray -BackgroundColor Black -ForegroundColor Cyan

                        1..$inOutArray | foreach {
                            $exportHeaders += "In"
                            $exportHeaders += "Out"

                            $soloExportHeaders += "In"
                            $soloExportHeaders += "Out"
                        }
                        

                        $exportHeaders | foreach {
                            $sheetToDest.Cells.Item($row,$column)=$_
                            $sheetToDest.Cells.Item($row,$column).Font.Bold = $True
                            $sheetToDest.Cells.Item($row,$column).Font.Color = 1
                            $sheetToDest.Cells.Item($row,$column).Interior.ColorIndex = 40
                            $sheetToDest.Cells.Item($row,$column).Font.Size = 12
                            $column++
                        }
                        
                        
                        #Users
                        Write-Host "Adding users" -BackgroundColor Black -ForegroundColor White
                        $row = 3
                        $column = 1
                        
                        $nameArray | foreach {
                            $sheetToDest.Cells.Item($row,$column) = $_
                            
                            $row++

                        }
                        
                        
                        #user last names
                        $row = 3
                        $column = 2

                        #vlookup 
                        Write-Host "vLookup" -BackgroundColor Black -ForegroundColor White
                        $rowVl = 3
                        $columnVl = 3
                        $n = 3
                        $lastNAmeArray | foreach {
                            $sheetToDest.Cells.Item($row,$column)=$_
                            
                            
                            $sheetNamesArray2 | foreach {
                            
                                
                                $sheetToDest.Cells.Item($rowVl,$columnVl).Formula = "=IFERROR(VLOOKUP(`$A$n,`'$_`'!`$A:`$B,2,FALSE),`"`")"
                                $sheetToDest.Cells.Item($rowVl,$columnVl).NumberFormat = "hh:mm"
                                
                                $columnVl++
                                

                            }
                            $n++
                            $columnVl = 3
                            $rowVl++
                            $row++
                        }

                        #Setting Line Weight
                        Write-Host "Setting line weight." -BackgroundColor Black -ForegroundColor White
                        $inputRange = "1"
                                                
                        $rows = $sheetToDest.UsedRange.Rows.Count
                        $columns = $sheetToDest.UsedRange.Columns.Count

                        #Write-Host $rows -BackgroundColor Black -ForegroundColor Red 
                        #Write-Host $rows -BackgroundColor Black -ForegroundColor Red 
                        
                        <#
                        $sdsa = foreach ($col in "A") {
                            Write-Host $col -BackgroundColor Black -ForegroundColor Magenta
                            $excel.WorksheetFunction.CountIf($sheetToDest.Range($col + "1:" + $col + $rows), "<>")
                        }
                        #>
                        #$sdsaMod = $sdsa + $inputRange
                        
                        

                        Function Convert-NumberToA1 { 
                            <# 
                            .SYNOPSIS 
                            This converts any integer into A1 format. 
                            .DESCRIPTION 
                            See synopsis. 
                            .PARAMETER number 
                            Any number between 1 and 2147483647 
                            #> 
   
                            Param([parameter(Mandatory=$true)] 
                                [int]$number) 
 
                            $a1Value = $null 
                            While ($number -gt 0) { 
                            $multiplier = [int][system.math]::Floor(($number / 26)) 
                            $charNumber = $number - ($multiplier * 26) 
                            If ($charNumber -eq 0) { $multiplier-- ; $charNumber = 26 } 
                            $a1Value = [char]($charNumber + 64) + $a1Value 
                            $number = $multiplier 
                            } 
                            Return $a1Value 
                        }
                        $columnLetter = Convert-NumberToA1 -number $columns
                        
                        #Write-Host $columnLetter -BackgroundColor Black -ForegroundColor DarkGray -NoNewline
                        #Write-Host $rows -BackgroundColor Black -ForegroundColor DarkGray


                        $dataRange = $sheetToDest.Range("A$inputRange : $columnLetter$rows")
                        7..12 | ForEach {
                            $dataRange.Borders.Item($_).LineStyle = 1
                            $dataRange.Borders.Item($_).Weight = 3
                        }
                        
                        $range = $sheet.usedRange
                        $range.EntireColumn.AutoFit() | out-null

                        
                        #Rename Sheet to Audit
                        #
                        #Warning: Dont put anything else after this line.
                        $workbook.worksheets.item('Export').name = "Audit"
                        
                        Write-host $excel.sheets.item($counter).name -BackgroundColor Black -ForegroundColor Yellow -NoNewline
                        Write-Host " is done. " -BackgroundColor Black -ForegroundColor White
                    }
                    Else
                    {
                        Write-host "Sheet name: " -BackgroundColor Black -ForegroundColor White -NoNewline
                        Write-Host $excel.sheets.item($counter).name -BackgroundColor Black -ForegroundColor Red

                    }
                    

                    
                    #Create a Title for the first worksheet
                    $row = 1
                    $Column = 1

                    #Save the initial row so it can be used later to create a border
                    $initalRow = $row

                    # Add the headers to the worksheet
                    <#
                    $headers = "FirstName", "Date","Location", "LastName"
                    $headers | foreach {
                    $worksheet.Cells.Item($row,$column)=$_
                        $worksheet.Cells.Item($row,$column).Font.Bold = $True
                        $worksheet.Cells.Item($row,$column).Font.Size = 12
                        $column++
                    }
                    $i++
                    #>

                    #Sort Headers
                    Write-Host "Sorting headers." -BackgroundColor Black -ForegroundColor White
                    
                    $setSheetToUsedrange = $worksheet.UsedRange
                    $setSheetToRange = $worksheet.Range("A1")
                    $setSheetToRange2 = $worksheet.Range("D1")
                    [void]$setSheetToUsedrange.Sort($setSheetToRange, 1,$null,$null,1,$null,1,1)
                    [void]$setSheetToUsedrange.Sort($setSheetToRange2, 1,$null,$null,1,$null,1,1)


                    $row = 2
                    $column = 2
                    Write-Warning "Starting for Loop."
                    
                    try {
                    
                        $rr.count -eq $sheetNamesArray2.count | Out-Null
                    }
                    catch {
                        Write-Host $rr.count -BackgroundColor Black -ForegroundColor Cyan
                        Write-Host $sheetNamesArray2.count -BackgroundColor Black -ForegroundColor Yellow

                    }

                    
                    for ($i = 2; $i -le $row; $i++)
                    {
                        $swisher = $excel.sheets.item(1).name

                        switch -Wildcard ($swisher)
                        {
                            "outFirst" 
                            {

                                if ([string]::IsNullOrEmpty($worksheet.cells.Item($i,2).Value2) -eq $False)
                                {

                                    $cellValue = $worksheet.cells.Item($i,2).Value()
                                    Write-Host "Cell " -BackgroundColor Black -ForegroundColor White -NoNewline
                                    Write-Host "B$i's " -BackgroundColor Black -ForegroundColor Green -NoNewline
                                                                    
                                    Write-Host "Value is: " -BackgroundColor Black -ForegroundColor White -NoNewline
                                    Write-Host $cellValue -BackgroundColor Black -ForegroundColor Green
                            
                                    $rr | foreach {
                                        

                                        If($cellValue -like "*$_*")
                                        {
                                            #Write-Host $_ -BackgroundColor Black -ForegroundColor Cyan -NoNewline
                                            #Write-Host $sheetNamesArray2[[array]::IndexOf($rr, $_)] -BackgroundColor Black -ForegroundColor Yellow
                                            #Write-Host "."

                                            $worksheet.cells.Item($i,1).EntireRow.Copy() | Out-Null
                                            $sheetRR = $workbook.sheets.item($sheetNamesArray2[[array]::IndexOf($rr, $_)])
                                    
                                            #Write-Host $sheetRR -BackgroundColor Black -ForegroundColor Cyan
                                    
                                            $sheetRR.Range("A$i").PasteSpecial() | Out-Null
                                            
                                            $row++
                                        }
                                        Else
                                        {
                                            #Write-Host $sheetNamesArray2[[array]::IndexOf($rr, $_)] -BackgroundColor Black -ForegroundColor Red
                                        }
                                        
                                    }
                            
                            
                                }
                                
                                Else
                                {
                                    
                                    break
                                }
                               
                            }

                            "outLast" 
                            {
                                if ([string]::IsNullOrEmpty($worksheet.cells.Item($i,2).Value2) -eq $False)
                                {

                            

                                    $cellValue = $worksheet.cells.Item($i,2).Value()
                                    Write-Host "Cell " -BackgroundColor Black -ForegroundColor White -NoNewline
                                    Write-Host "B$i's " -BackgroundColor Black -ForegroundColor Green -NoNewline
                                                                    
                                    Write-Host "Value is: " -BackgroundColor Black -ForegroundColor White -NoNewline
                                    Write-Host $cellValue -BackgroundColor Black -ForegroundColor Green
                            
                                    $rr | foreach {



                                        #write-host $sheetNamesArray2[[array]::IndexOf($rr, $_)] -BackgroundColor Black -ForegroundColor Yellow


                                        If($cellValue -like "*$_*")
                                        {
                                            

                                            

                                            #Write-Host $_ -BackgroundColor Black -ForegroundColor Cyan -NoNewline
                                            #Write-Host $sheetNamesArray2[[array]::IndexOf($rr, $_) + 1] -BackgroundColor Black -ForegroundColor Yellow
                                            #Write-Host "."

                                            $worksheet.cells.Item($i,1).EntireRow.Copy() | Out-Null
                                            $sheetRR = $workbook.sheets.item($sheetNamesArray2[[array]::IndexOf($rr, $_) + 1])
                                    
                                            #Write-Host $sheetRR -BackgroundColor Black -ForegroundColor Cyan
                                    
                                            $sheetRR.Range("A$i").PasteSpecial() | Out-Null
                                            
                                            $row++
                                            
                                        }
                                        Else
                                        {
                                            #Write-Host $sheetNamesArray2[[array]::IndexOf($rr, $_)] -BackgroundColor Black -ForegroundColor Red
                                        }
                                        
                                    }
                            
                            
                                }
                                
                                Else
                                {
                                    
                                    break
                                }
                                

                            }
                            default
                            {
                                Write-host "Indeterminable" -BackgroundColor Black -ForegroundColor Magenta
                            }

                        }
                       
                    }
                    Write-Host $excel.sheets.item(1).name -BackgroundColor Black -ForegroundColor Cyan -NoNewline
                    Write-Host " is done." -BackgroundColor Black -ForegroundColor White
                    
                       
                            
                } 
                        
            }
            End 
            {
                #Save spreadsheet
                $workbook.saveas("$output")
                Write-Host "File " -BackgroundColor Black -ForegroundColor Green -NoNewline
                Write-Host "$output" -BackgroundColor Black -ForegroundColor White -NoNewline
                Write-Host " saved." -BackgroundColor Black -ForegroundColor Green

                #Close Excel
                $excel.quit()  

                #Release processes for Excel
                $a = Release-Ref($range)
            }
        } 
        
         
      
        
        $xlsxOutput = "$getClientFolder\ExcelArchive\$year\$dateFormat3\$custID-EmployeeDCAccess-$DateFormat1.xlsx"
        
        ConvertCSV-ToExcel -inputfile @($outMergedBioReportLast, $outMergedBioReportFirst) -output $xlsxOutput
        
        #ii $xlsxOutput

        #Header file
        $tpOutFiles2 = Test-Path $outChangeHeaders
        if($tpOutFiles2 -eq $True)
        {
            Write-Host "The file: " -BackgroundColor Black -ForegroundColor Green -NoNewline
            write-host $outChangeHeaders -BackgroundColor Black -ForegroundColor White -NoNewline
            Write-Host " has been" -BackgroundColor Black -ForegroundColor Green -NoNewline
            Write-host " deleted." -BackgroundColor Black -ForegroundColor Red
            Remove-Item $outChangeHeaders
        }
        Else
        {
            Write-Warning "$outChangeHeaders does not exist."
        }

        #Out First file
        $tpOutFiles4 = Test-Path $outMergedBioReportFirst
        if($tpOutFiles4 -eq $True)
        {
            Write-Host "The file: " -BackgroundColor Black -ForegroundColor Green -NoNewline
            write-host $outMergedBioReportFirst -BackgroundColor Black -ForegroundColor White -NoNewline
            Write-Host " has been" -BackgroundColor Black -ForegroundColor Green -NoNewline
            Write-host " deleted." -BackgroundColor Black -ForegroundColor Red
            Remove-Item $outMergedBioReportFirst
        }
        Else
        {
            Write-Warning "$outMergedBioReportFirst does not exist."
        }
        #out Last file
        $tpOutFiles5 = Test-Path $outMergedBioReportLast
        if($tpOutFiles5 -eq $True)
        {
            Write-Host "The file: " -BackgroundColor Black -ForegroundColor Green -NoNewline
            write-host $outMergedBioReportLast -BackgroundColor Black -ForegroundColor White -NoNewline
            Write-Host " has been" -BackgroundColor Black -ForegroundColor Green -NoNewline
            Write-host " deleted." -BackgroundColor Black -ForegroundColor Red
            Remove-Item $outMergedBioReportLast
        }
        Else
        {
            Write-Warning "$outMergedBioReportLast does not exist."
        }


        #######################################################
        #
        #Excel to .pdf
        <#
        Filename
        $custID-DCAccess-$DateFormat1-thru-$DateFormat2.pdf
        #>

        $objExcel = New-Object -ComObject excel.application
        $objExcel.displayAlerts = $false
        $objExcel.AskToUpdateLinks = $false
        $objExcel.visible = $false 
    
        $excelFiles = Get-ChildItem -Path $getClientFolder\ExcelArchive\$year\$dateFormat3\$custID-EmployeeDCAccess-$DateFormat1.xlsx -Include *.xlsx -recurse | select -Last 1

        $xlFixedFormat = "Microsoft.Office.Interop.Excel.xlFixedFormatType" -as [type] | Out-Null
    
        foreach($wb in $excelFiles) 
        { 
            $filepath = Join-Path -Path $getClientFolder\PDFArchive\$year\$dateFormat3 -ChildPath ("$custID-EmployeeDCAccess-$dateFormat1.pdf") 
            $workbook = $objExcel.workbooks.open($wb.fullname, 3)
            $wsheet =  $workbook.sheets.item('Audit')
                        
            $workbook.Saved = $true 
            Write-Host "Converting" -BackgroundColor Black -ForegroundColor Green -NoNewline
            Write-Host " Excel " -BackgroundColor Black -ForegroundColor DarkGreen -NoNewline
            Write-Host "to" -BackgroundColor Black -ForegroundColor Green -NoNewline
            Write-Host " .PDF" -BackgroundColor Black -ForegroundColor Magenta
            $wsheet.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $filepath)
            
            Write-Host "PDF Process completed." -ForegroundColor White -BackgroundColor Black
        }
        $objExcel.Workbooks.close() 
        $objExcel.Quit()


        $source = "$getClientFolder\ExcelArchive\$year\$dateFormat3\$custID-EmployeeDCAccess-$DateFormat1.xlsx"

        $xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault

        #Open Excel
        $xl = new-object -com excel.application 
        $xl.displayAlerts = $false 
        $xl.AskToUpdateLinks = $false
        $xl.Visible = $false
        
        #open source, readonly 
        $wb2 = $xl.workbooks.open($source, $null, $true)  
        
        #Copy Skyscale Client Summary sheet
        $sh2_wb1 = $wb2.sheets.item("Audit")
        $sh2_wb1.UsedRange.Copy() | Out-Null
        Write-host "Export sheet selected" -BackgroundColor Black -ForegroundColor White

        #Create outlook Object: Facmon
        $Outlook = New-Object -comObject  Outlook.Application
        $Mail = $Outlook.CreateItem(0)
        $Mail.Recipients.Add("roemorales@scalematrix.com") | Out-Null

        #Add the text part I want to display first
        $Mail.Subject = $Department + " Biometric Access Reports " + ($dates.Start).tostring("MM-dd-yyyy") + " - " + ($dates.End).tostring("MM-dd-yyyy") 
        
        #Then Copy the Excel using parameters to format it
        $Mail.Getinspector.WordEditor.Range().PasteSpecial(13)
        
        #Then it becomes possible to insert text before
        $wdDoc = $Mail.Getinspector.WordEditor
        $wdRange = $wdDoc.Range()
                       
        
        $Mail.Attachments.add("$getClientFolder\PDFArchive\$year\$dateFormat3\$custID-EmployeeDCAccess-$dateFormat1.pdf") | Out-Null
        $Mail.Display()

        $xl.workbooks.OpenText($source,437,1,1,1,$True,$True,$False,$False,$True,$False)
        $xl.ActiveWorkbook.SaveAs($source, $xlFixedFormat)

        $xl.Workbooks.Close()
        $xl.Quit()



    }
}




#Get Script speed
$date2 = Get-Date
$Time = New-TimeSpan -Start $date -End $date2
Write-Host "This took" -BackgroundColor Black -ForegroundColor Green -NoNewline 
Write-Host " $time " -BackgroundColor Black -ForegroundColor Magenta -NoNewline
Write-host "to run." -BackgroundColor Black -ForegroundColor Green

