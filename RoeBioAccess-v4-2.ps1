#Start timer
$date = Get-Date
$year = $date.ToString("yyyy")

#Date format for Week '04-JAN-2018'
$dateFormat1 = $date.ToString("dd-MMM-yyyy").ToUpper()

#Date format for Header '01/04/2017'
$dateFormat2 = $date.ToString("MM/dd/yyyy")

#Date format for Header '01/04/2017'
$dateFormat3 = $date.ToString("MM - MMM").ToUpper()

#Date format for Cells 'HH:MM'


#File location for CSVs to be pulled from Digitus
$getExportFolder = "c:\Path\To\Biometric Audit Reports\Exports\*"

#Test the path to see if anything is in the folder
$tpExportsFolder = Test-Path $getExportFolder -Filter *.csv
If($tpExportsFolder -eq $True)
{
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
    $calendar.MaxSelectionCount = 1
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
        $DateFormat4 = $Date1.ToString("MM")
        
        #First Day selected
        $DateFormat5 = $Date1.ToString("dd")
        $DayName1 = $Date1.ToString("dddd")

        #Second Day Selected
        $DateFormat6 = $Date1.AddDays(1).ToString("dd")
        $DayName2 = $Date1.AddDays(1).ToString("dddd")
        
        #Thrid Day Selected
        $DateFormat7 = $Date1.AddDays(2).ToString("dd")
        $DayName3 = $Date1.AddDays(2).ToString("dddd")
        
        #Fourth Day Selected
        $DateFormat8 = $Date1.AddDays(3).ToString("dd")
        $DayName4 = $Date1.AddDays(3).ToString("dddd")

        #Fifth Day Selected
        $DateFormat9 = $Date1.AddDays(4).ToString("dd")
        $DayName5 = $Date1.AddDays(4).ToString("dddd")

        #Sixth Day Selected
        $DateFormat10 = $Date1.AddDays(5).ToString("dd")
        $DayName6 = $Date1.AddDays(5).ToString("dddd")

        #Seventh Day Selected
        $DateFormat11 = $Date1.AddDays(6).ToString("dd")
        $DayName7 = $Date1.AddDays(6).ToString("dddd")



        #Selected Year
        $DateFormat18 = $Date1.ToString("yyyy")
        
        Write-Host "Date selected: " -BackgroundColor Black -ForegroundColor Green -NoNewline
        Write-Host "$DayName1 $DateFormat4/$DateFormat5, $DayName2 $dateFormat4/$dateFormat6, $DayName3 $dateFormat4/$dateFormat7, $DayName4 $dateFormat4/$dateFormat8, $DayName5 $dateFormat4/$dateFormat9, $DayName6 $dateFormat4/$dateFormat10, $DayName7 $dateFormat4/$dateFormat11" -BackgroundColor Black -ForegroundColor White 
    }

    #Get the folder name for the client
    $getClientFolder = Get-ChildItem "c:\Path\To\Biometric Audit Reports\$custID*"
    $getClientFolderName = Get-ChildItem "c:\Path\To\Biometric Audit Reports\$custID*" | % {$_.basename}
    
    #Main Folder
    Write-Host "Main Folder location set to: " -BackgroundColor Black -ForegroundColor Green -NoNewline
    Write-Host $getClientFolder -BackgroundColor Black -ForegroundColor White 

    #
    Write-Host "Header Date selected: " -BackgroundColor Black -ForegroundColor Green -NoNewline
    Write-Host "$DateFormat1" -BackgroundColor Black -ForegroundColor White 
    
    Write-Host "Subject Date selected: " -BackgroundColor Black -ForegroundColor Green -NoNewline
    Write-host "$DateFormat2" -BackgroundColor Black -ForegroundColor White
    
    #Creating Archive folder for client if necessary
    Write-Host "Creating Archive folders..." -BackgroundColor Black -ForegroundColor White -NoNewline
    Write-Host "CSVArchive" -BackgroundColor Black -ForegroundColor Magenta -NoNewline
    Write-Host "..." -BackgroundColor Black -ForegroundColor White -NoNewline
    $newClientCSVArchive = New-Item -Path $getClientFolder\CSVArchive\$year\$dateFormat3 -ItemType directory -ErrorAction SilentlyContinue

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
        $outMergedBioReport = "$getClientFolder\CSVArchive\$year\out.csv"        
        $outChangeHeaders = "$getClientFolder\CSVArchive\$year\outHeaders.csv"
        
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


        
        $Users = "Dasan", "Emilie", "Jim", "Mike", "Riley", "Steph", "Steve"
        $Dates = "$dateFormat4/$dateFormat5", "$dateFormat4/$dateFormat6", "$dateFormat4/$dateFormat7", "$dateFormat4/$dateFormat8", "$dateFormat4/$dateFormat9", "$dateFormat4/$dateFormat10", "$dateFormat4/$dateFormat11"
        $daynames = "$DayName1", "$DayName2", "$DayName3", "$DayName4", "$DayName5", "$DayName6", "$DayName7"

        Write-Host "Getting Users" -BackgroundColor Black -ForegroundColor Green -NoNewline
        Write-Host " $Users " -BackgroundColor Black -ForegroundColor White -NoNewline
        Write-Host " and Dates" -BackgroundColor Black -ForegroundColor Green -NoNewline
        Write-Host " $daynames " -BackgroundColor Black -ForegroundColor White 

        foreach ($user in $users)
            {   
                Write-Host "$user " -BackgroundColor Black -ForegroundColor Green
                foreach ($d in $dates)
                    {
                        Write-Host "$d" -BackgroundColor Black -ForegroundColor White
                        Import-Csv $outChangeHeaders | select "Date", "Location", "User" | where {$_.date -match "$d" -and $_.user -match "$user"} | select -Last 1 -First 1 | Export-Csv -Path $outMergedBioReport -Append -NoTypeInformation -Force
                    }
            }

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
                    $excel.Visible = $false

                    #Add workbook
                    $workbook = $excel.workbooks.Add()

                    #Define initial worksheet number
                    $i = 1
                }

            Process 
                {

                    $workbook.worksheets.add() | Out-Null
                    $workbook.worksheets.add() | Out-Null
                    $workbook.worksheets.add() | Out-Null
                    $workbook.worksheets.add() | Out-Null
                    $workbook.worksheets.add() | Out-Null
                    $workbook.worksheets.add() | Out-Null

                    ForEach ($input in $inputfile) 
                        {
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
                            #          
                            $objRange = $Worksheet.UsedRange
                            #$objRange.Interior.ColorIndex = 16
                            $objRange.Font.Size = 10
                            $objRange.Font.Name = "Cambria"
            
                            #Create a Title for the first worksheet
                            $row = 1
                            $Column = 1

                            #Save the initial row so it can be used later to create a border
                            $initalRow = $row

                            # Add the headers to the worksheet
                            $headers = "Date","Location", "Name"
                            $headers | foreach {
                            $worksheet.Cells.Item($row,$column)=$_
                                $worksheet.Cells.Item($row,$column).Font.Bold = $True
                                $worksheet.Cells.Item($row,$column).Font.Size = 12
                                $column++
                            }
                            $i++
                
                            #Sort Headers
                            Write-Host "Sorting headers." -BackgroundColor Black -ForegroundColor White -NoNewline
                            Write-Host " headers." -BackgroundColor Black -ForegroundColor Magenta
                            $setSheetToUsedrange = $worksheet.UsedRange
                            $setSheetToRange = $worksheet.Range("A1")
                            [void]$setSheetToUsedrange.Sort($setSheetToRange, 1,$null,$null,1,$null,1,1)
                            
                            

                            

                            $excel.Worksheets.item(2).name = "$DayName1"
                            $excel.Worksheets.item(3).name = "$DayName2"
                            
                            
                            $excel.Worksheets.item(4).name = "$DayName3"
                            
                            $excel.Worksheets.item(5).name = "$DayName4"
                            $excel.Worksheets.item(6).name = "$DayName5"
                            $excel.Worksheets.item(7).name = "$DayName6"
                            $excel.Worksheets.item(8).name = "$DayName7"
                            $excel.Worksheets.item(9).name = "Export"

                            $row = 2

                            Write-Warning "Starting for Loop."
                            #For loop: For ($i = 1; $i -le $row; $i++)
                            <#
                            Do
                                {
                                    for ($i = 1; $i -le $row; $i++) 
                                        {
                                            foreach ($d in $dates)
                                                {
                                                    Foreach ($dy in $daynames)
                                                        {
                                                            If ($worksheet.cells.Item($i,1).Value() -match "$d")
                                                                {
                                                                    $cellValue = $worksheet.cells.Item($i,1).Value()
                                        
                                                                    Write-Host "Cell " -BackgroundColor Black -ForegroundColor White -NoNewline
                                                                    Write-Host "A$i's " -BackgroundColor Black -ForegroundColor Green
                                                                    
                                                                    Write-Host "Value is:" -BackgroundColor Black -ForegroundColor Green -NoNewline
                                                                    Write-Host " $cellValue " -BackgroundColor Black -ForegroundColor White -NoNewline 

                                                                    Write-Host "Date checked: " -BackgroundColor Black -ForegroundColor Green -NoNewline
                                                                    Write-Host "$d" -BackgroundColor Black -ForegroundColor White

                                                                    $selectRow = $worksheet.cells.Item($i,1).EntireRow.Copy()
                                        
                                                                    $sheetToDest = $workbook.sheets.item("$dy")
                                                                    $sheetToDest.Range("A$i").PasteSpecial() | Out-Null
                                                                    $sheetToDest.Cells.Item($i,1).NumberFormat = "hh:mm:ss"

                                                                    $Row++
                                                                }
                                                            Else
                                                                {
                                                                    $cellValue = $worksheet.cells.Item($i,1).Value()

                                                                    Write-Host "Cell A$i's value: $cellValue does not match $d" -BackgroundColor Black -ForegroundColor Red
                                                                    $row++
                                                                }
                                                        }

                                                }
                                        }

                                }
                            while ([string]::IsNullOrWhiteSpace($worksheet.cells.Item($i,1).Value2))
                            #>  
                            
                            #For loop: For ($i = 1; $i -le $row; $i++)
                            
                            Do
                                {
                                    for ($i = 1; $i -le $row; $i++) 
                                        {
                                            If ($worksheet.cells.Item($i,1).Value() -match "$d")
                                                {
                                                    $cellValue = $worksheet.cells.Item($i,1).Value()
                                        
                                                    Write-Host "Cell " -BackgroundColor Black -ForegroundColor White -NoNewline
                                                    Write-Host "A$i's " -BackgroundColor Black -ForegroundColor Green
                                                                    
                                                    Write-Host "Value is:" -BackgroundColor Black -ForegroundColor Green -NoNewline
                                                    Write-Host " $cellValue " -BackgroundColor Black -ForegroundColor White -NoNewline 

                                                    Write-Host "Date checked: " -BackgroundColor Black -ForegroundColor Green -NoNewline
                                                    Write-Host "$d" -BackgroundColor Black -ForegroundColor White

                                                    $selectRow = $worksheet.cells.Item($i,1).EntireRow.Copy()
                                        
                                                    $sheetToDest = $workbook.sheets.item("$dy")
                                                    $sheetToDest.Range("A$i").PasteSpecial() | Out-Null
                                                    $sheetToDest.Cells.Item($i,1).NumberFormat = "hh:mm:ss"

                                                    $Row++
                                                }
                                            Else
                                                {
                                                    $cellValue = $worksheet.cells.Item($i,1).Value()

                                                    Write-Host "Cell A$i's value: $cellValue does not match $d" -BackgroundColor Black -ForegroundColor Red
                                                    $row++
                                                }
                                        }
                                }
                            while ([string]::IsNullOrWhiteSpace($worksheet.cells.Item($i,1).Value2))
                             
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

        Get-ChildItem $outMergedBioReport | ConvertCSV-ToExcel -output $getClientFolder\ExcelArchive\$year\Output.xlsx
        
        $tpOutFiles1 = Test-path $outMergedBioReport
        If($tpOutFiles1 -eq $true)
            { 
                Remove-Item $outMergedBioReport
            }
        else
            {
                Write-Warning "$outMergedBioReport does not exist"
            }
        $tpOutFiles2 = Test-Path $outChangeHeaders
        if($tpOutFiles2 -eq $True)
            {
                Remove-Item $outChangeHeaders
            }
        Else
            {
                Write-Warning "$outChangeHeaders does not exist."
            }
        ii $getClientFolder\ExcelArchive\$year\output.xlsx
    }
    Else
    {
        Write-Warning "Folder not successfully created."
    }
}
Else
{
    Write-Warning "Exports folder does not contain CSV files."
}



#Get Script speed
$date2 = Get-Date
$Time = New-TimeSpan -Start $date -End $date2
Write-Host "This took" -BackgroundColor Black -ForegroundColor Green -NoNewline 
Write-Host " $time " -BackgroundColor Black -ForegroundColor Magenta -NoNewline
Write-host "to run." -BackgroundColor Black -ForegroundColor Green
