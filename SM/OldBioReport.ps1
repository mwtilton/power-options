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
