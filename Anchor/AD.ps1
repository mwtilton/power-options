[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "New User - Data Entry Form"
$objForm.Size = New-Object System.Drawing.Size(400,300) 
$objForm.StartPosition = "CenterScreen"

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {$x=$objTextBox.Text;$objTextBox2.Text;$objForm.Close()}})
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(40,200)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.Add_Click({$x=$objTextBox.Text;$objForm.Close()})
$objForm.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(120,200)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.Add_Click({$objForm.Close()})
$objForm.Controls.Add($CancelButton)

$objLabel1 = New-Object System.Windows.Forms.Label
$objLabel1.Location = New-Object System.Drawing.Size(10,20) 
$objLabel1.Size = New-Object System.Drawing.Size(75,20) 
$objLabel1.Text = "First Name:"
$objForm.Controls.Add($objLabel1) 

$objLabel2 = New-Object System.Windows.Forms.Label
$objLabel2.Location = New-Object System.Drawing.Size(10,100) 
$objLabel2.Size = New-Object System.Drawing.Size(150,20) 
$objLabel2.Text = "Last Name:"
$objForm.Controls.Add($objLabel2) 


$objTextBox = New-Object System.Windows.Forms.TextBox 
$objTextBox.Location = New-Object System.Drawing.Size(10,40) 
$objTextBox.Size = New-Object System.Drawing.Size(260,20) 
$objForm.Controls.Add($objTextBox) 

$objTextBox2 = New-Object System.Windows.Forms.TextBox 
$objTextBox2.Location = New-Object System.Drawing.Size(10,120) 
$objTextBox2.Size = New-Object System.Drawing.Size(200,20) 
$objForm.Controls.Add($objTextBox2) 

$objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()

$x
Write-Host "$X" -ForegroundColor Green