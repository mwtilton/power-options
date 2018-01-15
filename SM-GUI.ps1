Add-Type -AssemblyName System.Windows.Forms

$Form = New-Object system.Windows.Forms.Form
$Form.Text = "Shelly For the NOC"
$Form.BackColor = "#111111"
$Form.TopMost = $true
$Form.Width = 721
$Form.Height = 391

$label2 = New-Object system.windows.Forms.Label
$label2.Text = "PassDown"
$label2.AutoSize = $true
$label2.ForeColor = "#ffffff"
$label2.Width = 25
$label2.Height = 10
$label2.location = new-object system.drawing.point(28,15)
$label2.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label2)

$label3 = New-Object system.windows.Forms.Label
$label3.Text = "PowerAuditReports"
$label3.AutoSize = $true
$label3.ForeColor = "#ffffff"
$label3.Width = 25
$label3.Height = 10
$label3.location = new-object system.drawing.point(254,26)
$label3.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label3)

$label4 = New-Object system.windows.Forms.Label
$label4.Text = "Biometrics"
$label4.AutoSize = $true
$label4.ForeColor = "#ffffff"
$label4.Width = 25
$label4.Height = 10
$label4.location = new-object system.drawing.point(549,24)
$label4.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label4)

$label5 = New-Object system.windows.Forms.Label
$label5.Text = "In"
$label5.AutoSize = $true
$label5.ForeColor = "#ffffff"
$label5.Width = 25
$label5.Height = 10
$label5.location = new-object system.drawing.point(15,48)
$label5.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label5)

$label6 = New-Object system.windows.Forms.Label
$label6.Text = "Out"
$label6.AutoSize = $true
$label6.ForeColor = "#ffffff"
$label6.Width = 25
$label6.Height = 10
$label6.location = new-object system.drawing.point(23,90)
$label6.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label6)

$label7 = New-Object system.windows.Forms.Label
$label7.Text = "Time:"
$label7.AutoSize = $true
$label7.ForeColor = "#ffffff"
$label7.Width = 25
$label7.Height = 10
$label7.location = new-object system.drawing.point(15,137)
$label7.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label7)

$textBox8 = New-Object system.windows.Forms.TextBox
$textBox8.Width = 100
$textBox8.Height = 20
$textBox8.location = new-object system.drawing.point(62,53)
$textBox8.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($textBox8)

$textBox9 = New-Object system.windows.Forms.TextBox
$textBox9.Width = 100
$textBox9.Height = 20
$textBox9.location = new-object system.drawing.point(62,91)
$textBox9.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($textBox9)

$textBox10 = New-Object system.windows.Forms.TextBox
$textBox10.Width = 100
$textBox10.Height = 20
$textBox10.location = new-object system.drawing.point(61,137)
$textBox10.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($textBox10)

$label11 = New-Object system.windows.Forms.Label
$label11.Text = "Client ID"
$label11.AutoSize = $true
$label11.ForeColor = "#ffffff"
$label11.Width = 25
$label11.Height = 10
$label11.location = new-object system.drawing.point(462,57)
$label11.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label11)

$textBox12 = New-Object system.windows.Forms.TextBox
$textBox12.Width = 100
$textBox12.Height = 20
$textBox12.location = new-object system.drawing.point(543,54)
$textBox12.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($textBox12)

$button14 = New-Object system.windows.Forms.Button
$button14.Text = "button"
$button14.Width = 60
$button14.Height = 30
$button14.location = new-object system.drawing.point(44,184)
$button14.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($button14)

$button15 = New-Object system.windows.Forms.Button
$button15.Text = "button"
$button15.Width = 60
$button15.Height = 30
$button15.location = new-object system.drawing.point(284,184)
$button15.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($button15)

$button16 = New-Object system.windows.Forms.Button
$button16.Text = "button"
$button16.Width = 60
$button16.Height = 30
$button16.location = new-object system.drawing.point(548,184)
$button16.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($button16)

$button17 = New-Object system.windows.Forms.Button
$button17.BackColor = "#4d4d4d"
$button17.Text = "button"
$button17.Width = 60
$button17.Height = 30
$button17.location = new-object system.drawing.point(629,309)
$button17.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($button17)

[void]$Form.ShowDialog()
$Form.Dispose()