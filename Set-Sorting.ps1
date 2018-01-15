$CSVFileLocation = "C:\Temp"

#Process will error out if this folder is not created first or the folder does not exist
$ARCHIVE = "$csvfilelocation\Modified"
New-Item -Path $ARCHIVE -ItemType directory -ErrorAction SilentlyContinue

$GetBaseName = Get-ChildItem $CSVFileLocation -Filter "*.csv" | Where-Object {!$_.PSIsContainer} | ForEach-Object {$_.BaseName}

foreach($file in $GetBaseName)
{
    Write-Host $file
    Import-Csv $CSVFileLocation\$file.csv -Delimiter ";" -Encoding Default | Select-Object "Anst?llningsnummer","Fullst?ndigt namn","Avdelning","Alternativ e-postadress","Administrativ roll","Stad	" | Sort-Object "Fullst?ndigt namn","Avdelning" | Export-Csv -Path $ARCHIVE\$file.csv -Delimiter "," -NoTypeInformation -Force 
    
}