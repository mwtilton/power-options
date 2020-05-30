$csv = "C:\Temp\users.csv"   


Import-Csv $csv -Delimiter "," -Encoding UTF8 | Select-Object "name", "in", "out"  | Group-Object "name" -NoElement | Export-Csv C:\Temp\users-mod.csv

ii C:\Temp\users-mod.csv