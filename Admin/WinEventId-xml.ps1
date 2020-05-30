# Grab the events from a DC            
$Events = Get-WinEvent -FilterHashtable @{Logname='Application';Id=16384}            
            
# Parse out the event message data            
ForEach ($Event in $Events) {            
    # Convert the event to XML            
    $eventXML = [xml]$Event.ToXml()            
    # Iterate through each one of the XML message properties            
    For ($i=0; $i -lt $eventXML.Event.EventData.Data.Count; $i++) {            
        # Append these as object properties            
        Add-Member -InputObject $Event -MemberType NoteProperty -Force -Name $eventXML.Event.EventData.Data[$i].name -Value $eventXML.Event.EventData.Data[$i].'#text'            
    }            
}            
            
# View the results with your favorite output method            
$Events | Export-Csv .\events.csv            
$Events | Select-Object * | Out-GridView 