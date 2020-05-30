param (
    [Parameter(
        ValueFromPipeline = $true,
        Position = 0,
        Mandatory = $true,
        HelpMessage = "Set the "
    )]
    [ValidateNotNullOrEmpty()]
    [string]$input,
    [Parameter(
        ValueFromPipeline = $true,
        Position = 0,
        Mandatory = $true,
        HelpMessage = "Set the "
    )]
    [ValidateNotNullOrEmpty()]
    [string]$interest,
    [Parameter(
        ValueFromPipeline = $true,
        Position = 0,
        Mandatory = $true,
        HelpMessage = "Set the "
    )]
    [ValidateNotNullOrEmpty()]
    [string]$length,
    [Parameter(
        ValueFromPipeline = $true,
        Position = 0,
        Mandatory = $true,
        HelpMessage = "Set the "
    )]
    [ValidateNotNullOrEmpty()]
    [string]$name,
    [Parameter(
        ValueFromPipeline = $true,
        Position = 0,
        Mandatory = $true,
        HelpMessage = "Set the "
    )]
    [ValidateNotNullOrEmpty()]
    [string]$addition,
    

)

function get-compound {
	var accumulated = input
	for ( i=0; i -le $years; i++ ) {
		accumulated *= interest
		if ( addition ){
			accumulated += input
		}
	}
	console.log( name + ' will grow from  ' + input + ' to ' + accumulated + ' at ' + interest +  ' over ' + length + ' years' )
}

get-compound  500, 1.05, 40, 'cigarettes', true ) // How much money will a habit of cigarettes generate in retirement money?
