function get-mac {

param(
    $str
)

$colitems = Get-WmiObject -Class win32_networkadapterconfiguration -computername $str | where {$_.ipenabled -match $true}

    foreach($obj in $colitems){
        $obj | select Description, MACaddress
    }

}
get-mac localhost