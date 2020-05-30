$PSDefaultParameterValues=@{'Write-host:BackGroundColor'='Black';'Write-host:ForeGroundColor'='White'}
#requires -Version 2.0

Import-Module ActiveDirectory -Force
Import-Module Posh-Git -Force
Import-module "$env:ONEDRIVE\Scripts\UCSD\UCSD.psm1" -force

$psversiontable.PSVersion

#$TotalRAM = (systeminfo | Select-String 'Total Physical Memory:').ToString().Split(':')[1].Trim()
$TotalRAM = ((Get-WmiObject Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).sum/1024/1024/1024)

Function Prompt {
    Write-Host "[" -ForegroundColor DarkGray -NoNewline
    Write-Host $(Get-date).ToString("HH:mm:ss") -ForegroundColor Magenta -nonewline
    Write-Host "] " -ForegroundColor DarkGray -nonewline

    $procAVG = $(get-process | select @{name="CPU(s)";Expression="CPU"} | Measure-Object -property "CPU(s)" -Sum | select sum).sum
    $procAVGFormat = [math]::Round($procAVG, 2)
    Write-Host "RAM: " -ForegroundColor yellow -nonewline
    Write-Host "$procAVGFormat" -ForegroundColor Gray -nonewline
    Write-Host "/" -ForegroundColor Magenta -nonewline
    Write-Host "$TotalRAM " -ForegroundColor DarkGray -nonewline

    $user = [Security.Principal.WindowsIdentity]::GetCurrent()
    if ( (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
        $adminfg = "Red"
    }
    else {
        $adminfg = $host.ui.rawui.ForegroundColor
    }

    Switch ((get-location).provider.name) {
        "FileSystem" { $fg = "green"}
        "Registry" { $fg = "magenta"}
        "wsman" { $fg = "cyan"}
        "Environment" { $fg = "yellow"}
        "Certificate" { $fg = "darkcyan"}
        "Function" { $fg = "gray"}
        "alias" { $fg = "darkgray"}
        "variable" { $fg = "darkgreen"}
        Default { $fg = $host.ui.rawui.ForegroundColor}
    }

    #Write-Host "[$((Get-Date).timeofday.tostring().substring(0,8))] " -NoNewline
    Write-Host "PS " -nonewline -ForegroundColor $adminfg
    #$GitPromptSettings.DefaultPromptPrefix = '[$(hostname)] '
    #$GitPromptSettings.DefaultPromptPath = "$env:ONEDRIVE\Scripts\UCSD"
    $GitPromptSettings.DefaultPromptAbbreviateHomeDirectory = $true
    #$GitPromptSettings.DefaultPromptPath.ForegroundColor = 'Orange'
    #$GitPromptSettings.DefaultPromptPath.ForegroundColor = "orange"#$fg
    #$GitPromptSettings.DefaultPromptBeforeSuffix.ForegroundColor = $fg
    $GitPromptSettings.DefaultPromptSuffix = ' $((Get-History -Count 1).id + 1)$(" >" * ($nestedPromptLevel + 1)) '
    #$GitPromptSettings.DefaultPromptWriteStatusFirst = $true
    $prompt = & $GitPromptScriptBlock
    "$prompt"
    #Write-Output "' $((Get-History -Count 1).id + 1)$(">" * ($nestedPromptLevel + 1)) "
}

Set-Location "$env:ONEDRIVE\Scripts\UCSD"


Register-EngineEvent PowerShell.Exiting -Action {

Write-Host "  [>] " -NoNewline
Write-Host " Finishing post processing"

Write-Host "    [+] " -NoNewline
Write-Host " Removing pssessions"

Get-PSSession | Remove-PSSession

sleep 2
}