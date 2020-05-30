$PSDefaultParameterValues=@{'Write-host:BackGroundColor'='Black';'Write-host:ForeGroundColor'='White'}

# Powershell Version
Write-Host "Powershell Version: "$psversiontable.PSVersion.major"."$psversiontable.PSVersion.Minor"."$psversiontable.PSVersion.Build

# Operating System Version
Write-Host "Operating System Version Information: "
Write-Host $([System.Environment]::OSVersion.Version.major)"."$([System.Environment]::OSVersion.Version.Minor)"."$([System.Environment]::OSVersion.Version.Build)
(Get-WmiObject -class Win32_OperatingSystem).Caption
(Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ReleaseId).ReleaseId

Get-EventSubscriber -Force | Unregister-Event -Force

$TotalRAM = ((Get-WmiObject Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).sum/1024/1024)

Function Prompt {


    $user = [Security.Principal.WindowsIdentity]::GetCurrent()
    $role = (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    if ($role) {
        $adminfg = "Red"
    }
    else {
        $adminfg = $host.ui.rawui.ForegroundColor
    }

    Switch ((get-location).provider.name) {
        "FileSystem" { $fg = "Green"}
        "Registry" { $fg = "magenta"}
        "wsman" { $fg = "cyan"}
        "Environment" { $fg = "yellow"}
        "Certificate" { $fg = "darkcyan"}
        "Function" { $fg = "gray"}
        "alias" { $fg = "darkgray"}
        "variable" { $fg = "darkgreen"}
        Default { $fg = $host.ui.rawui.ForegroundColor}
    }
    #Admin
    Write-Host "[$Env:username][ADMIN]: PS " -nonewline -ForegroundColor $adminfg

    #time
    Write-Host "[" -ForegroundColor DarkGray -NoNewline
    Write-Host $(Get-date).ToString("HH:mm:ss") -ForegroundColor Magenta -nonewline
    Write-Host "] " -ForegroundColor DarkGray -nonewline

    #RAM
    $procAVG = $(get-process | select @{name="WorkingSet";Expression="WorkingSet"} | Measure-Object -property "WorkingSet" -Sum | select sum).sum/1024/1024
    $procAVGFormat = [math]::Round($procAVG, 2)
    Write-Host "RAM: " -ForegroundColor yellow -nonewline
    Write-Host "$procAVGFormat" -ForegroundColor Gray -nonewline
    Write-Host "/" -ForegroundColor Magenta -nonewline
    Write-Host "$TotalRAM " -ForegroundColor DarkGray -nonewline
    $p = Split-Path -leaf -path (Get-Location)
    Write-Host "\$($p) " -nonewline -ForegroundColor Green
}

Register-EngineEvent PowerShell.Exiting -Action {
    
    Clear-SavedHistory
    sleep 2

} | Out-Null