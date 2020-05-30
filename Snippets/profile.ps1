Get-ExecutionPolicy -List

$console = $host.ui.rawui
$console.BackgroundColor = "Black"
$console.ForegroundColor = "White"

<#
ForegroundColor       : Gray
BackgroundColor       : Black
CursorPosition        : 0,62
WindowPosition        : 0,42
CursorSize            : 25
BufferSize            : 94,3000
WindowSize            : 94,21
MaxWindowSize         : 94,171
MaxPhysicalWindowSize : 404,171
KeyAvailable          : True
WindowTitle           :


$console = $host.ui.rawui
$console.backgroundcolor = "black"
$console.foregroundcolor = "white"
$colors = $host.privatedata
$colors.verbosebackgroundcolor = "Magenta"
$colors.verboseforegroundcolor = "Green"
$colors.warningbackgroundcolor = "Red"
$colors.warningforegroundcolor = "white"
$colors.ErrorBackgroundColor = "DarkCyan"
$colors.ErrorForegroundColor = "Yellow"
set-location C:\
clear-host
#>