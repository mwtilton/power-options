# Created By:
# "Ta11ow"
#
# Reddit Link: 
iexplore "https://www.reddit.com/r/PowerShell/comments/7o56v8/get_hard_disk_information_wanimated_graph/ds72aob/"
# Run this script to see connected hard disk status with nifty bar graph!
Begin { 
    # Create bar graph
    function Get-BarGraph {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory, ValueFromPipeline)]
            [int]$Percent,

            [int]$Length = 30
        )

        # Tracks how many ticks to draw
        $Ticks = $Percent / 100 * $Length - 1
        $BarString = "[" + ("$([char]0x25A0)" * $Ticks) + (" " * ($Length - $Ticks)) + "]"
        Write-Output $BarString
    }

    $myDisks = Get-WmiObject -class Win32_LogicalDisk
}
Process {
    $myDisks | ForEach-Object {
        if ($_.Size -gt 0) {
            $TotalSpace = [Math]::Round($_.Size / 1gb)
            $UsedSpace = [Math]::Round(($_.Size - $_.FreeSpace) / 1gb)
            $FreeSpace = [Math]::Round($_.FreeSpace / 1gb)
            $PercentUsed = [Math]::Round((($_.Size - $_.FreeSpace) / $_.Size) * 100)
            [PsCustomObject]@{
                Volume      = $_.VolumeName
                DriveLetter = $_.DeviceID
                TotalSpace  = "$TotalSpace GB"
                UsedSpace   = "$UsedSpace GB"
                FreeSpace   = "$FreeSpace GB"
                PercentUsed = "$PercentUsed %"
                Graph       = ($PercentUsed | Get-BarGraph)
            }
        }
    } | Out-GridView
}