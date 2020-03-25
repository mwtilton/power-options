class CyberNinja
{
    # Properties
    [String] $Alias
    [int32] $HitPoints

    # Static Properties
    static [String] $Clan = "DevOps Library"

    # Hidden Properties
    hidden [String] $RealName

    # Parameterless Constructor
    CyberNinja ()
    {
    }

    # Constructor
    CyberNinja ([String] $Alias, [int32] $HitPoints)
    {
        $this.Alias = $Alias
        $this.HitPoints = $HitPoints
    }

    # Method
    [String] getAlias()
    {
       return $this.Alias
    }

    # Static Method
    static [String] getClan()
    {
        return [CyberNinja]::Clan
    }

    # ToString Method
    [String] ToString()
    {
        return $this.Alias + ":" + $this.HitPoints
    }
}

$ninja = New-Object CyberNinja

$ninja.Alias = "<xml>"
$ninja.HitPoints = "100"

$ninja.tostring()