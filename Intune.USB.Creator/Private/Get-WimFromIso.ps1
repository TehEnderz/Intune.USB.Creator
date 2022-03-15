function Get-WimFromIso {
    [cmdletbinding()]
    param (
        [parameter(Mandatory = $true)]
        [string]$isoPath,

        [parameter(Mandatory = $true)]
        [string]$wimDestination
    )
    try {
        $mount = Mount-DiskImage -ImagePath $isoPath -PassThru
        if ($mount) {
            $volume = Get-DiskImage -ImagePath $mount.ImagePath | Get-Volume
            if (!(Test-Path $wimDestination -ErrorAction SilentlyContinue)) {
                New-Item -Path $wimDestination -ItemType Directory -Force | Out-Null
            }

            $wimSource = "$($volume.DriveLetter)`:\sources"

            if (Test-Path "$wimSource\install.esd") {
                throw "You are trying to use a consumer ISO, which is not supported."
            }

            Invoke-FileTransfer -source "$wimSource\install.wim"  -destination "$wimDestination\install.wim"
        }
    }
    catch {
        Write-Warning $_
    }
    finally {
        Dismount-DiskImage -ImagePath $isoPath | Out-Null
        Write-Host $([char]0x221a) -ForegroundColor Green
    }
}