#region Functions
function Set-PowerPolicy {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory = $false)]
        [ValidateSet('PowerSaver', 'Balanced', 'HighPerformance')]
        [string]$powerPlan
    )
    try {
        switch ($powerPlan) {
            PowerSaver {
                Write-Host "Setting power policy to 'Power Saver'.." -ForegroundColor Cyan
                $planGuid = "a1841308-3541-4fab-bc81-f71556f20b4a"
            }
            Balanced {
                Write-Host "Setting power policy to 'Balanced Performance'.." -ForegroundColor Cyan
                $planGuid = "381b4222-f694-41f0-9685-ff5bb260df2e"
            }
            HighPerformance {
                Write-Host "Setting power policy to 'High Performance'.." -ForegroundColor Cyan
                $planGuid = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
            }
            default {
                throw "Incorrect selection.."
            }
        }
        Invoke-CmdLine -application powercfg -argumentList "/s $planGuid" -silent
    }
    catch {
        throw $_
    }
}
function Test-IsUEFI {
    try {
        $pft = Get-ItemPropertyValue -Path HKLM:\SYSTEM\CurrentControlSet\Control -Name 'PEFirmwareType'
        switch ($pft) {
            1 {
                Write-Host "BIOS Mode detected.." -ForegroundColor Cyan
                return "BIOS"
            }
            2 {
                Write-Host "UEFI Mode detected.." -ForegroundColor Cyan
                return "UEFI"
            }
            Default {
                Write-Host "BIOS / UEFI undetected.." -ForegroundColor Red
                return $false
            }
        }
    }
    catch {
        throw $_
    }
}
function Invoke-Cmdline {
    [cmdletbinding()]
    param (
        [parameter(mandatory = $true)]
        [string]$application,

        [parameter(mandatory = $true)]
        [string]$argumentList,

        [parameter(Mandatory = $false)]
        [switch]$silent
    )
    if ($silent) {
        cmd /c "$application $argumentList > nul 2>&1"
    } else {
        cmd /c "$application $argumentList"
    }

    if ($LASTEXITCODE -ne 0) {
        throw "An error has occurred.."
    }
}
function Get-DiskPartVolume {
    [cmdletbinding()]
    param (
        [parameter(mandatory = $false)]
        [string]$winPEDrive = "X:"
    )
    try {
        #region Map drive letter for install.wim
        $lvTxt = "$winPEDrive\listvol.txt"
        $lv = @"
List volume
exit
"@
        $lv | Out-File $lvTxt -Encoding ascii -Force -NoNewline
        $dpOutput = Invoke-CmdLine -application "diskpart" -argumentList "/s $lvTxt"
        $dpOutput = $dpOutput[6..($dpOutput.length - 3)]
        $vals = $dpOutput[2..($dpOutput.Length - 1)]
        $res = foreach ($val in $vals) {
            $dr = $val.Substring(10, 6).Replace(" ", "")
            [PSCustomObject]@{
                VolumeNum  = $val.Substring(0, 10).Replace(" ", "")
                DriveRoot  = if ($dr -ne "") { "$dr`:\" } else { $null }
                Label      = $val.Substring(17, 13).Replace(" ", "")
                FileSystem = $val.Substring(30, 7).Replace(" ", "")
                Type       = $val.Substring(37, 12).Replace(" ", "")
                Size       = $val.Substring(49, 9).Replace(" ", "")
                Status     = $val.Substring(58, 11).Replace(" ", "")
                Info       = $val.Substring($val.length - 10, 10).Replace(" ", "")
            }
        }
        return $res
        #endregion
    }
    catch {
        throw $_
    }
}
function Get-SystemDeviceId {
    try {
       $dataDrives = $drives | ?{ $_.BusType -ne "USB"}
       if (@($DataDrives).count -eq 1) {
            $targetDrive = $DataDrives[0].DeviceId
            return $targetDrive
       }
       elseif (@($DataDrives).count -gt 1) {
            Write-Host "More than one disk has been detected. Select disk where Windows should be installed" -ForegroundColor Yellow
            $DataDrives | ft DeviceId, FriendlyName, Size| Out-String | % {Write-Host $_ -ForegroundColor Cyan}
            $targetDrive = Read-Host "Please make a selection..."
            return $targetDrive
        }
       else {
           throw "Error while getting DeviceId of potiential Windows target drives"
       }
    }
    catch {
        throw $_
    }
}

function Set-DrivePartition {
    [cmdletbinding()]
    param (
        [parameter(mandatory = $false)]
        [string]$winPEDrive = "X:",

        [parameter(mandatory = $false)]
        [string]$targetDrive = "0"
    )
    try {
        $txt = "$winPEDrive\winpart.txt"
        New-Item $txt -ItemType File -Force | Out-Null
        Write-Host "Checking boot system type.." -ForegroundColor Cyan
        $bootType = Test-IsUEFI
        #region Boot type switch
        switch ($bootType) {
            "BIOS" {
                $winpartCmd = @"
select disk $targetDrive
clean
create partition primary size=100
active
format quick fs=fat32 label="System"
assign letter="S"
create partition primary
format quick fs=ntfs label="Windows"
assign letter="W"
shrink desired=450
create partition primary
format quick fs=ntfs label="Recovery"
assign letter="R"
exit
"@
            }
            "UEFI" {
                $winpartCmd = @"
select disk $targetDrive
clean
convert gpt
create partition efi size=100
format quick fs=fat32 label="System"
assign letter="S"
create partition msr size=16
create partition primary
format quick fs=ntfs label="Windows"
assign letter="W"
shrink desired=950
create partition primary
format quick fs=ntfs label="Recovery"
assign letter="R"
set id="de94bba4-06d1-4d40-a16a-bfd50179d6ac"
gpt attributes=0x8000000000000001
exit
"@
            }
            default {
                throw "Boot type could not be detected.."
            }
        }
        #endregion
        #region Partition disk
        $winpartCmd | Out-File $txt -Encoding ascii -Force -NoNewline
        Write-Host "Setting up partition table.." -ForegroundColor Cyan
        Invoke-Cmdline -application diskpart -argumentList "/s $txt" -silent
        #endregion

    }
    catch {
        throw $_
    }

}
function Find-InstallWim {
    [cmdletbinding()]
    param (
        [parameter(Mandatory = $true)]
        $volumeInfo
    )
    try {
        foreach ($vol in $volumeInfo) {
            if ($vol.DriveRoot) {
                if (Test-Path "$($vol.DriveRoot)images\install.wim" -ErrorAction SilentlyContinue) {
                    Write-Host "Install.wim found on drive: $($vol.DriveRoot)" -ForegroundColor Cyan
                    $res = $vol
                }
            }
        }
    }
    catch {
        Write-Warning $_.Exception.Message
    }
    finally {
        if (!($res)) {
            Write-Warning "Install.wim not found on any drives.."
        }
        else {
            $res
        }
    }
}
#endregion


function printKeyValue([String] $Key, [String] $Value) {
    Write-Host -ForegroundColor Yellow -NoNewline $key
    Write-Host -ForegroundColor Cyan $value
}

class ImageDeploy {
    [string]                    $winPEDrive = $null
    [string]                    $winPESource = $env:winPESource
    [PSCustomObject]            $volumeInfo = $null
    [string]                    $installPath = $null
    [string]                    $installRoot = $null
    [System.IO.DirectoryInfo]   $scratch = $null
    [string]                    $scRoot = $null
    [System.IO.DirectoryInfo]   $recovery = $null
    [string]                    $reRoot = $null
    [System.IO.DirectoryInfo]   $driverPath = $null

    ImageDeploy ([string]$winPEDrive) {
        $this.winPEDrive = $winPEDrive
        $this.volumeInfo = Get-DiskPartVolume -winPEDrive $winPEDrive
        $this.installRoot = (Find-InstallWim -volumeInfo $($this.volumeInfo)).DriveRoot
        $this.installPath = "$($this.installRoot)images"

        $this.model = Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty Model
        $this.serial = Get-CimInstance -ClassName Win32_Bios | Select-Object -ExpandProperty SerialNumber
        $this.driverPath = "$($this.installRoot)Drivers\$($this.deviceModel)"
    }

    loadDriversInWinPE() {
        if (-Not Test-Path "$($this.driverPath)\WinPE") {
            Write-Host "No WinPE drivers detected.." -ForegroundColor Yellow
            return
        }

        $drivers = Get-ChildItem "${usb.driverPath}\WinPE" -Filter *.inf -Recurse
        foreach ($d in $drivers) {
            printKeyValue("Loading driver: ", $d.fullName)
            . drvload $d.fullName
        }
    }

    prompt() {
        Clear-Host

        $welcomeScreen = "IF9fICBfXyAgICBfXyAgX19fX19fICBfX19fX18gIF9fX19fXwovXCBcL1wgIi0uLyAgXC9cICBfXyBcL1wgIF9fX1wvXCAgX19fXApcIFwgXCBcIFwtLi9cIFwgXCAgX18gXCBcIFxfXyBcIFwgIF9fXAogXCBcX1wgXF9cIFwgXF9cIFxfXCBcX1wgXF9fX19fXCBcX19fX19cCiAgXC9fL1wvXy8gIFwvXy9cL18vXC9fL1wvX19fX18vXC9fX19fXy8KIF9fX19fICAgX19fX19fICBfX19fX18gIF9fICAgICAgX19fX19fICBfXyAgX18KL1wgIF9fLS4vXCAgX19fXC9cICA9PSBcL1wgXCAgICAvXCAgX18gXC9cIFxfXCBcClwgXCBcL1wgXCBcICBfX1xcIFwgIF8tL1wgXCBcX19fXCBcIFwvXCBcIFxfX19fIFwKIFwgXF9fX18tXCBcX19fX19cIFxfXCAgIFwgXF9fX19fXCBcX19fX19cL1xfX19fX1wKICBcL19fX18vIFwvX19fX18vXC9fLyAgICBcL19fX19fL1wvX19fX18vXC9fX19fXy8KICAgICAgIF9fX19fX19fX19fX19fX19fX19fX19fX19fX19fX19fX19fCiAgICAgICBXaW5kb3dzIDEwIERldmljZSBQcm92aXNpb25pbmcgVG9vbAogICAgICAgKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKioqKio="
        Write-Host $([system.text.encoding]::UTF8.GetString([system.convert]::FromBase64String($welcomeScreen)))
        Write-Host "=================================================================="
        printKeyValue("Model: ", $this.deviceModel)
        printKeyValue("Serial: ", $this.serial)
        Write-Host "===================== Press ENTER to install ====================="

        Read-Host
    }

    prepareDisk() {
        Write-Host "Configuring drive partitions.." -ForegroundColor Yellow
        $targetDrive = Get-SystemDeviceId
        Set-DrivePartition -winPEDrive $usb.winPEDrive -targetDrive $targetDrive

        Write-Host "Setting up Scratch & Recovery paths.." -ForegroundColor Yellow

        $this.scratch = "W:\recycler\scratch"
        $this.scRoot = $this.scratch.Root
        $this.recovery = "R:\RECOVERY\WINDOWSRE"
        $this.reRoot = $this.recovery.Root

        New-Item -Path $this.scratch.FullName -ItemType Directory -Force | Out-Null
        New-Item -Path $this.recovery.FullName -ItemType Directory -Force | Out-Null
    }

    deployImage() {
        Write-Host "Applying the windows image from the USB.." -ForegroundColor Yellow
        $imageIndex = Get-Content "$($this.installPath)\imageIndex.json" -Raw | ConvertFrom-Json -Depth 20
        Invoke-Cmdline -application "DISM" -argumentList "/Apply-Image /ImageFile:$($this.installPath)\install.wim /Index:$($imageIndex.imageIndex) /ApplyDir:$($this.scRoot) /EA /ScratchDir:$($this.scratch)"
    }

    postDeploy() {
        $this.injectAutopilotConfig()
        $this.setupWindowsRecovery()
        $this.makeDiskBootable()
        $this.applyUnattend()
        $this.injectProvisioningPackages()
        $this.injectOSDrivers()
    }

    injectAutopilotConfig() {
        if (-Not Test-Path "$PSScriptRoot\AutopilotConfigurationFile.json" -ErrorAction SilentlyContinue) {
            return
        }

        if (Test-Path "$($this.scRoot)Windows\Provisioning\Autopilot") {
            return
        }

        Write-Host "`nInjecting AutoPilot configuration file.." -ForegroundColor Yellow
        Copy-Item "$PSScriptRoot\AutopilotConfigurationFile.json" -Destination "$($usb.scRoot)Windows\Provisioning\Autopilot\AutopilotConfigurationFile.json" -Force | Out-Null
    }

    setupWindowsRecovery() {
        Write-Host "Move WinRE to recovery partition.." -ForegroundColor Yellow

        $reWimPath = "$($this.scRoot)Windows\System32\recovery\winre.wim"
        if (-Not Test-Path $reWimPath -ErrorAction SilentlyContinue) {
            return
        }

        Write-Host "Moving the recovery wim into place.." -ForegroundColor Yellow
        (Get-ChildItem -Path $reWimPath -Force).attributes = "NotContentIndexed"
        Move-Item -Path $reWimPath -Destination "$($this.recovery.FullName)\winre.wim"
        (Get-ChildItem -Path "$($this.recovery.FullName)\winre.wim" -Force).attributes = "ReadOnly", "Hidden", "System", "Archive", "NotContentIndexed"
        Write-Host "`nSetting the recovery environment.." -ForegroundColor Yellow
        Invoke-Cmdline -application "$($this.scRoot)Windows\System32\reagentc" -argumentList "/SetREImage /Path $($this.recovery.FullName) /target $($this.scRoot)Windows" -silent
    }

    makeDiskBootable() {
        Write-Host "Setting the boot environment.." -ForegroundColor Yellow
        Invoke-Cmdline -application "$($this.scRoot)Windows\System32\bcdboot" -argumentList "$($this.scRoot)Windows /s s: /f all"
    }

    applyUnattend() {
        Write-Host "Looking for unattended.xml.." -ForegroundColor Yellow
        if (-Not Test-Path "$($this.winPESource)scripts\unattended.xml" -ErrorAction SilentlyContinue) {
            Write-Host "Nothing found. Moving on.." -ForegroundColor Red
        }

        Write-Host "Found it! Copying over to scratch drive.." -ForegroundColor Green
        if(-not (Test-Path "$($this.scRoot)Windows\Panther" -ErrorAction SilentlyContinue)){
            New-Item -Path "$($this.scRoot)Windows\Panther" -ItemType Directory -Force | Out-Null
        }

        Copy-Item -Path "$($this.winPESource)\scripts\unattended.xml" -Destination "$($this.scRoot)Windows\Panther\unattended.xml" | Out-Null
    }

    injectProvisioningPackages() {
        Write-Host "`nlooking for *.ppkg files.." -ForegroundColor Yellow
        if (-Not Test-Path "$($usb.winPESource)scripts\*.ppkg" -ErrorAction SilentlyContinue) {
            Write-Host "Nothing found. Moving on.." -ForegroundColor Yellow
            return
        }

        Write-Host "Found them! Copying over to scratch drive.." -ForegroundColor Yellow
        Copy-Item -Path "$($usb.winPESource)\scripts\*.ppkg" -Destination "$($usb.scRoot)Windows\Panther\" | Out-Null
    }

    injectOSDrivers() {
        $modelDriverPath = "$($usb.DriverPath)\$deviceModel"
        if (-Not Test-Path $modelDriverPath) {
            return
        }

        Write-Host "`nApplying drivers.." -ForegroundColor Yellow
        Invoke-Cmdline -application "DISM" -argumentList "/Image:$($usb.scRoot) /Add-Driver /Driver:$modelDriverPath /recurse"
    }
}

#region Main process
try {
    Set-PowerPolicy -powerPlan HighPerformance

    $imageDeploy = [ImageDeploy]::new($env:SystemDrive)

    $imageDeploy.loadDriversInWinPE()
    $imageDeploy.prompt()

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    Write-Host "`nSetting Install.Wim location.." -ForegroundColor Yellow
    if (!($usb.installRoot)) {
        throw "Coudn't find install.wim anywhere..."
    }

    $imageDeploy.prepareDisk()
    $imageDeploy.deployImage()
    $imageDeploy.postDeploy()

    $completed = $true
}
catch {
    $errorMsg = $_.Exception.Message
}
finally {
    $sw.stop()

    if ($errorMsg) {
        Write-Warning $errorMsg
    } else {
        if ($completed) {
            Write-Host "Provisioning process completed..`nTotal time taken: $($sw.elapsed)" -ForegroundColor Green
        } else {
            Write-Host "Provisioning process stopped prematurely..`nTotal time taken: $($sw.elapsed)" -ForegroundColor Green
        }
    }
}
#endregion

Write-Host "Press ENTER to restart"
Read-Host
Restart-Computer