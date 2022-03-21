#requires -Modules @{ ModuleName="WindowsAutoPilotIntune"; ModuleVersion="4.3" }
#requires -Modules @{ ModuleName="Microsoft.Graph.Intune"; ModuleVersion="6.1907.1.0"}
function Get-AutopilotPolicy {
    [cmdletbinding()]
    param (
        [parameter(Mandatory = $true)]
        [System.IO.FileInfo]$FileDestination
    )
    try {
        if (Test-Path "$FileDestination" -ErrorAction SilentlyContinue) {
            Write-Host "Autopilot Configuration directory found locally: $FileDestination" -ForegroundColor Green
            return
        }

        New-Item -ItemType Directory -Path "$FileDestination" | Out-Null

        $modules = @(
            "WindowsAutoPilotIntune",
            "Microsoft.Graph.Intune"
        )
        if ($PSVersionTable.PSVersion.Major -eq 7) {
            $modules | ForEach-Object {
                Import-Module $_ -UseWindowsPowerShell -ErrorAction SilentlyContinue 3>$null
            }
        } else {
            $modules | ForEach-Object {
                Import-Module $_
            }
        }
        #region Connect to Intune
        Connect-MSGraph | Out-Null
        #endregion Connect to Intune
        #region Get policies
        $apPolicies = Get-AutopilotProfile
        if (!($apPolicies)) {
            Write-Warning "No Autopilot policies found.."
            return
        }

        foreach ($policy in $apPolicies) {
            # NTFS does not allow colons in filesnames, which can occur in the display name.
            $strippedName = $policy.displayName.Replace(":", "")

            $policy | ConvertTo-AutopilotConfigurationJSON | Out-File "$FileDestination\$strippedName.json" -Encoding ascii -Force
            Write-Host "Saved autopilot config: " -ForegroundColor Yellow -NoNewline
            Write-Host $policy.displayName
        }
        #endregion Get policies
    }
    catch {
        $errorMsg = $_
    }
    finally {
        if ($PSVersionTable.PSVersion.Major -eq 7) {
            $modules = @(
                "WindowsAutoPilotIntune",
                "Microsoft.Graph.Intune"
            ) | ForEach-Object {
                Remove-Module $_ -ErrorAction SilentlyContinue 3>$null
            }
        }
        if ($errorMsg) {
            Write-Warning $errorMsg
        }
    }
}