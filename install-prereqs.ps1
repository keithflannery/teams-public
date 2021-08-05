Write-Host "Checking for elevated permissions..."
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
            [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Insufficient permissions to run this script. Open the PowerShell console as an administrator and run this script again."
    Break
}
else {
    if ($PSVersionTable.PSVersion.Major -eq '5') {
        Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
        $Installedmodules = Get-InstalledModule
        $modules = "MicrosoftTeams", "MSOnline", "AzureAD", "AZ", "ImportExcel"

        foreach ($checkModule in $modules) {
            if ($Installedmodules.Name -notcontains $checkModule) {
                Install-Module $checkModule
            }
        }   
    }
    else {
        Write-Warning "This needs to be run in PS version 5!"
    }
}
