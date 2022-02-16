Write-Host "Checking for elevated permissions..."
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
            [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Insufficient permissions to run this script. Open the PowerShell console as an administrator and run this script again."
    Break
}
else {
    Write-Host "Running as elevated user...." -ForegroundColor Green
    if ($PSVersionTable.PSVersion.Major -eq '5') {
        Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted


        $Installedmodules = Get-InstalledModule
        if ($Installedmodules.Name -contains "MicrosoftTeams"){
            $version = Get-InstalledModule -Name MicrosoftTeams | Select-Object Version
            if ($version -ne 3.1.1){
                Uninstall-Module -Name MicrosoftTeams
                Install-Module -Name MicrosoftTeams -RequiredVersion 3.1.1
            }
        }
        else 
        {
           Install-Module -Name MicrosoftTeams -RequiredVersion 3.1.1
        }
        
        $modules = "MSOnline", "AzureAD", "ImportExcel", "AzureAD"

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
