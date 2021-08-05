Write-Host "Checking for elevated permissions..."
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
            [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Insufficient permissions to run this script. Open the PowerShell console as an administrator and run this script again."
    Break
}
else {
    $Installedmodules = Get-InstalledModule
    $modules = "MicrosoftTeams", "MSOnline", "AzureAD", "AZ", "ImportExcel"

    foreach ($checkModule in $modules) {
        if ($Installedmodules.Name -notcontains $checkModule) {
            Install-Module $checkModule
        }
    }   
}