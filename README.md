# teams-public

## Install required PS Modules
[Prereqs PS1](./install-prereqs.ps1)

Open Powershell 5 as Admin, and run:
```
iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/keithflannery/teams-public/main/install-prereqs.ps1'))
```

## Using my Teams PS Functions

Create a new PS script and paste the following:
```
iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/keithflannery/teams-public/main/teams-includes.ps1'))


$username = 'XXXX@xxxx.onmicrosoft.com'
$password = 'XXX'

Disconnect-Sessions
Get-Connected -username $username  -password $password

$voiceskus = "BUSINESS_VOICE_DIRECTROUTING", "MCOCV", "MCOEV", "BUSINESS_VOICE_DIRECTROUTING", "ENTERPRISEPREMIUM_NOPSTNCONF"

```

Once you've executed the script, and you should be connected to teams online, and be able to use the functions inside of [Teams Includes](./teams-includes.ps1)