#Requires -Modules @{ ModuleName="MicrosoftTeams"; RequiredVersion="3.1.1" }
#Requires -Module MSOnline
#Requires -Module AzureAD
#Requires -Module ImportExcel

import-Module MicrosoftTeams, MSOnline, ImportExcel

function Get-Connected {

    param (
        [parameter(Mandatory = $true)]
        $username
        ,
        [parameter(Mandatory = $true)]
        $password
    )

    $SecurePassword = ConvertTo-SecureString $password -AsPlainText -Force
    $credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $username, $SecurePassword
    Connect-MsolService -Credential $credential
    Connect-AzureAD -Credential $credential
    Connect-MicrosoftTeams -Credential $credential
}

function Get-Connected-MFA {
    write-host "Connecting to 365 Powershell..." -ForegroundColor Green
    Connect-MsolService
    Write-Host "Connecting to AzureAD Powershell..." -ForegroundColor Green
    Connect-AzureAD
    Write-Host "Connecting to Microsoft Teams Powershell..." -ForegroundColor Green
    Connect-MicrosoftTeams -Credential $credential
}

function Disconnect-Sessions {
    Disconnect-MicrosoftTeams
    Get-PSSession | Remove-PSSession
}

Function XString {
    Param (
        [Parameter(Mandatory = $true)]
        [string]$string
        ,
        [Parameter(Mandatory = $true)]
        [char]$character
        ,
        [Parameter(Mandatory = $false)]
        [ValidateSet("Right", "Left")]
        [string]$range
        ,
        [Parameter(Mandatory = $false)]
        [int]$afternumber
        ,
        [Parameter(Mandatory = $false)]
        [int]$tonumber
    )
    Process {
        [string]$return = ""
        
        if ($range -eq "Right") {
            $return = $string.Split("$character")[($string.Length - $string.Replace("$character", "").Length)]
        }
        elseif ($range -eq "Left") {
            $return = $string.Split("$character")[0]
        }
        elseif ($tonumber -ne 0) {
            for ($i = $afternumber; $i -le ($afternumber + $tonumber); $i++) {
                $return += $string.Split("$character")[$i]
            }
        }
        else {
            $return = $string.Split("$character")[$afternumber]
        }
        
        return $return
    }
}


function Get-KF-LicensedUsers	{

    param (
        [string[]]$licenseskus
    )

    $licensedusers = @()

    $msolusers = Get-MsolUser -All | Where-Object { ($_.isLicensed -eq "TRUE" -and $_.BlockCredential -ne $true ) }
    
    foreach ($user in $msolusers) {
        # if ($user.isLicensed -eq $true) {
        foreach ($sku in $user.Licenses.AccountSkuID) {
            $fuxsku = XString -string $sku -character ":" -range Right
            if ($licenseskus.Contains($fuxsku)) {
                $licensedusers += $user.UserPrincipalName
            }
        }
        # }
    }

    return $licensedusers
}


function Get-KF-Data {

    param (
        [parameter(Mandatory = $true)]
        $licensedusers
    )

    $ht = @{}
    foreach ($user in $licensedusers) {

        $data = "FirstName", "LastName", "EnterpriseVoiceEnabled", "HostedVoiceMail", "LineURI", "UsageLocation", "UserPrincipalName", "WindowsEmailAddress", "SipAddress", "OnPremLineURI", "OnlineVoiceRoutingPolicy", "TenantDialPlan", "HostingProvider", "TeamsUpgradeEffectiveMode", "OnPremLineURIManuallySet", "TeamsIPPhonePolicy"

        $teamsdata = get-csonlineuser -id $user

        $datahash = @{}
        foreach ($x in $data) { $datahash += @{$x = $teamsdata.$x } }
        $dataobject = [pscustomobject]$datahash

        $ht += @{$user = $dataobject }
    }

    return $ht
}

function new-kf-csuser {

    param (
        [parameter(Mandatory = $true)]
        $UPN
        ,
        [parameter(Mandatory = $true)]
        $LineURI
        ,
        [parameter()]
        $OnlineVoiceRoutingPolicy = 'MyTeamsVoice'
        ,
        [parameter()]
        $TenantDialPlan
    )

    $licensedusers2 = Get-KF-LicensedUsers -licenseskus $voiceskus

    # if ($licensedusers2.Contains($UPN)) {
    if ($licensedusers2 -contains $UPN) {

        $userdata = get-csonlineuser -id $UPN

        if ($TenantDialPlan -eq $null) {
            if ($LineURI -match '^tel:\+612\d{8}(?:|;ext\=\d+)$') {
                $TenantDialPlan = "AU-02"
            }
            elseif ($LineURI -match '^tel:\+613\d{8}(?:|;ext\=\d+)$') {
                $TenantDialPlan = "AU-03"
            }
            elseif ($LineURI -match '^tel:\+617\d{8}(?:|;ext\=\d+)$') {
                $TenantDialPlan = "AU-07"
            }
            elseif ($LineURI -match '^tel:\+618\d{8}(?:|;ext\=\d+)$') {
                $TenantDialPlan = "AU-08"
            }
            else {
                Write-Error "Line URI not a valid Australian Landline Number e.g.: tel:+61736249100"
                exit
            }
        }

        Set-CsUser -Identity $UPN -EnterpriseVoiceEnabled $true 
        #if ($userdata.OnPremLineURIManuallySet -eq $false){
        Set-CSUser -Identity $UPN -LineURI $LineURI
        #}
        Grant-CsOnlineVoiceRoutingPolicy -id $UPN -PolicyName $OnlineVoiceRoutingPolicy
        Grant-CsTenantDialPlan -id $UPN -PolicyName $TenantDialPlan

    }
    else {
        Write-Error "User does not have a Voice Enabled license assigned, or doesn't exist! Please assign License in Office 365 portal to $UPN!"
    }
}

function get-phonenumbers {
    get-csonlineuser | where-object { $_.LineUri -match '^tel:\+?61[2378]\d{8}(?:|;ext\=\d+)$' } | Select-Object UserPrincipalName, FirstName, LastName, DisplayName, LineUri, City | export-excel
}

function validate-allusers {
    
    param (
        [parameter()]
        $IgnoreEVDisabled = '$false'
    )

    if ($IgnoreEVDisabled -eq $true) {
        $csonlineusers = get-csonlineuser | Where-Object { $_.EnterpriseVoiceEnabled -eq $true } | select-object FirstName, LastName, EnterpriseVoiceEnabled, HostedVoiceMail, LineURI, UsageLocation, UserPrincipalName, WindowsEmailAddress, SipAddress, OnPremLineURI, OnlineVoiceRoutingPolicy, TenantDialPlan, HostingProvider, TeamsUpgradeEffectiveMode, OnPremLineURIManuallySet, TeamsIPPhonePolicy
    }
    else {
        $csonlineusers = get-csonlineuser | select-object FirstName, LastName, EnterpriseVoiceEnabled, HostedVoiceMail, LineURI, UsageLocation, UserPrincipalName, WindowsEmailAddress, SipAddress, OnPremLineURI, OnlineVoiceRoutingPolicy, TenantDialPlan, HostingProvider, TeamsUpgradeEffectiveMode, OnPremLineURIManuallySet, TeamsIPPhonePolicy        
    }
    $licensedusers2 = Get-KF-LicensedUsers -licenseskus $voiceskus

    $data = @()

    foreach ($user in $csonlineusers) {
        if ($licensedusers2 -contains $user.UserPrincipalName) {
            $data += $user
        }
    }

    foreach ($user in $data) {
        $borkedusers = @{}
        $BORKED = $false
        $reasons = New-Object System.Collections.Generic.List[string]

        if ($user.LineURI -notmatch '^tel:\+61[2378]\d{8}(?:|;ext\=\d+)$') {
            $BORKED = $true
            $reasons.Add("LineURI Invalid!")
        }
        if ($user.EnterpriseVoiceEnabled -eq $false) {
            $BORKED = $true
            $reasons.Add("EnterpriseVoiceEnabled is False!")
        }
        if ($user.TenantDialPlan -eq $null) {
            $BORKED = $true
            $reasons.Add("TenantDialPlan is Empty!")
        }
        if ($user.OnlineVoiceRoutingPolicy -eq $null) {
            $BORKED = $true
            $reasons.Add('OnlineVoiceRoutingPolicy is Empty!')
        }

        if ($BORKED -eq $true) {
            $borkedusers += @{$user.UserPrincipalName = $reasons }


        }

        $borkedusers

    }
}


function get-userdetails {
    get-csonlineuser | Where-Object { $_.EnterpriseVoiceEnabled -eq $true } | select-object FirstName, LastName, EnterpriseVoiceEnabled, HostedVoiceMail, LineURI, UsageLocation, UserPrincipalName, WindowsEmailAddress, SipAddress, OnPremLineURI, OnlineVoiceRoutingPolicy, TenantDialPlan, HostingProvider, TeamsUpgradeEffectiveMode, OnPremLineURIManuallySet, TeamsIPPhonePolicy | Export-Excel
}




function delete-kf-csuser {

    param (
        [parameter(Mandatory = $true)]
        $UPN
    )
    Set-CSUser -Identity $UPN -OnPremLineURI $null
    Set-CsUser -Identity $UPN -EnterpriseVoiceEnabled $false
}

function new-kf-resourceaccount {
    param (
        [parameter(Mandatory = $true)]
        $ratype,
        $UPN,
        $DisplayName,
        $URI,
        $IgnoreWarning = $false
    )
    $sku = Get-MsolAccountSku | where-object {$_.AccountSkuId -match '^.+\:PHONESYSTEM_VIRTUALUSER'}

    write-host "This will take approximatly 'Microsoft' 4 minutes to run..." -ForegroundColor Green
    if ($IgnoreWarning -eq $false) {
        write-host "Please ensure you have a spare 'Virtual Phone System' User!" -ForegroundColor Yellow
        $confirmation = Read-Host "Ok? [y/n]"
        while($confirmation -ne "y"){
            if ($confirmation -eq 'n') {exit}
            $confirmation = Read-Host "Ok? [y/n]"
        }
    }
    if ($ratype -eq 'aa'){
        $appid = 'ce933385-9390-45d1-9512-c8d228074e07'
    }
    Elseif ($ratype -eq 'cq'){
        $appid = '11cd3e2e-fccb-42ad-ad00-878b93575e07'
    }
    else {
        write-host 'Need to specifiy resourse account type -ratype aa (auto attendent), -ratype cq (call queue)!' -ForegroundColor Red
        exit
    }
    New-CsOnlineApplicationInstance -UserPrincipalName $UPN -DisplayName $DisplayName -ApplicationId $appid
    start-sleep -Seconds 120
    Set-MsolUser -UserPrincipalName $UPN -UsageLocation "AU"
    Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses $sku.AccountSkuId
    Start-Sleep -Seconds 120
    Set-CsOnlineApplicationInstance -Identity $UPN -OnpremPhoneNumber $URI
    start-sleep -Seconds 5
    write-host "process complete - please make sure the phone number is set below (if you set one!):" -ForegroundColor Yellow
    Get-CsOnlineApplicationInstance -Identity $UPN
}