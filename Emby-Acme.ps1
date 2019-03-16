#requires -runasadministrator
$ErrorActionPreference = "Stop"

Install-PackageProvider -Name NuGet -Force
Install-Module -Name ACMESharp.Providers.IIS -Force
Import-Module ACMESharp
Enable-ACMEExtensionModule -ModuleName ACMESharp.Providers.IIS -ErrorAction SilentlyContinue

if (-not (Get-ACMEVault)) {
    Initialize-ACMEVault
}

try {
    (Get-ACMERegistration).Contacts
}
catch {
    New-ACMERegistration -Contacts "mailto:$(Read-Host -Prompt 'Enter email address')" -AcceptTos
}

if ($serviceName = (Get-Service | Where-Object {$_.name -match "emby"} | Select-Object -first 1).name) {
    $appLocation = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\$serviceName\Parameters").Application
    $location = (Get-Item $appLocation).Directory.Parent.FullName
}
elseif (Test-Path "$Env:APPDATA\Emby-Server") {
    $location = "$Env:APPDATA\Emby-Server"
}
else {
    $appLocation = Read-Host -Prompt "Enter Emby server location"
    if ($appLocation.Length -eq 0) {
        throw "Unknown Location"
    }
    elseif (Test-Path -Path $appLocation -PathType Leaf) {  
        $location = (Get-Item $appLocation).Directory.Parent.FullName
    }
    elseif (Test-Path -Path $appLocation -PathType Container) {
        $location = (Get-Item $appLocation).Parent.FullName
    }
    else {
        throw "Cannot locate file or folder named $appLocation"
    }
}

if (Test-Path "$location\programdata\config\system.xml") {
    $serverConfiguration = ([xml](Get-Content "$location\programdata\config\system.xml")).ServerConfiguration
}
elseif (Test-Path "$location\config\system.xml") {
    $serverConfiguration = ([xml](Get-Content "$location\config\system.xml")).ServerConfiguration
}
else {
    throw "Cannot find system.xml at either $location\programdata\config\system.xml or $location\config\system.xml"
}

$address = $serverConfiguration.WanDdns
if ($address.Length -eq 0) {
    throw "Domain name not found in emby config"
}
$alias = "emby-$($address.Split(".")[0])-$(get-date -format yyyy-MM-dd--HH-mm)"

if ((Get-WindowsOptionalFeature -Online -FeatureName IIS-WebServerRole).State -ne "Enabled") {
    Enable-WindowsOptionalFeature -Online -FeatureName IIS-WebServerRole
}
    
if ((Get-WebConfiguration -filter /system.webServer/handlers -PSPath IIS:\ -Location 'Default Web Site' -metadata).metadata.overrideMode -ne 'Allow') {
    Set-WebConfiguration -filter /system.webServer/handlers -PSPath IIS:\ -Location 'Default Web Site' -metadata overrideMode -value Allow
}

function New-Identifiter {
    New-ACMEIdentifier -Dns $address -Alias $alias
    Complete-ACMEChallenge $alias -ChallengeType http-01 -Handler iis -HandlerParameters @{ WebSiteRef = 'Default Web Site' }
    Submit-ACMEChallenge $alias -ChallengeType http-01
    $i = 0
    do {
        $identinfo = (Update-ACMEIdentifier $alias -ChallengeType http-01).Challenges | Where-Object {$_.Status -eq "valid"}
        $i++
        Write-Progress "Completing Identifiter" -PercentComplete ($i * 10)
        Start-Sleep 1
        if ($i -ge 10) {
            throw "Did not receive a completed Identifiter"
        }
    } until($identinfo.Length -ne 0)
    "Valid Identifier: $alias"
}

function New-Certificate {
    New-ACMECertificate $alias -Generate -Alias $alias
    Submit-ACMECertificate $alias
    $i = 0
    do {
        $certinfo = Update-AcmeCertificate $alias
        $i++
        Write-Progress "Completing Certificate" -PercentComplete ($i * 10)
        Start-Sleep 1
        if ($i -ge 10) {
            throw "Did not receive a completed certificate"
        }
    } until($certinfo.SerialNumber -ne "")
    "Valid Certificate: $alias"
}

try {
    $Identifiers = Get-ACMEIdentifier 
}
catch {
    continue
}

$validIdentifiers = @()
ForEach ($Identifier in $Identifiers) {
    if ($Identifier.Dns -eq $address) {
        try {
            $vaildIdentifier = (Update-ACMEIdentifier $Identifier.Alias -ChallengeType http-01).Challenges | Where-Object {$_.Status -eq "valid"}
        }
        catch {
            continue
        }
        if ($vaildIdentifier.Length -ne 0) {
            $validIdentifiers += $Identifier.Alias
        }
    }
}
if ($validIdentifiers.Length -ne 0) {
    $alias = $validIdentifiers[0]
    "Valid Identifier: $alias"
}
else {
    New-Identifiter
}

try {
    Get-ACMECertificate $alias
}
catch {
    New-Certificate
}

$certPath = $serverConfiguration.CertificatePath
if ($certPath.Length -eq 0) {
    $certPath = "$location\programdata\$address.pfx"
}

Get-ACMECertificate $alias -ExportPkcs12 $certPath -Overwrite
