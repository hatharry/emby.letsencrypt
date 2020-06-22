#requires -runasadministrator
$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-PackageProvider -Name NuGet -Force
Install-Module -Name Posh-ACME -Scope AllUsers -Force

Set-PAServer LE_PROD
if (-not (Get-PAAccount | Select-Object -first 1).Contact) {
    New-PAAccount -AcceptTOS -Contact "$(Read-Host -Prompt 'Enter email address')"
}

if ($serviceName = (Get-Service | Where-Object { $_.name -match "emby" } | Select-Object -first 1).name) {
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

New-PAOrder $address -PfxPass "" -Force

Invoke-HttpChallengeListener -Verbose

New-PACertificate $address

$certPath = $serverConfiguration.CertificatePath
if ($certPath.Length -eq 0) {
    $certPath = "$location\programdata\$address.pfx"
}

$pfxFile = (Get-PACertificate $address | Where-Object { $_.NotAfter -gt (Get-Date) } | Select-Object -first 1).PfxFile

if ($pfxFile.Length -gt 0) {
    Copy-Item $pfxFile $certPath -Force
}
