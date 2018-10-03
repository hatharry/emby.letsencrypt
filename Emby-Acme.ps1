#requires –runasadministrator
$ErrorActionPreference = "Stop"

Install-PackageProvider -Name NuGet -Force
Install-Module -Name ACMESharp -AllowClobber -Force
Install-Module -Name ACMESharp.Providers.IIS -Force
Import-Module ACMESharp
Enable-ACMEExtensionModule -ModuleName ACMESharp.Providers.IIS -ErrorAction SilentlyContinue

if (-not (Get-ACMEVault)){
    Initialize-ACMEVault
}

try {
    (Get-ACMERegistration).Contacts
} catch {
    New-ACMERegistration -Contacts "mailto:$(Read-Host -Prompt 'Enter email address')" -AcceptTos
}

$serviceName = (Get-Service | Where-Object {$_.name -match "emby"} | Select-Object -first 1).name
if ($serviceName.Length -eq 0){
    $appLocation = Read-Host -Prompt "Emby exe location"
    if ($appLocation.Length -eq 0){
        throw "Unknown Location"
    }
    #Test if appLocation specified is embyserver.exe or the folder that embyserver.exe is in.
    if (Test-Path -Path $appLocation -PathType Leaf){  
        $location = (Get-Item $appLocation).Directory.Parent.FullName  #is File
    }
    elseif (Test-Path -Path $appLocation -PathType Container){
        $location = (Get-Item $appLocation).Parent.FullName #is Folder
    }
    else {
        throw ("Cannot locate file or folder named", $appLocation)
    }
} else {
    try {
        $appLocation = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\$serviceName\Parameters").Application
        $location = (Get-Item $appLocation).Directory.Parent.FullName
    } catch {
        $appLocation = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Services\$serviceName").ImagePath.split("`"")[1]
        $location = (Get-Item $appLocation).Directory.Parent.FullName
    }
}

$serverConfiguration = ([xml](Get-Content "$location\config\system.xml")).ServerConfiguration
$address = $serverConfiguration.WanDdns
$alias = "emby-$($address.Split(".")[0])-$(get-date -format yyyy-MM-dd--HH-mm)"

if ((Get-WindowsOptionalFeature -Online -FeatureName IIS-WebServerRole).State -ne "Enabled"){
    Enable-WindowsOptionalFeature -Online -FeatureName IIS-WebServerRole}
    
if ((Get-WebConfiguration -filter /system.webServer/handlers -PSPath IIS:\ -Location 'Default Web Site' -metadata).metadata.overrideMode -ne 'Allow'){
    Set-WebConfiguration -filter /system.webServer/handlers -PSPath IIS:\ -Location 'Default Web Site' -metadata overrideMode -value Allow
}

function New-Identifiter {
    New-ACMEIdentifier -Dns $address -Alias $alias
    Complete-ACMEChallenge $alias -ChallengeType http-01 -Handler iis -HandlerParameters @{ WebSiteRef = 'Default Web Site' }
    Submit-ACMEChallenge $alias -ChallengeType http-01
    $i = 0
    do {
        $identinfo = (Update-ACMEIdentifier $alias -ChallengeType http-01).Challenges | Where-Object {$_.Status -eq "valid"}
        if($identinfo.Length -eq 0) {
            Start-Sleep 6
            $i++
        }
    } until($identinfo.Length -ne 0 -or $i -gt 10)
    if ($i -gt 10){
        throw "Did not receive a completed Identifiter after 60 seconds"
    } else {
        "Valid Identifier: $alias"
    }
}

function New-Certificate {
    New-ACMECertificate $alias -Generate -Alias $alias
    Submit-ACMECertificate $alias
    $i = 0
    do {
        $certinfo = Update-AcmeCertificate $alias
        if($certinfo.SerialNumber -eq "") {
            Start-Sleep 6
            $i++
        }
    } until($certinfo.SerialNumber -ne "" -or $i -gt 10)
    if ($i -gt 10){
        throw "Did not receive a completed certificate after 60 seconds"
    } else {
        "Valid Certificate: $alias"
    }
}

try {
    $Identifiers = Get-ACMEIdentifier
    $validIdentifiers = @()
    ForEach ($Identifier in $Identifiers) {
        if ($Identifier.Dns -eq $address){
            try {
                $vaildIdentifier = (Update-ACMEIdentifier $Identifier.Alias -ChallengeType http-01).Challenges | Where-Object {$_.Status -eq "valid"}
            } catch {
                continue
            }
            if ($vaildIdentifier.Length -ne 0) {
                $validIdentifiers += $Identifier.Alias
            }
        }
    }
    if($validIdentifiers.Length -ne 0){
        $alias = $validIdentifiers[0]
        "Valid Identifier: $alias"
    }else{
        New-Identifiter
    }
} catch {
    New-Identifiter
}

try {
    Get-ACMECertificate $alias
} catch {
    New-Certificate
}

$certPath = $serverConfiguration.CertificatePath
if ($certPath.Length -eq 0) {
    $certPath = "$location\ssl\$address.pfx"
}

Get-ACMECertificate $alias -ExportPkcs12 $certPath -Overwrite
