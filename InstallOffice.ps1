# Onattended Office installatie via ODT

# Pad naar ODT en config files
$odtPath = "D:\ODT\setup.exe"
$configDir = "D:\ODT\Configs"

# Download ODT setup indien niet aanwezig
if (!(Test-Path $odtPath)) {
    Write-Host "ODT setup.exe niet gevonden, downloaden..."
    $odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17531-20046.exe"
    $odtInstaller = "D:\ODT\OfficeDeploymentTool.exe"
    if (!(Test-Path "D:\ODT")) { New-Item -ItemType Directory -Path "D:\ODT" | Out-Null }
    Invoke-WebRequest -Uri $odtUrl -OutFile $odtInstaller
    Write-Host "Uitpakken van ODT setup..."
    Start-Process -FilePath $odtInstaller -ArgumentList "/quiet /extract:D:\ODT" -Wait
    Remove-Item $odtInstaller
}

# Definieer beschikbare configs in het script
$configs = @{
    "Office2021" = "D:\ODT\Configs\Office2021.xml"
    "Office365"      = "D:\ODT\Configs\Office365.xml"
}

# Controleer of officeVersion is gezet
if (-not $env:officeVersion) {
    Write-Error "Omgevingsvariabele 'officeVersion' is niet gezet! Kies bijvoorbeeld 'Office2021' of 'Office365'."
    exit 1
}

# Controleer of de gekozen versie bestaat in configs
if (-not $configs.ContainsKey($env:officeVersion)) {
    Write-Error "Ongeldige officeVersion: $env:officeVersion. Kies uit: $($configs.Keys -join ', ')"
    exit 1
}

$configFile = $configs[$env:officeVersion]

# Download config file indien niet aanwezig
if (!(Test-Path $configFile)) {
    # Controleer of de Configs map bestaat, maak aan indien nodig
    if (!(Test-Path $configDir)) {
        New-Item -ItemType Directory -Path $configDir | Out-Null
    }
    Write-Host "Config file $configFile niet gevonden, downloaden van GitHub..."
    $githubUrl = "https://raw.githubusercontent.com/MelvinJongeling/OfficeInstallatie/refs/heads/main/Configs/$env:officeVersion.xml"
    Invoke-WebRequest -Uri $githubUrl -OutFile $configFile
}

if (!(Test-Path $configFile)) {
    Write-Error "Config file $configFile niet gevonden of downloaden mislukt!"
    exit 1
}

Write-Host "Start installatie van Office $env:officeVersion ..."
Start-Process -FilePath $odtPath -ArgumentList "/configure `"$configFile`"" -Wait

Write-Host "Installatie afgerond."
