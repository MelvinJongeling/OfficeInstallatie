# OfficeInstallatie

Dit project bevat een PowerShell-script (`InstallOffice.ps1`) waarmee je automatisch Microsoft Office kunt installeren op Windows-systemen. Het script regelt het downloaden van de benodigde Office Deployment Tool (ODT), het ophalen van de juiste configuratiebestanden en het uitvoeren van de installatie.

## Functionaliteit
- **Automatische installatie van Office**: Het script downloadt en installeert Office zonder handmatige stappen.
- **Automatische installatie van dependencies**: De benodigde ODT wordt automatisch gedownload en uitgepakt.
- **Configuratie via variabelen**: Je kunt de te installeren Office-versie instellen via de omgevingsvariabele `$env:officeVersion`. Dit is handig voor gebruik met tools zoals NinjaOne. Gebruik je geen NinjaOne, dan kun je deze variabele handmatig instellen in het script.

## Gebruik
1. Plaats het script `InstallOffice.ps1` op je systeem.
2. Stel de omgevingsvariabele `officeVersion` in op de gewenste Office-versie (bijvoorbeeld `Office21` of `Office365`).
   - Met NinjaOne kun je deze variabele als parameter meegeven.
   - Zonder NinjaOne kun je deze variabele bovenaan het script instellen: ` $env:officeVersion = "Office21" `
3. Voer het script uit met PowerShell als administrator.

## Voorbeeld
```powershell
$env:officeVersion = "Office21"
./InstallOffice.ps1
```

## Opmerking
- Het script downloadt automatisch de benodigde bestanden als ze ontbreken.
- De configuratiebestanden voor Office moeten in de map `Configs` staan. Het script probeert ze te downloaden van GitHub als ze niet aanwezig zijn.

## NinjaOne integratie
Het script is geschikt voor gebruik met NinjaOne door de variabele `$env:officeVersion`. Je kunt deze variabele als parameter meegeven via NinjaOne policies.


## Belangrijk: Configuratiebestanden

De meegeleverde configuratiebestanden in de map `Configs` zijn specifiek gemaakt voor mijn eigen omgeving. Deze werken waarschijnlijk niet voor andere organisaties of situaties. 

**Wil je zelf Office installeren met eigen instellingen?**
- Maak je eigen configuratiebestanden en sla deze lokaal op in de map `Configs` op de werkplek.
- Of geef in het script je eigen URL op waar de configuratiebestanden te downloaden zijn.

Zorg ervoor dat je configuratiebestanden aansluiten bij jouw licentie, gewenste Office-versie en instellingen.
