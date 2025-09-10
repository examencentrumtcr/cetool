# Cetool

Een tool om Eduarte-examenbestanden om te zetten naar het Facet-formaat

## 1. Introductie

Cetool is ontwikkeld voor de mbo-school Techniek College Rotterdam in Nederland waar **Facet** wordt gebruikt om digitale examens klaar te zetten.

Facet heeft specifieke gegevens nodig, zoals studenteninformatie, groepen, examentypen en examendata. Deze gegevens komen uit het studentenadministratiesysteem **Eduarte** in de vorm van een Excel-bestand. Omdat Facet alleen XML-bestanden kan verwerken, moet dit bestand eerst worden omgezet. Cetool automatiseert dit proces.

Daarnaast accepteert Facet alleen groepen studenten van één school. Om examens met studenten uit meerdere scholen toch mogelijk te maken, splitst Cetool het Eduarte Excel-bestand automatisch in meerdere XML-bestanden. Elk bestand bevat uitsluitend studenten van één school.

### Belangrijkste functies (v1 – uitgebracht september 2025)

* **Excel → XML conversie**

  * Zet een Eduarte Excel-bestand om naar meerdere XML-bestanden, per examen georganiseerd.
  * Studenten met een speciale indicatie (bijv. dyslexie) komen in een apart XML-bestand zodat ze extra tijd kunnen krijgen.

* **Logboek**

  * Bewaart gebeurtenissen in een logboek.
  * Het logboek kan handmatig worden geleegd of automatisch oude meldingen verwijderen bij het opstarten (ouder dan ingesteld aantal dagen).

* **Voorkeuren**

  * Slaat gebruikersinstellingen op, zoals:

    * Map met in te lezen Excel-bestanden
    * Map voor de uitvoerbestanden (ZIP)
    * Naam van het uitvoerbestand
    * Automatisch logboekbeheer
    * Sluitgedrag van de console

* **Automatische updates**

  * Haalt updates direct op via GitHub.
  * Beheerders kunnen ook prerelease-versies downloaden.

---

## 2. Installatie

### Systeemvereisten

* Windows 10 of hoger
* PowerShell 5.1 of hoger

### Installatie via setup (aanbevolen)

1. Download de installer vanuit:
   [Setup-map op GitHub](https://github.com/examencentrumtcr/cetool/tree/main/setup)
2. Pak `setup-cetool.zip` uit.
3. Start `setup-cetool.exe`.

> ⚠️ Let op: Windows SmartScreen kan een waarschuwing tonen omdat het bestand niet digitaal is ondertekend. Wanneer het bestand van de officiële repository komt, kan het veilig uitgevoerd worden.

Tijdens de installatie kun je:

* De installatiemap wijzigen
* Een snelkoppeling op het bureaublad maken
* Een schone installatie uitvoeren

Wij adviseren om de standaardopties te gebruiken en **Installeren** te kiezen.

### Handmatige installatie

1. Download de laatste release vanuit:
   [Release-map op GitHub](https://github.com/examencentrumtcr/cetool/tree/main/release)
2. Pak de ZIP uit in een map waarin de gebruiker schrijfrechten heeft.

> Let op: Bij handmatige installatie moet je zelf een snelkoppeling maken en rekening houden met PowerShell-beperkingen.

---

## 3. Configuratie

Configuratie is alleen nodig als:

* Er geen snelkoppeling is aangemaakt, of
* PowerShell het uitvoeren van scripts blokkeert.

### Snelkoppeling maken

Voer `snelkoppeling_maken.ps1` uit in PowerShell. Dit maakt een snelkoppeling en stelt direct de juiste instellingen in.

### Uitvoeringsbeleid aanpassen

Je kunt ook `Wijzig_Executionpolicy_bypass.bat` uitvoeren. Hiermee wordt PowerShell in *Bypass*-modus gezet.

Als geen van beide opties wordt gebruikt, kan PowerShell bij het starten van Cetool een waarschuwing geven over het uitvoeringsbeleid.

---

## 4. Gebruik

* Als er een snelkoppeling is aangemaakt: **dubbelklik** om Cetool te starten.
* Zonder snelkoppeling: klik met de rechtermuisknop op het script en kies **Openen met > Windows PowerShell** of **Uitvoeren met PowerShell**.

---

## 5. Bestandenoverzicht

Het project bevat de volgende bestanden en mappen:

```
Cetool.ps1                        Hoofdscript
script_icoon.ico                  Icoontje gebruikt door het script
Logboek.png                       Afbeelding gebruikt door het script
start.png                         Afbeelding gebruikt door het script
stoppen.png                       Afbeelding gebruikt door het script
Snelkoppeling_maken.ps1           Tool om snelkoppeling te maken
Wijzig_Executionpolicy_bypass.bat Tool om PowerShell uitvoeringsbeleid op 'Bypass' te zetten
Changelog.md                      Lijst met wijzigingen per versie
Readme.md                         Deze documentatie
License                           GNU General Public Licentie overeenkomst
```

---

## 6. Licentie

Cetool is gelicenseerd onder de **GNU General Public License v3.0 (GPLv3)**.

Dit programma is vrije software: je mag het verspreiden en/of aanpassen onder de voorwaarden van de GPL, uitgegeven door de Free Software Foundation.
Het wordt geleverd **zonder enige garantie**, ook niet voor verkoopbaarheid of geschiktheid voor een bepaald doel.

Zie de [GNU GPL v3.0](https://www.gnu.org/licenses/gpl-3.0.html) voor meer informatie.

---

## 7. Bekende problemen

Er zijn momenteel geen bekende problemen.
Mocht je een fout ontdekken, meld dit dan via [GitHub Issues](https://github.com/examencentrumtcr/cetool/issues).

---

## 8. Auteurs en dankwoord

**Hoofdauteur**

* [Benvindo Neves](https://neveshuis.nl/over-mij)

**Bijdrage geleverd door**

* Rob Schortemeijer

---

## 9. Changelog

Bekijk de [Changelog](https://github.com/examencentrumtcr/cetool/tree/main/CHANGELOG.md) voor details.

---
