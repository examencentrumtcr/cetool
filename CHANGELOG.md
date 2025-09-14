# Changelog

Alle belangrijke wijzigingen in dit project worden hier gedocumenteerd.

## [1.0.0] - 2025-09-12
### Added
- Mogelijkheid om het logboek handmatig te legen of automatisch bij de start van het script oudere gebeurtenissen te verwijderen.  
- Automatisch updaten van het script via GitHub.  
  - Beheerders kunnen ook een prerelease versie downloaden door dit in te stellen in het gebruikersbestand.  
- Uitvoermap instelbaar via variabelen.  
- Toegevoegd: verschillende modes in het script:  
  - **alpha**: logbestanden verwijderen = uit, automatische updates = uit, console blijft open.  
  - **beta**: automatische updates = uit, console blijft open.  
  - **prerelease**: controle op prerelease updates.  
  - **release**: normaal gebruik.  
- Extra controles toegepast voor betere foutafhandeling.  
- Persoonlijke voorkeuren worden bewaard voor hergebruik, o.a.:  
  - map met Excel-bestanden,  
  - uitvoermap voor gecomprimeerde bestanden,  
  - naam van het uitvoerbestand,  
  - logboek automatisch legen (ja/nee),  
  - aantal dagen waarna gebeurtenissen worden verwijderd,  
  - console sluiten (ja/nee).  
- Bij start van het script worden een aantal opschoon acties uitgevoerd.

### Fixed
- Probleem opgelost bij het vergelijken van versienummers (werden eerder als karakters gezien).  

---

## [0.3.0] -  2025-07-15
### Added
- Hoofdfunctie: Excel-bestand uit **Eduarte** kan worden omgezet naar meerdere XML-bestanden en als één gecomprimeerd bestand bewaard (compatibel met **Facet**).  
- XML-bestanden zijn per examen verdeeld.  
- Studenten met speciale indicaties (zoals dyslexie) worden in een apart XML-bestand geplaatst zodat zij extra tijd krijgen.  
- Startvenster met drie keuzes:  
  - **Start omzetten**  
  - **Logboek bekijken**  
  - **Afsluiten**  
- Interactieve workflow bij het omzetten:  
  - Excel-bestand kiezen → voortgangsvenster → overzicht ingelezen data → bevestigen of opnieuw kiezen → laatste venster met overzicht en instelbare uitvoernaam/-map → **Omzetten**.  
  - Voortgang tijdens omzetten wordt weergegeven.  
- Alle gebeurtenissen worden in het logboek bewaard.  
- Afsluiten-optie sluit het script netjes af.  
