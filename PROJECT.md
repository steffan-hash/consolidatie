# Project: Pallet Consolidatie

## Overzicht
Kleine interne tool voor logistiek/operations werk. Filtert een WMS-
voorraadexport (.xlsx) op artikelen die verspreid over meerdere pallets op
bulklocaties staan — kandidaten om fysiek samen te voegen tot 1 pallet.
Persoonlijk project. Draait op GitHub voor version control, o.a. omdat het
op meerdere plekken (thuis/werk) wordt bewerkt.

## Status
Techkeuze gemaakt: één losse `index.html` + `scripts/script.js`, geen
build/install-stap. Excel-verwerking draait volledig client-side in de
browser via SheetJS (inlezen) en ExcelJS (opmaak/output van het resultaat),
beide via CDN. Geen Python/Streamlit.

De tool is functioneel af (upload → filter → export). Tot nu toe gebouwd
buiten deze repo om (los mapje op het bureaublad); vanaf deze sessie is dat
overgezet naar deze repo zodat de vaste werkwijze (git, Log.md, LESSONS.md)
erop van toepassing is.

## Belangrijkste commando's
```
# Geen install/build nodig — direct openen in de browser:
start index.html
```

## Data & bestanden
- Brondata: .xlsx voorraadexport vanuit het WMS, handmatig geëxporteerd
  (o.a. via Gmail) en via de UI geüpload — geen command-line verwerking.
- Omdat het handmatige exports zijn: het format kan per export licht
  verschillen (kolomvolgorde, extra tabbladen, lege rijen bovenaan). De
  tool zoekt daarom zelf de koprij op, over alle tabbladen en de eerste
  ~20 regels heen.
- Verwachte kolommen (moeten aanwezig zijn, volgorde maakt niet uit):
  `Product`, `Location Code`, `Urn`, `StockLocationTypeName`. Optioneel:
  `Product Group` (voor de groepsfilter in de UI).
- Voorbeeldbestand staat in `data/input/` voor lokaal testen. Deze map en
  `data/output/` zijn gitignored — WMS-exports zijn bedrijfsdata en horen
  niet op GitHub.

## Conventies
- Scripts in `scripts/`
- Downloadbestand vanuit de tool: `rapportage_consolidatie_pallets_YYYY-MM-DD.xlsx`

## Context die niet vanzelfsprekend is
- **Alleen "Bulk Location" telt mee als pallet-locatie.** "Bulk Location
  Extern" en pick-locaties tellen niet mee, want die pallets kun je niet
  fysiek samenvoegen met de rest van het magazijn.
- **Consolidatie-kandidaat** = een artikel dat op 2 of meer unieke pallets
  (verschillende `Urn`-waarden) op een Bulk Location staat. Artikelen op
  precies 1 pallet worden standaard verborgen (instelbaar via checkbox).
- Het exportbestand bevat altijd maar 4 vaste kolommen (niet instelbaar):
  Location Code, Product Name (= `Description` uit de bron), Quantity, Urn
  — de rest van de brondata is ruis voor het fysiek consolideren van
  pallets.
- Resultaat wordt gesorteerd op product, dan locatie — zodat alle pallets
  van hetzelfde artikel bij elkaar staan in de export.
- Printopmaak van de export staat vast op A4 liggend, geschaald naar 1
  pagina breed.
