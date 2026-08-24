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
- `data/reference/products.xlsx` en `data/reference/locations.xlsx`: vaste
  referentiebestanden voor de vulgraadberekening (zie Roadmap 2.0
  hieronder). Deze staan wél gewoon in git (bewuste keuze, zie Roadmap).
  Bij een nieuwe export: bestand met dezelfde naam vervangen, én
  `REF_DATA_VERSION` bovenaan `scripts/script.js` ophogen — anders toont
  een refresh van de live tool tot 10 minuten lang nog de oude versie
  (browsercache van GitHub Pages).

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

## Roadmap 2.0 — slimme consolidatie
Huidige versie telt alleen *hoeveel* pallets een artikel inneemt. 2.0 moet
rekening houden met hoe vol/leeg die pallets écht staan, zodat de
reachtruck-planning gericht kan worden op consolidaties die ook
daadwerkelijk ruimte opleveren — niet op artikelen die toevallig op 2
pallets staan maar al bijna vol zijn.

**Vastgestelde aannames (product owner):**
- Magazijn-specifiek: deze tool wordt maar in 1 magazijn gebruikt, dus
  locatie-afmetingen wijzigen zelden.
- Bulklocaties verschillen sterk van afmeting per locatie — geen
  standaardmaat, dus per Location Code de eigen afmetingen gebruiken.
- Producten worden vrijwel altijd rechtop opgeslagen — vaste oriëntatie,
  geen rekening houden met roteren/kantelen van een product.
- Gewicht is geen harde grens — alleen volume (ruimte) telt mee.
- Veel producten zitten in een omdoos, maar die afmetingen zijn niet
  bekend. Bewuste keuze: omdoos negeren bij de berekening, en in plaats
  daarvan een vaste marge aanhouden op de vulgraad (pallet iets eerder
  als "vol" beschouwen dan de kale volumeberekening aangeeft).
- Referentiedata (locatie- en productafmetingen) komt uit een los
  WMS/ERP-bestand en wordt als vaste bestanden in deze repo gezet (map
  nog te bepalen, bijv. `data/reference/`) — niet elke sessie opnieuw
  te uploaden, alleen bijwerken bij wijzigingen.
- Repo is en blijft public. Voor nu gaat referentiedata gewoon plain
  (leesbaar) mee in git. Optie om dit later te versleutelen met een
  wachtwoord (echte encryptie, geen ingebakken sleutel) staat als losse
  fase op de roadmap, nog niet gebouwd.

**Fase 1 — Referentiedata (gebouwd)**
`data/reference/products.xlsx` en `data/reference/locations.xlsx` staan
vast in de repo en worden door de tool zelf ingeladen (fetch, geen
upload-stap). Koppelveld producten: `Product ID` ↔ `Product` in de
voorraadexport. Koppelveld locaties: `Location` ↔ `Location Code`.
Bevestigd: `Quantity` in de voorraadexport is altijd losse eenheden,
rechtstreeks te vermenigvuldigen met de productafmetingen (geen
dozen/verpakkingsfactor nodig). `products.xlsx` is bewust een subset
(niet elk product staat erin) — ontbrekende afmetingen geven "onbekend",
geen foutmelding of gok.
Let op: fetch() van deze bestanden werkt alleen als de pagina via een
webserver bediend wordt (dus via de GitHub Pages-URL), niet bij lokaal
openen door dubbelklikken op `index.html`. Voor lokaal testen is een
kleine lokale server nodig.

**Fase 2 — Vulgraad per pallet (gebouwd)**
Per regel: (aantal × productvolume) ÷ (bruikbare locatie-inhoud) =
vulgraad %. Bruikbare locatie-inhoud = locatielengte × locatiebreedte ×
(locatiehoogte − 200 mm voor de europallet zelf) × (1 − 15% marge voor de
onbekende omdoos). Beide getallen (200 mm, 15%) zijn een inschatting, geen
gemeten waarde — staan als `PALLET_HOOGTE_MM` en `OMDOOS_MARGE` bovenaan
`scripts/script.js` en zijn daar aan te passen als de praktijk daar
aanleiding toe geeft.

Zichtbaar als kolom in de preview én in het geëxporteerde .xlsx-bestand.
De statistieken splitsen "onbekend" uit naar reden (product niet in
`products.xlsx`, locatie niet in `locations.xlsx`, ongeldig aantal, of
locatie te laag na de pallet-aftrek) — zo is in de tool zelf te zien
wáárom iets onbekend is, zonder de browserconsole nodig te hebben.
(Extra: bij het laden logt de tool ook een samenvatting in de
browserconsole (F12 → Console), voor verdere diagnose.)

Dekking hangt volledig af van hoe compleet `products.xlsx` is. Bij de
eerste test (276 producten, een subset) was vrijwel alles "onbekend".
Na een volledigere export (2856 producten) is dat gemeten op de
voorbeelddata: ~30% van de resultaatregels bekend, en van de unieke
artikelen in het resultaat stond nog 67% niet in `products.xlsx` — dat is
dus een dekkingsvraagstuk in het bronbestand, geen bug in de tool. Als de
dekking omhoog moet, is een vollediger productbestand (zonder filter op
bijv. "actief") de aangewezen oplossing.

**Fase 3 — Consolidatiepotentieel (kern van 2.0)**
Per artikel op 2+ pallets: totaal benodigd volume afzetten tegen het
volume van de grootste locatie die het artikel al gebruikt → minimaal
aantal pallets nodig. Verschil met huidig aantal pallets = aantal
locaties dat écht vrijgemaakt kan worden. Dit wordt de nieuwe
prioriteitsscore/sortering, in plaats van alfabetisch.

**Fase 4 — Werklijst voor de reachers**
Resultaat wordt een prioriteitenlijst (meeste winst bovenaan) i.p.v. een
platte tabel, met filter/sortering op vulgraad en vrij te maken locaties.

**Fase 5 — Later / optioneel**
- Wachtwoord-encryptie voor de referentiebestanden in de repo (zie boven).
- Preciezere volumeberekening (per laag/oriëntatie i.p.v. simpele
  volumeratio) als de eenvoudige aanpak in de praktijk niet nauwkeurig
  genoeg blijkt.
- Duidelijke melding bij artikelen/locaties waarvan afmetingen ontbreken,
  i.p.v. laten verdwijnen of fout laten rekenen.

**Nog open voordat Fase 1 gebouwd kan worden:** voorbeeld van het
locatie-afmetingenbestand en het product-afmetingenbestand (kolomnamen,
eenheden) nog aan te leveren door de product owner.
