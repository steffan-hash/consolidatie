# Project: Pallet Consolidatie

## Overzicht
Kleine interne tool voor logistiek/operations werk. Filtert een WMS-
voorraadexport (.xlsx) op artikelen die verspreid over meerdere pallets op
bulklocaties staan — kandidaten om fysiek samen te voegen tot 1 pallet.
Persoonlijk project. Draait op GitHub voor version control, o.a. omdat het
op meerdere plekken (thuis/werk) wordt bewerkt.

## Waarom deze tool bestaat (het echte doel)
De tool moet klaar zijn vóór de piek: de goederen voor het nieuwe seizoen
komen binnen terwijl er nog weinig verkocht wordt. Vorig jaar moest daarvoor
een **extern magazijn met 7000 gevulde europallet-plekken** bijgehuurd
worden. Elke pallet-plek die intern vrijgemaakt kan worden vóór die instroom
is dus directe winst.

Daarom is **"vrij te maken pallet-plekken" de kernmaat** van deze tool — niet
de vulgraad, niet het aantal regels. Alles in de UI en de export is daarop
gericht: hoeveel plekken levert deze lijst op, en welke pallets moeten
daarvoor leeg.

## Status
Techkeuze: één losse `index.html` + `scripts/script.js`, geen build/install-
stap. Excel-verwerking draait volledig client-side in de browser via SheetJS
(inlezen) en ExcelJS (opmaak/output van het resultaat), beide via CDN. Geen
Python/Streamlit. Live via GitHub Pages.

De tool is functioneel af (upload → filter → export). Het rekenmodel is in
augustus 2026 herbouwd van een volumeratio naar **capaciteit in stuks** (zie
"Model 3.0" hieronder) — het oude model rekende structureel verkeerd en liet
placeholder-data bovenaan de werklijst komen.

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
- Het exportbestand bevat 4 vaste kolommen uit de bron (niet instelbaar):
  Location Code, Product Name (= `Description` uit de bron), Quantity, Urn —
  de rest van de brondata is ruis voor het fysiek consolideren van pallets.
  Daarnaast 4 berekende kolommen: Vulgraad, Restruimte, Vrij te maken
  locaties en Actie (zie Model 3.0).
- Resultaat wordt gesorteerd op vrij te maken locaties (meeste winst
  bovenaan), dan op product zodat alle pallets van hetzelfde artikel bij
  elkaar staan, en binnen een artikel op hoeveelheid oplopend — zodat de
  pallets die leeg moeten (Empty) bovenaan staan.
- Printopmaak van de export staat vast op A4 liggend, geschaald naar 1
  pagina breed.
- **Ruisreductie:** (1) Producten met "DOOS", "BOX" of "TOP" als los woord in
  de naam (verpakkingsmateriaal) worden helemaal genegeerd — staat als
  `NOISE_PRODUCT_KEYWORDS` in `scripts/script.js`. (2) Locaties waarvan de
  `Location Code` "CHITA" bevat (bijv. `CHITA_AM1`, `CHITA_DOOS10`) worden
  genegeerd — dat zijn geen gewone bulklocaties in het magazijnrek
  (`NOISE_LOCATION_KEYWORD`). (3) Pallets waarop meerdere artikelen door
  elkaar staan (dezelfde `Urn`, verschillende producten) krijgen geen advies:
  de capaciteit is dan niet aan één artikel toe te rekenen. Alle uitsluitingen
  zijn zichtbaar in de statistieken (niet stilzwijgend).
- **Vervallen regel:** de eerdere ruisreductie "gelijkmatige stapeling" (een
  artikel op >10 pallets met overal dezelfde hoeveelheid én vulgraad negeren)
  is in 3.0 geschrapt. Die bestond alleen om artikelen te verbergen die het
  oude volumemodel onterecht als kans aanmerkte. Met een capaciteitsmodel valt
  een volle pallet vanzelf weg (0 vrij te maken plekken), dus de regel zou nu
  echte kansen gaan verbergen. Dit sluit de openstaande vraag uit de sessies
  van 24/25 augustus af.
- **Alleen echte kansen tonen** (checkbox, standaard aan): artikelen waar
  consolideren geen enkele plek vrijmaakt, of waar de capaciteit onbekend is,
  staan niet in de lijst. Uitzetten laat alles zien.

## Model 3.0 — capaciteit in stuks (huidige aanpak)
Versie 2.0 rekende de vulgraad uit als volume-tegen-volume. Dat bleek
structureel verkeerd: een pallet wordt *gestapeld*, niet volgegoten. Een
artikel van 980 mm lang laat zich niet halveren om de laatste 220 mm te
vullen. Gemeten gevolg: 3 stuks Eurom Flameheater kwamen uit op 48% "vol",
terwijl er fysiek precies 3 op passen — de pallet was 100% vol en werd dus
onterecht als kans op de werklijst gezet.

Daar kwam bij dat een derde van `products.xlsx` (10.401 van 29.996) op
placeholder-afmetingen 1×1×1 mm staat. Zo'n product past volgens een
volumeberekening altijd op één pallet, kreeg dus de **maximale** score bij
"vrij te maken locaties" en een vulgraad van 0% — en kwam daarmee bovenaan de
gesorteerde lijst. De bovenkant van de werklijst werd zo gevuld door
ontbrekende data i.p.v. echte kansen. Dat is de belangrijkste reden dat het
resultaat onbruikbaar aanvoelde.

**De kern van 3.0: capaciteit in stuks per pallet.**
```
stuks per laag = beste van de twee liggingen op de palletvoetprint
aantal lagen   = bruikbare stapelhoogte / producthoogte
capaciteit     = stuks per laag × aantal lagen   (beide naar beneden afgerond)
```
Daaruit volgt alles: vulgraad (`aantal / capaciteit`), **restruimte in stuks**
(het getal waar een chauffeur iets mee kan), het minimaal benodigde aantal
pallets, en dus het aantal vrij te maken plekken.

**Vastgestelde aannames (product owner):**
- Magazijn-specifiek: deze tool wordt maar in 1 magazijn gebruikt, dus
  locatie-afmetingen wijzigen zelden.
- Producten worden vrijwel altijd rechtop opgeslagen — de hoogte staat dus
  vast, maar in het platte vlak mag een product een kwartslag gedraaid worden
  (de tool neemt de beste van die twee liggingen).
- Gewicht is geen harde grens — alleen ruimte telt mee. Bevestigd in de data:
  `Maximum Weight` staat voor alle 9013 locaties op dezelfde waarde en bevat
  dus geen informatie.
- **Stapelhoogte: in de praktijk niet hoger dan ongeveer 2 meter** — een
  natuurlijke grens, geen WMS-regel. Locaties die hoger zijn leveren dus geen
  extra capaciteit op (`MAX_STAPELHOOGTE_MM`).
- **De afmetingen in `products.xlsx` zijn die van de verkoopverpakking.**
  Verkoopt Toppy iets als set, dan staat dat in de productnaam ("... - 8
  stuks") en zijn de afmetingen die van de hele set. `Quantity` telt dus
  dezelfde eenheid als waar de afmetingen bij horen — geen omrekenfactor
  nodig. Hierdoor is de oude omdoos-marge van 15% vervallen: die was een
  dubbele veiligheidsaftrek die het beeld alleen vertekende.
- **Palletvoetprint komt uit de kolom `Urn Type`, niet uit de Length/Width van
  de locatie.** Die laatste spreken elkaar tegen (122 locaties met
  `Urn Type = Euro Pallet` staan als 1800 mm breed geregistreerd) terwijl
  `Urn Type` consistent is: Euro Pallet 7869, Blok Pallet 730, 180 Pallet 320,
  270 Pallet 90, Hottub 15. Zo is het model niet meer gevoelig voor precies de
  datafouten die eerder steeds handmatig gerepareerd zijn. Alleen de
  locatie**hoogte** komt nog uit de locatie zelf.
- **Referentiedata wordt eerst gewantrouwd** (plausibiliteitstoets). Producten
  met 1×1×1, ontbrekende, absurd kleine of absurd grote afmetingen worden
  afgekeurd; locaties met een onmogelijke hoogte of onbekende palletsoort ook.
  Afgekeurd = "onbekend", en onbekend doet **niet** mee aan de rangschikking.
  Gemeten resultaat: 18.220 van 29.996 producten en 9006 van 9039 locaties
  blijven bruikbaar.
- **Zelfcorrectie op de waarneming:** staat er méér op een pallet dan volgens
  de berekening past, dan is de berekening fout (of er wordt ruimer gestapeld
  dan aangenomen) — niet de werkelijkheid. Dan wordt de waargenomen
  hoeveelheid de capaciteit. Dat is conservatief (levert nooit een valse kans
  op) en het aantal keer dat het gebeurt staat in de statistieken als
  signaal over de datakwaliteit.
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
Uit `locations.xlsx` gebruikt de tool alleen `Location`, `Height` en
`Urn Type` — de Length/Width worden bewust genegeerd (zie aannames).
Ontbrekende of onbetrouwbare afmetingen geven "onbekend", geen gok.
Let op: fetch() van deze bestanden werkt alleen als de pagina via een
webserver bediend wordt (dus via de GitHub Pages-URL), niet bij lokaal
openen door dubbelklikken op `index.html`. Voor lokaal testen is een
kleine lokale server nodig.

**Fase 2 — Capaciteit per pallet (gebouwd, 3.0)**
Per regel wordt de capaciteit in stuks berekend volgens de formule bovenaan
dit hoofdstuk. Getoond worden: **Vulgraad** (% van de capaciteit),
**Restruimte** (hoeveel stuks er nog bij kunnen) en de reden als het niet
lukt. Instelbaar bovenaan `scripts/script.js`: `PALLET_HOOGTE_MM` (200),
`MAX_STAPELHOOGTE_MM` (2000), `PALLET_FOOTPRINTS` en de vier
plausibiliteitsgrenzen.

De statistieken splitsen "onbekend" uit naar reden (placeholder-afmetingen,
locatie/palletsoort onbruikbaar, gemengde pallet, product past niet op de
pallet, product te hoog, ongeldig aantal) én tonen de kwaliteit van de
referentiebestanden zelf. Zo is in de tool te zien wáárom iets onbekend is,
zonder de browserconsole nodig te hebben. (Bij het laden logt de tool ook een
samenvatting in de console, F12 → Console.)

**Fase 3 — Vrij te maken plekken (gebouwd, 3.0)**
Per artikel wordt bepaald hoeveel pallets er minimaal nodig zijn: capaciteiten
van groot naar klein optellen tot de totale hoeveelheid erin past. Het
verschil met het huidige aantal pallets is **"Vrij te maken locaties"** — de
kernmaat van de tool. Er wordt per **pallet (Urn)** gerekend, niet per regel,
zodat een artikel met twee regels op dezelfde pallet niet dubbel telt.
Artikelen waarvan ook maar één pallet een onbekende capaciteit heeft krijgen
géén score: dan is niet aan te tonen dat het echt in minder pallets past.
Dit is ook de sortering van het resultaat (meeste winst bovenaan).

**Fase 4 — Werklijst voor de reachers (gebouwd, 3.0)**
Kolom **Actie** geeft per pallet aan wat er moet gebeuren: **Empty**
(leeghalen, voorraad overhevelen) of **Keep** (blijft staan, ontvangt de
voorraad). De pallets met het **meeste** erop blijven staan — dat kost de
minste ritten — en worden aangevuld tot de totale hoeveelheid erin past; wat
dan overblijft kan leeg. Binnen elk artikel staan de Empty-pallets bovenaan.
De labels blijven Engels (kort en duidelijk voor de chauffeurs), de rest van
de tool is Nederlands.

**Getest tot nu toe:** het capaciteitsmodel is buiten de browser om nagerekend
op de échte referentiebestanden (via Excel COM, want er staat geen Node/Python
op deze machine). Uitkomsten kloppen met de verwachting: het Eurom-voorbeeld
komt op 100% (vol, geen kans) en de ondertegels op 65% met 38 stuks
restruimte. Nog **niet** end-to-end in een browser met een echte
voorraadexport getest — die wordt aangeleverd door de product owner.

## Nog open / bewust geparkeerd
**Geparkeerd — de tijdsdimensie (omloopsnelheid).** Idee: met een pick- of
verkoopexport per artikel kan de tool onderscheid maken tussen "nu aanpakken"
en "laat maar, is zo weg" — een pallet die volgende week toch leeggepickt
wordt hoeft geen rit. Bewuste keuze van de product owner: goed idee, maar
buiten scope voor nu. Bewaard als suggestie voor later. (Merk op dat de waarde
hiervan in het piekvenster juist laag is: dan komt er voorraad binnen en wordt
er weinig verkocht, dus is bijna alles een langzame loper.)

**Geparkeerd — AI-verrijking.** Twee onderzochte toepassingen: producten
classificeren op naam (i.p.v. het grove DOOS/BOX/TOP-woordfilter), en het
doosvolume schatten uit gewicht + productgroep voor de 10.401
placeholder-producten. Voor nu overgeslagen op verzoek van de product owner.
Afgeschreven na test: afmetingen uit de productnaam halen werkt niet — slechts
270 van de 29.996 namen bevatten maten, en die beschrijven meestal het
opgebouwde product en niet de doos (een Intex-zwembad van 975×488×132 cm zit
in een doos van 1150×800×1600 mm). Dit gat moet aan de bron dicht.

**Geparkeerd — wachtwoord-encryptie** voor de referentiebestanden in de repo
(zie de aanname over de public repo hierboven).

**Openstaand — datakwaliteit aan de bron.** De grootste beperking is nu
`products.xlsx`: 10.401 van 29.996 producten (35%) staan op placeholder
1×1×1 mm. Elk van die artikelen levert "onbekend" op en verdwijnt dus uit de
werklijst. Wil de dekking omhoog, dan moeten die afmetingen in het bron-WMS
gevuld worden. De tool maakt zichtbaar hoeveel het zijn, maar kan het niet
oplossen.

**Openstaand — te overwegen na de eerste echte test.** In `locations.xlsx`
staat de kolom `Stock on Location`, waaruit blijkt dat er nu al 2.174 locaties
leegstaan (1.824 europallet-plekken, 23%). Die kan als context in de
statistieken (hoeveel plekken zijn er al vrij, hoeveel komen er bij) of om
consolidatie naar een betere lege locatie voor te stellen. Nog niet gebouwd.
