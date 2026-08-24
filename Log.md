# Log

Dit bestand wordt door Claude bijgehouden, niet handmatig door de product
owner. Nieuwste sessie bovenaan. Zie CLAUDE.md → "Sessie einde" voor het
format.

## Sessie 2026-08-24
**Status:** De "-45"-locatiehoogtefix (100mm placeholder → hoogte van de bijbehorende "-40"-locatie) is opnieuw doorgevoerd in `data/reference/locations.xlsx`, nu op een manier die het laden in de browser niet meer zou moeten breken. Nog niet gepusht — wacht op akkoord van de product owner (zie "Nog open").

**Wat gedaan — oorzaak van de vorige laadfout gevonden:**
- De vorige sessie eindigde met: BOM-theorie (byte-order-mark) klopte niet, want ook het bevestigd werkende bestand had al een BOM op dezelfde plekken. Ware oorzaak nog onbekend.
- Zonder een echte browser te gebruiken (geen Node/Python beschikbaar op deze machine) is dit alsnog hard aangetoond: de eerder gepushte "gefixte" bestanden (commits `73d0f16` en `3a49acf`) uit de git-historie gehaald en elk via Excel COM-automatisering geopend (dus onafhankelijk van elke browser). Excel weigerde beide bestanden te openen — ook in reparatiemodus — terwijl het bevestigd werkende bestand (`ef7cdd2`/huidig) foutloos opende. Reproduceerbaar met verse Excel-instanties, dus geen toeval.
- Zip-structuur (Central Directory/Local File Headers) en XML-validiteit (geen ongeldige controletekens, geen locale-decimalen zoals "1500,0") van het gefixte bestand zelf gecontroleerd met een handgeschreven validatiescript — beide in orde. De corruptie zit dus ergens in hoe `[System.Xml.XmlDocument]::Save()` de volledige (1MB+) worksheet-XML herserialiseerde, niet in een makkelijk te isoleren enkel detail. Conclusie: xlsx-bestanden niet meer rechtstreeks als XML/zip bewerken via .NET — zie bijgewerkte `LESSONS.md` (LESSON 1, tekst gecorrigeerd i.p.v. een nieuwe tegenstrijdige les toegevoegd, conform CLAUDE.md).
- Fix opnieuw uitgevoerd, nu via Excel's eigen COM-objectmodel (Excel zelf leest het bestand in, past de 70 hoogtes aan via `Range.Value2`, en slaat zelf op als .xlsx) — dus Excel schrijft het bestand, niet een los .NET-scriptje. Zelfde 70 locaties gevonden en gefixt als in de vorige sessie (0 zonder "-40"-tegenhanger, 0 met een zelf verdacht lage "-40"-hoogte).
- Geverifieerd: Excel opent het resulterende bestand weer foutloos (verse instantie), 9040 rijen intact, 0 van de 630 "-45"-locaties heeft nog hoogte 100, steekproef van onaangeraakte rijen/kolommen ongewijzigd. `scripts/script.js` gebruikt van dit bestand alleen `Location`, `Length`, `Width`, `Height` (gecontroleerd) — andere kolommen zijn dus sowieso niet relevant, ook al zou Excel's bulk-lezen/schrijven daar een klein typeverschil in geven.
- Backup gemaakt vóór het overschrijven: `data/reference/locations.xlsx.backup-20260824-154337` (niet meegecommit, staat er puur lokaal ter herstel — de vorige bevestigd werkende versie staat toch al in git-historie).

**Nog open:**
- Nog niet gecommit/gepusht — gezien dit exacte bestand de live tool twee keer eerder heeft gebroken, eerst expliciet akkoord van de product owner vragen voordat dit weer naar `main`/GitHub Pages gaat, ook al staat in `CLAUDE.md` dat direct committen naar `main` normaal geen aparte goedkeuring nodig heeft.
- Er staat ook nog een oudere losse back-up in de werkmap van 2026-08-19 (`data/reference/locations.xlsx.backup-20260819-112331`, niet in git) — niet verwijderd, ter beoordeling aan de product owner of die weg kan.
- Fase 3 van de roadmap (prioriteitsscore: hoeveel locaties een consolidatie daadwerkelijk vrijmaakt) staat nog steeds open.
**Volgende stap:** Na akkoord: committen en pushen, en de product owner vragen de live tool nog een keer te verversen (F5) om te bevestigen dat de referentiebestanden weer laden én dat de "-45"-locaties nu een geldige vulgraad tonen i.p.v. "onbekend".

## Sessie 2026-08-19
**Status:** Tool is overgezet naar deze repo en live via GitHub Pages. Roadmap 2.0 (vulgraad-gebaseerde consolidatie) is besproken; Fase 1 (referentiedata) en Fase 2 (vulgraad per pallet) zijn gebouwd, getest en na een terugkoppeling van de product owner verder afgesteld (pallethoogte-aftrek, reden-diagnostiek). Fase 3 (prioriteitsscore "locaties vrij te maken") staat nog open.

**Vervolg dezelfde sessie — veel vollediger productbestand:**
- Product owner had een opgeschoonde Google Sheet met productdata; op advies (xlsx, sluit aan bij hoe de tool referentiebestanden inleest) geëxporteerd en `data/reference/products.xlsx` mee vervangen (665 KB → 1,55 MB).
- Gecontroleerd (buiten de browser om, via een PowerShell-uitlezing van de xlsx zelf, want er was geen voorbeeld-voorraadexport meer lokaal aanwezig om de tool end-to-end te draaien): koprij bevat exact de verwachte kolommen (`Product ID`, `Description`, `Product Group`, `Barcode`, `Length`, `Width`, `Height`, `Weight`). 29.996 productregels, geen lege of dubbele Product ID's, en 28.914 daarvan (96%) hebben een volledige, geldige lengte/breedte/hoogte. Dat is een grote sprong t.o.v. de vorige 2.856 producten (waarvan destijds 67% van de artikelen in de resultaten geen match had) — dekking van de vulgraadberekening zou hier flink door moeten verbeteren.
- Oude bestand (2.856 producten) staat nog gewoon in de git-historie (commit `478246d`) — geen aparte back-up nodig.
- Nog niet end-to-end getest in de browser met een echte voorraadexport (die stond niet meer lokaal klaar); structuur en volledigheid van het bestand zelf zijn wel bevestigd.

**Vervolg dezelfde sessie — diagnose en fix "-45"-locaties met onbekende vulgraad:**
- Product owner meldde: bijna alle locaties die eindigen op "-45" (bijv. `16-3-A-136-45`) geven een onbekende vulgraad.
- Diagnose (via de xlsx-bestanden zelf uitgelezen, buiten de browser om): 70 van de 630 "-45"-locaties in `locations.xlsx` hadden een hoogte van precies 100 mm — steeds exact dezelfde waarde, geen enkele variatie. Na de vaste pallet-hoogte-aftrek van 200 mm (`PALLET_HOOGTE_MM`) wordt de bruikbare hoogte dan 0, vandaar "onbekend" (reden "locatie te laag"). Geen bug in de tool — wel een plausibele datafout in het referentiebestand: 10 cm is geen realistische pallethoogte, en het steeds identieke getal wijst op een placeholder/standaardwaarde in het bron-WMS in plaats van een echte meting.
- Op verzoek van de product owner opgelost: voor alle 70 betrokken locaties is de hoogte overschreven met de hoogte van de corresponderende "-40"-locatie op dezelfde positie (bijv. `16-3-A-136-45` kreeg de hoogte van `16-3-A-136-40`). Alle 70 hadden een bruikbare "-40"-tegenhanger, geen enkele hoefde overgeslagen te worden, en geen van de bron-"-40"-hoogtes was zelf verdacht laag.
- Uitgevoerd met een directe bewerking van `data/reference/locations.xlsx` (via een PowerShell-script dat de xlsx als zip/XML behandelt, dus zonder Excel of extra libraries nodig te hebben). Vooraf een lokale kopie gemaakt (`locations.xlsx.backup-20260819-112331`, niet meegecommit — de originele versie staat toch al in de git-historie van vóór deze wijziging). Na de wijziging gecontroleerd: bestand blijft geldig (rijaantal ongewijzigd, 9040), en 0 van de "-45"-locaties heeft nog een hoogte ≤ 200 mm.
- Niet aangepast: de overige ~57 lage (≤200 mm) locaties die niet op "-45" eindigen — dat viel buiten deze vraag en is niet onderzocht.

**Vervolg dezelfde sessie — gepushte fix bleek de tool zelf te breken, hersteld:**
- Na het pushen meldde de product owner op de live tool: "referentiebestanden konden niet geladen worden, vulgraad wordt niet berekend."
- Live URL gecontroleerd (`curl` op de GitHub Pages-URL): deployment was up-to-date, beide bestanden gaven gewoon 200 OK met de juiste bestandsgrootte en CORS-header — dus geen deploy-vertraging of CORS-probleem.
- Oorzaak gevonden door de ruwe bytes van `xl/worksheets/sheet1.xml` in het net-bewerkte `locations.xlsx` te inspecteren: het bestand begon met een UTF-8 BOM (EF BB BF), veroorzaakt doordat `XmlDocument.Save(Stream)` in .NET die standaard toevoegt. Browsers (SheetJS) accepteren dat niet in een OOXML-onderdeel, ook al leest .NET zelf het bestand gewoon terug — vandaar dat dit niet opviel bij de eigen controle na de vorige stap. Zie `LESSONS.md` (LESSON 1).
- Fix opnieuw uitgevoerd, nu vanaf de eerder gemaakte backup (dus zonder op de BOM-fout voort te bouwen) en met een `XmlWriter` die expliciet zonder BOM schrijft. Dezelfde 70 locaties bijgewerkt, gecontroleerd dat de BOM weg is en dat 0 "-45"-locaties nog een hoogte ≤ 200 mm hebben. Bestandsgrootte nu 1.028.285 bytes — nagenoeg gelijk aan het origineel (1.028.434), in lijn met een kleine, gerichte wijziging.
- Gecommit en gepusht; live URL bevestigde de nieuwe bestandsgrootte — maar product owner meldde dat de waarschuwing op de live tool nog steeds verscheen.
- Bij nader onderzoek van het live-bestand bleek ook `xl/workbook.xml` een BOM te hebben — een onderdeel dat ik niet had aangeraakt, dus kennelijk al aanwezig vóór alle wijzigingen van vandaag. Dat betekent dat de BOM-theorie (in ieder geval voor `workbook.xml`) niet de (hele) verklaring kan zijn — de vorige sessie bevestigde immers dat dit bestand toen al gewoon laadde in een echte browser. De precieze oorzaak van de nieuwe laadfout is dus nog niet gevonden.
- Om de live tool niet langer kapot te laten staan tijdens verder uitzoeken: `data/reference/locations.xlsx` teruggezet naar de laatst bevestigd werkende versie (commit `ef7cdd2`, vóór de -45-hoogtefix), gecommit en gepusht (`a9e753e`) en bevestigd dat deze versie weer live staat.
- **Openstaand:** de -45-hoogtefix (100 mm → hoogte van de "-40"-locatie) staat dus weer NIET in het live bestand — die moet opnieuw gedaan worden zodra duidelijk is waarom de directe XML-bewerking de browser-laadfout veroorzaakt. Volgende stap: exacte foutmelding uit de browserconsole (F12) van de product owner opvragen, en/of testen of de fout mogelijk aan `products.xlsx` ligt in plaats van `locations.xlsx` (nog niet end-to-end in een echte browser getest, alleen met .NET gecontroleerd — wat nu bewezen onvoldoende is gebleken).

**Vervolg dezelfde sessie — diagnose "alles onbekend":**
- Product owner testte de live tool: alle regels toonden "onbekend" bij Vulgraad, en vroeg om logging om te zien waarom.
- Diagnose (met de PowerShell-nabouw van de rekenlogica op de échte bestanden): de eerste `products.xlsx` was een subset van 276 producten uit één productgroep (gevaarlijke stoffen) — de eerste 200 (alfabetisch gesorteerde) resultaatregels waren toevallig allemaal andere producten, dus geen van alle kon gematcht worden. Geen bug, wel onvoldoende dekking.
- Product owner leverde een vollediger `products.xlsx` aan (2856 producten). Daarmee: 1823 van 5983 resultaatregels (~30%) hebben nu een bekende vulgraad. Van de 933 unieke artikelen in het resultaat staat 67% (621) nog steeds niet in `products.xlsx` — puur een dekkingsvraagstuk in het bronbestand (0 gevallen van een matchend product zonder afmetingen, dus geen normalisatieprobleem in de koppeling zelf).
- `scripts/script.js` uitgebreid met reden-codes voor "onbekend" (`product-onbekend`, `locatie-onbekend`, `aantal-ongeldig`, `locatie-te-laag`, `referentiedata-niet-geladen`) — nu zichtbaar als uitsplitsing in de statistieken, plus een samenvatting in de browserconsole bij het laden. Zo is voortaan in de tool zelf te zien wáárom een regel onbekend is, zonder devtools nodig te hebben.
- Op verzoek: locatiehoogte wordt nu eerst met 200 mm verminderd (`PALLET_HOOGTE_MM`) voor de vulgraadberekening, om ruimte voor de europallet zelf mee te rekenen. Marge voor de onbekende omdoos blijft 15% (`OMDOOS_MARGE`) — beide instelbaar bovenaan `scripts/script.js`, gedocumenteerd in `PROJECT.md`.
- Opnieuw getest in een echte (headless) browser: geen JS-fouten, referentiedata laadt en de consolelog toont exact dezelfde aantallen als de PowerShell-controle (2855/2856 producten, 9039/9039 locaties).
**Wat gedaan:**
- `index.html`, `scripts/script.js` en een voorbeeldbestand (`data/input/Stock (6).xlsx`) overgezet vanuit een los mapje op het bureaublad (`Desktop/Consalidatie`) waar de tool tot nu toe buiten git om gebouwd was.
- `.gitignore` uitgebreid: `data/input/` en `data/output/` worden genegeerd (met `.gitkeep`), want WMS-voorraadexports zijn bedrijfsdata en horen niet op GitHub.
- `PROJECT.md` ingevuld op basis van de bestaande tool, `.nojekyll` toegevoegd voor GitHub Pages (hosting handmatig aangezet door product owner).
- Roadmap 2.0 besproken en vastgelegd in `PROJECT.md`, incl. vastgestelde aannames (locaties verschillen sterk per stuk, producten altijd rechtop, gewicht telt niet mee, omdoos genegeerd met een marge).
- Product owner leverde `Products.xlsx` en `warehouse_locations.xlsx` aan (in `data/input/`). Structuur gecontroleerd: koppelveld producten = `Product ID` ↔ `Product` in de voorraadexport, koppelveld locaties = `Location` ↔ `Location Code`. Bevestigd: `products.xlsx` is bewust een subset, en `Quantity` in de voorraadexport is altijd losse eenheden.
- Bestanden verplaatst naar `data/reference/products.xlsx` en `data/reference/locations.xlsx` (wél in git, bewuste keuze uit de roadmap-afweging over publieke repo).
- `scripts/script.js`: `findHeaderRow` generiek gemaakt (herbruikbaar voor referentiebestanden), referentiedata wordt nu bij het laden van de pagina automatisch opgehaald (`fetch`, geen upload) en gekoppeld. Nieuwe functie `computeFillRatio` berekent per regel de vulgraad. Preview, statsbox en de geëxporteerde .xlsx tonen nu een "Vulgraad"-kolom; bij ontbrekende afmetingen staat er "onbekend" i.p.v. een gok. `index.html` kreeg een waarschuwingsbalk (`#refDataHint`) die alleen verschijnt als de referentiebestanden niet geladen konden worden.
- Getest: de vulgraad-formule losstaand nagebouwd in PowerShell op de échte data (via de xlsx-bestanden zelf uitgelezen) — resultaten zijn plausibel (1–50% in de steekproef, geen negatieve/absurde waarden). Daarnaast in een echte (headless) browser bevestigd dat `script.js` foutloos laadt en de referentiebestanden succesvol ophaalt en verwerkt via dezelfde CDN-libraries als de tool zelf gebruikt.
**Nog open:**
- Fase 3 van de roadmap (prioriteitsscore: hoeveel locaties een consolidatie daadwerkelijk vrijmaakt) is nog niet gebouwd.
- Dekking van `products.xlsx` is nu ~30% van de resultaatregels (67% van de unieke artikelen mist nog een match). Als hogere dekking gewenst is: een vollediger productbestand aanleveren (mogelijk zit er een "actief"-filter op de huidige export).
- Het oude mapje `Desktop/Consalidatie` staat nog op het bureaublad (incl. een oudere, afwijkende `CLAUDE.md`). Niet verwijderd — graag zelf beoordelen of dat weg kan.
**Volgende stap:** Zelf de bijgewerkte statistieken-uitsplitsing checken via GitHub Pages. Als de aanpak (200 mm pallet-aftrek, 15% omdoos-marge) logisch aanvoelt: Fase 3 bouwen (sorteren op daadwerkelijk vrij te maken locaties i.p.v. alfabetisch).

<!--
Voorbeeld van een entry:

## Sessie 2026-08-18
**Status:** [één regel: waar staat het project]
**Wat gedaan:** [wat er echt gebouwd/aangepast is, met bestandsnamen]
**Nog open:** [wat de product owner moet weten of beslissen]
**Volgende stap:** [wat er logisch als eerste opgepakt wordt]
-->
