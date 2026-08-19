# Log

Dit bestand wordt door Claude bijgehouden, niet handmatig door de product
owner. Nieuwste sessie bovenaan. Zie CLAUDE.md → "Sessie einde" voor het
format.

## Sessie 2026-08-19
**Status:** Tool is overgezet naar deze repo en live via GitHub Pages. Roadmap 2.0 (vulgraad-gebaseerde consolidatie) is besproken; Fase 1 (referentiedata) en Fase 2 (vulgraad per pallet) zijn gebouwd en getest. Fase 3 (prioriteitsscore "locaties vrij te maken") staat nog open.
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
- De volledige upload-en-bekijk-flow (bestand kiezen → Vulgraad-kolom zien in de preview) is nog niet door een mens in een browser gecontroleerd — wel de onderliggende berekening en het laden van de referentiedata. Aanrader: eenmaal zelf de live GitHub Pages-versie proberen met een voorraadbestand.
- Het oude mapje `Desktop/Consalidatie` staat nog op het bureaublad (incl. een oudere, afwijkende `CLAUDE.md`). Niet verwijderd — graag zelf beoordelen of dat weg kan.
- Dekking van de vulgraad is nu laag (`products.xlsx` is een subset) — dat is verwacht gedrag, geen bug, maar goed om te weten bij het beoordelen van de eerste resultaten.
**Volgende stap:** Zelf de nieuwe Vulgraad-kolom checken via GitHub Pages, dan (als de marge/aanpak logisch aanvoelt) Fase 3 bouwen: sorteren op daadwerkelijk vrij te maken locaties i.p.v. alfabetisch.

<!--
Voorbeeld van een entry:

## Sessie 2026-08-18
**Status:** [één regel: waar staat het project]
**Wat gedaan:** [wat er echt gebouwd/aangepast is, met bestandsnamen]
**Nog open:** [wat de product owner moet weten of beslissen]
**Volgende stap:** [wat er logisch als eerste opgepakt wordt]
-->
