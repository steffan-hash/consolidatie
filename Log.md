# Log

Dit bestand wordt door Claude bijgehouden, niet handmatig door de product
owner. Nieuwste sessie bovenaan. Zie CLAUDE.md → "Sessie einde" voor het
format.

## Sessie 2026-08-19
**Status:** Tool is functioneel af (upload → filter → export) en nu overgezet naar deze repo, zodat de vaste werkwijze (git, deze logbestanden) erop van toepassing is.
**Wat gedaan:**
- `index.html`, `scripts/script.js` en een voorbeeldbestand (`data/input/Stock (6).xlsx`) overgezet vanuit een los mapje op het bureaublad (`Desktop/Consalidatie`) waar de tool tot nu toe buiten git om gebouwd was.
- `.gitignore` uitgebreid: `data/input/` en `data/output/` worden genegeerd (met `.gitkeep` om de mapstructuur te behouden), want WMS-voorraadexports zijn bedrijfsdata en horen niet op GitHub.
- `PROJECT.md` ingevuld op basis van de bestaande tool: techkeuze (losse `index.html`, client-side via SheetJS/ExcelJS, geen build), verwachte kolommen, en de niet-vanzelfsprekende businessregels (alleen "Bulk Location" telt mee, consolidatie = 2+ pallets, vaste output-kolommen).
**Nog open:**
- Het oude mapje `Desktop/Consalidatie` staat nog op het bureaublad (incl. het CLAUDE.md-bestand daar, dat een oudere/afwijkende versie is). Niet verwijderd — graag zelf beoordelen of dat weg kan nu alles overgezet is.
- `PROJECT.md` is door Claude ingevuld op basis van de code; graag even doorlezen of dit klopt met wat er bedoeld is, met name de sectie "Context die niet vanzelfsprekend is".
**Volgende stap:** Doorbouwen aan de tool vanuit deze repo (bijv. verdere filters of exportopties), of eerst `PROJECT.md` samen scherpstellen als er dingen niet kloppen.

<!--
Voorbeeld van een entry:

## Sessie 2026-08-18
**Status:** [één regel: waar staat het project]
**Wat gedaan:** [wat er echt gebouwd/aangepast is, met bestandsnamen]
**Nog open:** [wat de product owner moet weten of beslissen]
**Volgende stap:** [wat er logisch als eerste opgepakt wordt]
-->
