# Log

Dit bestand wordt door Claude bijgehouden, niet handmatig door de product
owner. Nieuwste sessie bovenaan. Zie CLAUDE.md → "Sessie einde" voor het
format.

## Sessie 2026-08-19
**Status:** Tool is functioneel af (upload → filter → export), overgezet naar deze repo, en live via GitHub Pages. Roadmap voor 2.0 (slimme consolidatie op basis van vulgraad) is besproken en vastgelegd in `PROJECT.md`, bouwen moet nog beginnen.
**Wat gedaan:**
- `index.html`, `scripts/script.js` en een voorbeeldbestand (`data/input/Stock (6).xlsx`) overgezet vanuit een los mapje op het bureaublad (`Desktop/Consalidatie`) waar de tool tot nu toe buiten git om gebouwd was.
- `.gitignore` uitgebreid: `data/input/` en `data/output/` worden genegeerd (met `.gitkeep` om de mapstructuur te behouden), want WMS-voorraadexports zijn bedrijfsdata en horen niet op GitHub.
- `PROJECT.md` ingevuld op basis van de bestaande tool: techkeuze (losse `index.html`, client-side via SheetJS/ExcelJS, geen build), verwachte kolommen, en de niet-vanzelfsprekende businessregels (alleen "Bulk Location" telt mee, consolidatie = 2+ pallets, vaste output-kolommen).
- `.nojekyll` toegevoegd zodat GitHub Pages de site direct serveert zonder Jekyll-verwerking; hosting handmatig door product owner aangezet via repo-instellingen.
- Roadmap 2.0 besproken (idee: vulgraad per pallet berekenen met locatie- en productafmetingen, zodat consolidatie-prioriteit op écht vrij te maken ruimte gebaseerd wordt i.p.v. alleen pallet-aantal) en vastgelegd in `PROJECT.md` onder "Roadmap 2.0", incl. vastgestelde aannames: locaties verschillen sterk per stuk, producten altijd rechtop, gewicht telt niet mee, omdoos wordt genegeerd met een marge op de vulgraad, referentiedata komt als vaste bestanden in de repo (voor nu plain/leesbaar; wachtwoord-encryptie is een latere fase).
**Nog open:**
- Het oude mapje `Desktop/Consalidatie` staat nog op het bureaublad (incl. het CLAUDE.md-bestand daar, dat een oudere/afwijkende versie is). Niet verwijderd — graag zelf beoordelen of dat weg kan nu alles overgezet is.
- Voor Fase 1 van de roadmap heb ik een voorbeeld nodig van het locatie-afmetingenbestand en het product-afmetingenbestand (kolomnamen, eenheden) — nog aan te leveren.
- Te checken zodra dat bestand er is: staat de kolom "Quantity" in de voorraadexport altijd in losse eenheden, of soms in dozen/verpakkingen? Bepaalt of er een omrekenfactor nodig is.
**Volgende stap:** Referentiebestanden (locaties + producten) aanleveren, dan Fase 1 van de roadmap bouwen (referentiedata inladen in de tool).

<!--
Voorbeeld van een entry:

## Sessie 2026-08-18
**Status:** [één regel: waar staat het project]
**Wat gedaan:** [wat er echt gebouwd/aangepast is, met bestandsnamen]
**Nog open:** [wat de product owner moet weten of beslissen]
**Volgende stap:** [wat er logisch als eerste opgepakt wordt]
-->
