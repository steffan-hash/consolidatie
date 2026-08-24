# Lessons

Dit bestand wordt door Claude bijgehouden, niet handmatig door de product
owner. Een les komt er alleen bij als iets daadwerkelijk een keer fout is
gegaan — niet vooraf verzonnen op basis van wat er mis zou kunnen gaan.

Format per les: wat ging fout, waarom, wat is nu de regel.

<!--
Voorbeeld:

## LESSON 1: [korte titel]
**Wat ging fout:** [concrete situatie]
**Waarom:** [oorzaak]
**Regel:** [controleerbare regel, geen advies]
-->

## LESSON 1: xlsx rechtstreeks als XML bewerken via .NET `XmlDocument` levert een bestand op dat Excel zelf weigert — BOM was niet de oorzaak
**Wat ging fout:** Na het direct bewerken van `locations.xlsx` via een PowerShell-script (xlsx als zip/XML behandeld met `XmlDocument`, om 70 locatiehoogtes aan te passen) gaf de live tool "referentiebestanden konden niet geladen worden" — terwijl het bestand met .NET zelf prima te openen was en de juiste waarden bevatte. Eerste verklaring was een ontbrekende/foutieve BOM (byte-order-mark) in de XML-onderdelen, en de "fix" daarvoor (BOM verwijderen) loste het probleem niet op — de live tool bleef de foutmelding geven.
**Waarom:** Bij het uitzoeken bleek de BOM-theorie niet te kloppen: het (nooit aangepaste) bevestigd werkende bestand had zélf óók al een BOM in dezelfde XML-onderdelen (`xl/workbook.xml`, `xl/worksheets/sheet1.xml`) en laadde daarmee altijd al gewoon in de browser. De BOM was dus nooit het probleem. Onafhankelijk bevestigd door beide "gefixte" bestanden via Excel COM-automatisering te openen (los van elke browser): Excel weigerde ze allebei te openen (ook in reparatiemodus), terwijl het bevestigd werkende bestand foutloos opende. Zip-structuur en XML-validiteit van het gefixte bestand waren op zich in orde (gecontroleerd: geldige Central Directory/Local File Headers, geen ongeldige controletekens, geen locale-decimalen) — de corruptie zat dus specifiek in hoe `XmlDocument.Save()` de hele (1MB+) worksheet-XML herserialiseerde, niet in een makkelijk te isoleren detail. .NET's eigen lezer is daar kennelijk lakser in dan Excel/browsers.
**Regel:** Een xlsx-bestand niet rechtstreeks als XML/zip bewerken via .NET (`XmlDocument`, `ZipArchive`, etc.) voor waarde-wijzigingen — ook niet voor een klein aantal cellen. Gebruik in plaats daarvan Excel zelf via COM-automatisering (`New-Object -ComObject Excel.Application`, bulk inlezen/wegschrijven via `Range.Value2`, opslaan met `SaveAs(...,51)`) zodat Excel zelf een gegarandeerd geldig bestand schrijft. Controleren dat .NET het bestand kan terugleses is geen bewijs dat het bestand valide is — openen met Excel zelf (niet alleen de browser/live tool) is een snelle, betrouwbare test die niet afhankelijk is van een echte browsersessie.
