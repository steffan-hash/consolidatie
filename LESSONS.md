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

## LESSON 1: BOM in xlsx-XML-onderdeel breekt browser-parsing, terwijl .NET het gewoon leest
**Wat ging fout:** Na het direct bewerken van `locations.xlsx` via een PowerShell-script (xlsx als zip/XML behandeld, om 70 locatiehoogtes aan te passen) gaf de live tool "referentiebestanden konden niet geladen worden" — terwijl het bestand met .NET zelf prima te openen was en de juiste waarden bevatte. Pas zichtbaar na een echte test op de GitHub Pages-URL, niet bij de eigen (.NET-)controle.
**Waarom:** `[System.Xml.XmlDocument]::Save(Stream)` schrijft in .NET standaard een UTF-8 BOM (bytes EF BB BF) vóór de XML-declaratie. OOXML-onderdelen zoals `xl/worksheets/sheet1.xml` hebben normaal geen BOM, en de browser (SheetJS) kon het bestand met BOM niet meer parsen. .NET's eigen zip/XML-lezer is daar niet kieskeurig over, dus de fout viel niet op bij het herlezen met hetzelfde gereedschap waarmee de wijziging is gemaakt.
**Regel:** Bij het direct bewerken van een xlsx-XML-onderdeel via .NET/PowerShell altijd expliciet zonder BOM wegschrijven (`New-Object System.Text.UTF8Encoding($false)` als encoding voor de `XmlWriter`), en de eerste bytes van het gewijzigde onderdeel na afloop controleren op EF BB BF. Alleen checken of .NET het bestand terugleest is onvoldoende bewijs dat de browser het ook kan.
