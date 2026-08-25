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

## LESSON 2: Een gevuld veld is geen geldige waarde — controleer referentiedata op plausibiliteit, niet op aanwezigheid
**Wat ging fout:** In de sessie van 19-08 is `products.xlsx` gecontroleerd en gerapporteerd als "28.914 van de 29.996 producten (96%) hebben een volledige, geldige lengte/breedte/hoogte". Die conclusie is een week lang als waarheid gebruikt en stond zo in `Log.md`. Bij het herontwerp op 25-08 bleek dat **10.401 van die producten (35%) op 1×1×1 mm staan** — een placeholder, geen meting. In het oude volumemodel betekende dat een vrijwel oneindige palletcapaciteit, dus de maximale consolidatiescore en een vulgraad van 0%: precies de artikelen die bovenaan de werklijst kwamen, waren de artikelen waarvan we niets wisten.
**Waarom:** De controle was `if (!length || !width || !height) skip` — die test alleen of het veld gevuld en niet-nul is. 1×1×1 komt daar zonder meer door. Er is nooit getoetst of de waarde *fysiek plausibel* was.
**Regel:** Bij het inlezen van referentiedata niet alleen op aanwezigheid controleren, maar op plausibiliteit: een ondergrens, een bovengrens, en expliciet op bekende placeholder-patronen (1×1×1, ronde standaardwaarden die overal identiek zijn — zoals de 100 mm-hoogtes bij de "-45"- en "-35"-locaties). Afgekeurde data levert "onbekend" op, en onbekend mag nooit meedoen in een rangschikking. Rapporteer het aantal afgekeurde records in de UI, zodat het niet stilzwijgend gebeurt.

## LESSON 3: Conclusies over het resultaat mogen niet uit de referentiedata alleen komen
**Wat ging fout:** Na het vinden van de 35% placeholder-afmetingen (LESSON 2) is als hoofdconclusie gepresenteerd dat "de bovenkant van de werklijst gevuld wordt door placeholder-data in plaats van echte kansen". Toen de echte voorraadexport er later diezelfde sessie was, bleek dat te grof: van de regels die daadwerkelijk in bulk liggen heeft maar **1,5% (96 regels, 49 artikelen)** onbetrouwbare afmetingen. De artikelen die in bulk staan zijn juist goed onderhouden. Er was wél een echt effect (onder het oude model zou "Reparatie overig" met 17 pallets op plek 1 zijn gekomen), maar veel smaller dan beweerd.
**Waarom:** Een percentage uit de referentiedata (de hele catalogus, 29.996 producten) is doorgetrokken naar een uitspraak over het resultaat (de ~2.000 artikelen die echt in bulk liggen). Dat zijn twee verschillende populaties, en de kwaliteit verschilt er sterk tussen.
**Regel:** Zolang er geen echte voorraadexport beschikbaar is, uitspraken over de referentiedata scheiden van uitspraken over de werklijst. Een percentage over de catalogus zegt niets over de verdeling in de voorraad — noem het dan als hypothese ("dit zou de lijst kunnen vervuilen"), niet als bevinding, en toets het zodra er een export is.

## LESSON 4: Na een push "zie ik de wijziging niet" op de live tool is meestal browsercache, geen codefout
**Wat ging fout:** Na het pushen van een nieuwe knop in `index.html` meldde de product owner "Ik zie geen knop". Eerste reflex zou zijn geweest om de code na te lopen op een bug.
**Waarom:** `curl` op de live GitHub Pages-URL bevestigde dat de knop al correct in de geserveerde HTML stond (Last-Modified kwam overeen met de laatste push) — er was dus geen codefout. GitHub Pages stuurt `Cache-Control: max-age=600` mee voor alle bestanden, en `index.html`/`scripts/script.js` hebben (anders dan `data/reference/*.xlsx`, die al een `?v=`-cache-buster hebben sinds een eerdere sessie) geen cache-buster. Een gewone refresh (F5) binnen die 10 minuten haalt daardoor de oude, lokaal gecachte versie op, ook al staat de server allang bij.
**Regel:** Als de product owner meldt dat een zojuist gepushte wijziging niet zichtbaar is op de live tool: eerst met `curl -I`/`curl` verifiëren dat GitHub Pages de nieuwe versie al serveert (vergelijk `Last-Modified`/bestandsgrootte met de laatste commit) vóórdat er naar een bug in de code gezocht wordt. Is de server al bij, vraag dan een harde refresh (Ctrl+Shift+R) i.p.v. een gewone refresh.
