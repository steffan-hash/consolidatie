/*
  script.js — Pallet Consolidatie

  Wat: Leest een .xlsx voorraadexport uit het WMS in, en filtert artikelen die
       op meerdere pallets op bulklocaties staan (kandidaten om te consolideren
       naar 1 pallet). Artikelen die al maar op 1 pallet staan worden er
       standaard uitgefilterd, want die hoeven niet aangepakt te worden.

  Input: Een .xlsx bestand, geüpload via de UI. Herkende kolommen (moeten
         aanwezig zijn, kolomvolgorde maakt niet uit): "Product",
         "Location Code", "Urn", "StockLocationTypeName". Overige kolommen
         worden 1-op-1 meegenomen.

         Omdat dit een handmatige export is, kan de koprij niet altijd op
         regel 1 staan (lege regels bovenaan) en kunnen er extra tabbladen
         bijzitten. Daarom wordt er over alle tabbladen en de eerste ~20
         regels heen gezocht naar de koprij.

  Output: Een nieuw .xlsx bestand dat gedownload wordt
          (rapportage_consolidatie_pallets_YYYY-MM-DD.xlsx). Kolomvolgorde:
          Location Code, Product Name, Quantity, Fill Rate, Action, To,
          Locations Freed, Remaining, Urn (zie "Exportkolommen" hieronder
          voor welke daarvan standaard wél/niet in het exportbestand zitten).
          Printopmaak staat op A4 liggend, geschaald naar 1 pagina breed,
          met Product Name zo breed als de andere kolommen toelaten.

  Doel (3.0): de tool moet vóór het piekseizoen zoveel mogelijk europallet-
          plekken vrijmaken. Vorig jaar moest er een extern magazijn met 7000
          gevulde europallet-plekken bijgehuurd worden omdat de goederen voor
          het nieuwe seizoen binnenkwamen terwijl er nog weinig verkocht werd.
          "Vrij te maken locaties" is daarom de kernmaat van deze tool.

  Capaciteit i.p.v. volume (3.0): een pallet wordt gestapeld, niet volgegoten.
          Daarom rekent de tool niet meer met volume-tegen-volume, maar met
          het aantal stuks dat er fysiek op past:

            stuks per laag  = beste van de twee liggingen op de palletvoetprint
                              (product staat rechtop, mag plat gedraaid worden)
            aantal lagen    = bruikbare stapelhoogte / producthoogte
            capaciteit      = stuks per laag × aantal lagen   (beide afgerond
                              naar beneden — je kunt een artikel niet halveren)

          Dit maakt een groot verschil. Voorbeeld: 3 stuks Eurom Flameheater
          (980×500×500 mm) op een 2 m hoge locatie is volgens het oude
          volumemodel 48% vol, terwijl er fysiek precies 3 stuks op passen —
          de pallet is dus 100% vol en biedt géén consolidatiekans.

          De palletvoetprint komt uit de kolom "Urn Type" in locations.xlsx
          (Euro Pallet, Blok Pallet, 180/270 Pallet), NIET uit de Length/Width
          van de locatie. Die kolommen bleken elkaar tegen te spreken (122
          "Euro Pallet"-locaties stonden als 1800 mm breed geregistreerd), en
          Urn Type is wel consistent. Zo is het model niet meer gevoelig voor
          precies de datafouten die eerder steeds handmatig gerepareerd zijn.

          Er wordt géén omdoos-marge meer afgetrokken: de afmetingen in
          products.xlsx zijn al die van de verkoopverpakking. Bevestigd door de
          product owner: verkopen ze iets als set, dan staat dat in de naam
          ("... - 8 stuks") en zijn de afmetingen die van de hele set. Quantity
          telt dus dezelfde eenheid als waar de afmetingen bij horen.

  Plausibiliteitstoets (3.0): referentiedata wordt eerst gewantrouwd. 10.401 van
          de 29.996 producten in products.xlsx staan op 1×1×1 mm — een
          placeholder, geen meting. Zonder toets krijgt zo'n product oneindig
          veel capaciteit, dus de maximale consolidatiescore, en komt het
          bovenaan de werklijst te staan. Producten/locaties die de toets niet
          halen worden "onbekend" en doen niet mee aan de rangschikking.

  Zelfcorrectie op wat we zien staan (3.0): staat er meer op een pallet dan
          volgens de berekening past, dan is de berekening fout (of er wordt
          hoger/ruimer gestapeld dan aangenomen) — niet de werkelijkheid. In
          dat geval wordt de waargenomen hoeveelheid de capaciteit. Dat is
          conservatief, en het aantal keer dat dit gebeurt is zichtbaar in de
          statistieken als signaal over de datakwaliteit.

  Werklijst voor de reachers (3.0): per artikel wordt bepaald hoeveel pallets
          er minimaal nodig zijn (vul de volste pallets, tel hun capaciteit op
          tot de totale hoeveelheid erin past). Het verschil met het huidige
          aantal pallets is kolom "Locations Freed". De pallets die
          overblijven krijgen actie "Empty" (leeghalen, minst erop = eerst),
          de rest "Keep". Kolom "Remaining" laat zien hoeveel stuks er nog
          bij kunnen — dat is het getal waar een chauffeur iets mee kan.

  Van→naar (3.0): kolom "To" wijst voor een "Empty"-pallet een concrete
          bestemming aan — maar alleen als de hele inhoud in één keer bij één
          "Keep"-pallet past (best fit: de ontvanger met de kleinste
          restruimte die nog wel groot genoeg is). Past een donorpallet
          nergens in zijn geheel, dan blijft hij "Empty" zonder doel ("-"),
          in plaats van een versnipperde instructie te tonen (voorraad over
          meerdere ontvangers verdelen). Bewuste keuze van de product owner
          na een proof of concept: bij grote artikelen leidde volledige
          consolidatie tot tientallen losse regels van een paar stuks — niet
          uitvoerbaar voor een chauffeur. Zie computeVanNaarMoves().

  Kolomvolgorde en -namen (3.0): op verzoek van de product owner staan de
          berekende kolommen in het Engels (Fill Rate, Action, To, Locations
          Freed, Remaining) — kort en duidelijk voor de reachtruck-
          chauffeurs, de rest van de tool (UI-teksten, statistieken) blijft
          Nederlands. Volgorde: Location Code, Product Name, Quantity, Fill
          Rate, Action, To, Locations Freed, Remaining, Urn — Urn staat
          bewust als allerlaatste kolom.

  Exportkolommen (3.0): de export paste niet meer op 1 A4 liggend met alle
          kolommen erbij. Urn, Fill Rate en Locations Freed staan daarom
          standaard NIET in het geëxporteerde bestand — wel altijd in de
          preview-tabel. Een toggle ("exportExtraColumns") zet ze er op
          verzoek weer bij (dan ook in dezelfde volgorde, Urn als laatste).

  Zoekfilter: een zoekveld boven de resultaattabel filtert op de zichtbare
          kolommen (Location Code, Product Name, Quantity, Urn). Werkt als
          laatste stap ná de score/actie-berekening, zodat zoeken naar 1
          pallet van een artikel niet de "vrij te maken locaties"/Legen-Keep
          bepaling voor dat artikel verstoort — die blijft naar alle pallets
          van het artikel kijken, ook de pallets die het zoekresultaat niet
          toont.

  Ruisreductie: (1) Producten met "DOOS", "BOX" of "TOP" als los woord in de
          naam (verpakkingsmateriaal) tellen nergens in mee. (2) Locaties
          waarvan de code "CHITA" bevat tellen nergens in mee — geen gewone
          bulklocaties in het magazijnrek. (3) Pallets waarop meerdere
          artikelen door elkaar staan (dezelfde Urn, verschillende producten)
          krijgen geen advies: de capaciteit is dan niet aan één artikel toe te
          rekenen. Alle uitsluitingen zijn zichtbaar in de statistieken, niet
          stilzwijgend.

          De eerdere regel "gelijkmatige stapeling" (artikel op >10 pallets met
          overal dezelfde hoeveelheid én vulgraad negeren) is in 3.0 vervallen.
          Die regel bestond alleen om artikelen te verbergen die het oude
          volumemodel onterecht als kans aanmerkte. Met een capaciteitsmodel
          valt een volle pallet vanzelf weg (0 vrij te maken locaties), dus de
          regel zou nu echte kansen gaan verbergen.
*/

(function () {
  const fileInput = document.getElementById('fileInput');
  const pickFileBtn = document.getElementById('pickFileBtn');
  const fileChip = document.getElementById('fileChip');
  const fileChipName = document.getElementById('fileChipName');
  const removeFileBtn = document.getElementById('removeFileBtn');
  const errorBox = document.getElementById('errorBox');
  const filterCard = document.getElementById('filterCard');
  const resultCard = document.getElementById('resultCard');
  const hideSinglePallet = document.getElementById('hideSinglePallet');
  const onlyOpportunities = document.getElementById('onlyOpportunities');
  const productGroupSection = document.getElementById('productGroupSection');
  const productGroupList = document.getElementById('productGroupList');
  const selectAllGroups = document.getElementById('selectAllGroups');
  const selectNoneGroups = document.getElementById('selectNoneGroups');
  const searchInput = document.getElementById('searchInput');
  const statsBox = document.getElementById('statsBox');
  const statusEl = document.getElementById('status');
  const previewTable = document.getElementById('previewTable');
  const exportBtn = document.getElementById('exportBtn');
  const exportExtraColumns = document.getElementById('exportExtraColumns');
  const refDataHint = document.getElementById('refDataHint');

  // Thema-knop rechtsboven: wisselt de "dark" class op <html> (waar
  // index.html al een eerste voorkeur voor heeft ingesteld vóór de pagina
  // rendert, om een lichtflits te voorkomen) en onthoudt de keuze.
  const themeToggle = document.getElementById('themeToggle');
  const iconSun = document.getElementById('iconSun');
  const iconMoon = document.getElementById('iconMoon');
  function syncThemeIcon() {
    const isDark = document.documentElement.classList.contains('dark');
    iconSun.classList.toggle('hidden', !isDark);
    iconMoon.classList.toggle('hidden', isDark);
  }
  syncThemeIcon();
  themeToggle.addEventListener('click', () => {
    const isDark = document.documentElement.classList.toggle('dark');
    document.documentElement.style.colorScheme = isDark ? 'dark' : 'light';
    localStorage.setItem('theme', isDark ? 'dark' : 'light');
    syncThemeIcon();
  });

  // Kolomnamen die we nodig hebben om het bestand te kunnen interpreteren.
  // Vergelijking gebeurt case-insensitive en na trimmen van spaties, zodat
  // kleine verschillen tussen exports geen probleem zijn.
  const STOCK_REQUIRED_HEADERS = ['product', 'location code', 'urn', 'stocklocationtypename'];
  const PALLET_LOCATION_TYPE = 'bulk location'; // exact deze waarde telt mee, "Bulk Location Extern" dus niet

  // Ruisreductie 1: producten met een van deze losse woorden in de naam
  // (Description) zijn verpakkingsmateriaal (dozen/omdoos-toppers), geen
  // fysiek te consolideren magazijnartikel — worden volledig genegeerd, ook
  // in de tellingen/statistieken. Hoofdletterongevoelig en op woordgrens,
  // dus "TOPPY" matcht niet op "TOP".
  const NOISE_PRODUCT_KEYWORDS = /\b(doos|box|top)\b/i;

  // Ruisreductie 3: locaties waarvan de code "CHITA" bevat (bijv. CHITA_AM1,
  // CHITA_DOOS10) zijn geen gewone bulklocaties in het magazijnrek, maar
  // apart benoemde plekken — worden op verzoek van de product owner volledig
  // genegeerd, net als ruisreductie 1.
  const NOISE_LOCATION_KEYWORD = /chita/i;

  // Referentiebestanden voor de vulgraadberekening (2.0) — vaste bestanden in
  // de repo, geen upload. Koppelveld voor producten is "Product ID" (komt
  // overeen met de "Product"-kolom in de voorraadexport), voor locaties is
  // dat "Location" (komt overeen met "Location Code").
  //
  // REF_DATA_VERSION als cache-buster: GitHub Pages laat browsers deze
  // bestanden 10 minuten cachen, dus zonder dit zou een refresh na het
  // bijwerken van products.xlsx/locations.xlsx alsnog de oude versie tonen.
  // Ophogen (bijv. datum) bij elke wijziging aan een van beide bestanden.
  const REF_DATA_VERSION = '2026-08-25b';
  const REF_PRODUCTS_URL = `data/reference/products.xlsx?v=${REF_DATA_VERSION}`;
  const REF_LOCATIONS_URL = `data/reference/locations.xlsx?v=${REF_DATA_VERSION}`;
  const PRODUCT_REF_HEADERS = ['product id', 'length', 'width', 'height'];
  const LOCATION_REF_HEADERS = ['location', 'height', 'urn type'];

  // De locatiehoogte uit het referentiebestand is de hoogte van de hele
  // opslaglocatie, maar de producten staan op een pallet die zelf ook hoogte
  // inneemt. Dit aantal mm gaat van de locatiehoogte af voordat de capaciteit
  // berekend wordt, want dat is geen ruimte voor het product zelf.
  const PALLET_HOOGTE_MM = 200;

  // In de praktijk wordt niet hoger gestapeld dan ongeveer 2 meter (opgave van
  // de product owner: een natuurlijke grens, geen regel uit het WMS). Locaties
  // die hoger zijn leveren dus geen extra capaciteit op. Dit is de totale
  // hoogte van de beladen pallet, dus inclusief PALLET_HOOGTE_MM.
  const MAX_STAPELHOOGTE_MM = 2000;

  // Palletvoetprint per "Urn Type" uit locations.xlsx, in mm. Dit is de
  // betrouwbare bron voor hoeveel vloeroppervlak er per pallet beschikbaar is
  // — de Length/Width-kolommen van de locatie spreken elkaar tegen (er staan
  // "Euro Pallet"-locaties op 1800 mm breed geregistreerd) terwijl Urn Type
  // consistent is. De twee maten zijn onderling verwisselbaar: bij het
  // berekenen van de stuks per laag proberen we het product in beide liggingen.
  // Locatiesoorten die hier niet in staan (bijv. "Hottub", die als 200000 mm
  // breed geregistreerd staat) leveren bewust "onbekend" op i.p.v. een gok.
  const PALLET_FOOTPRINTS = {
    'euro pallet': { a: 1200, b: 800 },
    'blok pallet': { a: 1200, b: 1000 },
    '180 pallet': { a: 1800, b: 1200 },
    '270 pallet': { a: 2700, b: 1200 },
  };

  // Plausibiliteitsgrenzen voor de referentiedata. Een derde van de producten
  // staat op 1×1×1 mm (placeholder); zonder deze toets krijgt zo'n product
  // een vrijwel oneindige capaciteit en dus de hoogste consolidatiescore.
  const PRODUCT_MIN_VOLUME_MM3 = 8000;  // kleiner dan een blokje van 2×2×2 cm bestaat niet als handelsartikel
  const PRODUCT_MAX_SIDE_MM = 2400;     // groter dan dit past nergens in het rek
  const LOCATION_MIN_HEIGHT_MM = 300;   // lager dan dit is geen opslaglocatie
  const LOCATION_MAX_HEIGHT_MM = 4000;  // hoger dan dit is een invoerfout

  // Overhang: voorraad mag over de rand van de pallet uitsteken. Zonder deze
  // regel zou een strikte pasrekening onnodig veel artikelen als "past niet"
  // wegzetten — gemeten op een echte export: 447 regels over 169 artikelen,
  // waaronder een hangstoel van 940×940 mm op een europallet van 1200×800 (die
  // staat er in de praktijk gewoon op) en een spacover van 2400×1210 mm op een
  // 270-pallet van 2700×1200 (10 mm te breed op papier).
  //
  // Regel: past het product niet netjes binnen de voetprint, maar is het
  // grondoppervlak niet groter dan deze factor × het palletoppervlak, dan gaan
  // we uit van 1 stuk per laag. Dat is de laagst mogelijke aanname, dus het kan
  // de capaciteit alleen onderschatten — nooit een valse consolidatiekans
  // opleveren. Echt buitenmaatse artikelen blijven "past niet op pallet".
  const OVERHANG_MAX_AREA_FACTOR = 2;

  // Vaste set kolommen voor het resultaat — niet instelbaar in de UI. De rest
  // van de brondata is ruis voor het consolideren van pallets. Urn staat
  // hier los van (zie URN_COLUMN): die staat op verzoek van de product owner
  // altijd als laatste kolom, ná de berekende kolommen.
  const OUTPUT_COLUMNS = [
    { key: 'location code', label: 'Location Code' },
    { key: 'description', label: 'Product Name' },
    { key: 'quantity', label: 'Quantity' },
  ];
  const URN_COLUMN = { key: 'urn', label: 'Urn' };

  let originalHeaders = [];   // koprij in originele volgorde, zoals in het bestand
  let outputColumns = [];     // OUTPUT_COLUMNS aangevuld met de bijbehorende originele headernaam
  let urnColumn = null;       // URN_COLUMN aangevuld met de bijbehorende originele headernaam
  let baseRows = [];          // rijen op een "Bulk Location", als objecten {header: waarde}
  let palletCountByProduct = new Map(); // Product -> aantal unieke pallets (Urn's)
  let resultRows = [];        // rijen die momenteel getoond/geëxporteerd worden
  let productGroupHeader = null;        // originele headernaam van "Product Group", of null als kolom ontbreekt
  let selectedProductGroups = new Set(); // welke Product Group waarden momenteel getoond worden

  let productDimsById = new Map();    // genormaliseerde Product ID -> {length,width,height} in mm, alleen plausibele
  let locationInfoByCode = new Map(); // genormaliseerde Location Code -> {height, urnType, footprint}
  let referenceDataReady = false;     // true zodra beide referentiebestanden geladen en gekoppeld zijn

  let totalQtyByProduct = new Map();    // Product -> totale hoeveelheid over al zijn Bulk Location-pallets
  let scoreByProduct = new Map();       // Product -> {minPalletsNeeded, locationsFreed}, zie computeConsolidationScores
  let mixedPalletUrns = new Set();      // Urn's waarop meerdere verschillende artikelen staan

  let noiseExcludedRowCount = 0;         // aantal regels genegeerd door NOISE_PRODUCT_KEYWORDS (verpakkingsmateriaal)
  let chitaExcludedRowCount = 0;         // aantal regels genegeerd door NOISE_LOCATION_KEYWORD (CHITA-locaties)
  let observedOverrideCount = 0;         // aantal pallets waar meer op staat dan berekend past (zelfcorrectie)
  let overhangAssumedCount = 0;          // aantal pallets waarbij overhang is aangenomen (1 stuk per laag)
  let refStats = null;                   // telling van afgekeurde referentiedata, voor de statistieken

  function norm(v) {
    return String(v ?? '').trim();
  }
  function normKey(v) {
    return norm(v).toLowerCase();
  }

  function showError(msg) {
    errorBox.textContent = msg;
    errorBox.style.display = 'block';
  }
  function clearError() {
    errorBox.style.display = 'none';
    errorBox.textContent = '';
  }

  // Zoekt, over alle tabbladen en de eerste 20 regels van elk tabblad, naar de
  // rij die het meest lijkt op de koprij (de meeste verwachte kolomnamen
  // erin). Nodig omdat handmatige exports soms lege regels bovenaan hebben
  // of meerdere tabbladen bevatten. requiredHeaders is generiek zodat dit
  // ook voor de referentiebestanden (producten/locaties) gebruikt kan worden.
  function findHeaderRow(workbook, requiredHeaders) {
    let best = null; // { sheetName, rowIndex, headerRow, matchCount }

    for (const sheetName of workbook.SheetNames) {
      const ws = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true, blankrows: false });
      const scanLimit = Math.min(rows.length, 20);

      for (let r = 0; r < scanLimit; r++) {
        const row = rows[r];
        const normalizedCells = row.map(normKey);
        const matchCount = requiredHeaders.filter(h => normalizedCells.includes(h)).length;
        if (!best || matchCount > best.matchCount) {
          best = { sheetName, rowIndex: r, headerRow: row, matchCount, allRows: rows };
        }
      }
    }
    return best;
  }

  // Leest één referentiebestand (producten of locaties) en zet elke datarij
  // om naar een object {header: waarde}, net als bij de voorraadexport.
  async function loadRefFile(url, requiredHeaders) {
    const res = await fetch(url);
    if (!res.ok) throw new Error(`${url} gaf status ${res.status}`);
    const data = await res.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });
    const best = findHeaderRow(workbook, requiredHeaders);
    if (!best || best.matchCount < requiredHeaders.length) {
      throw new Error(`${url}: verwachte kolommen niet gevonden`);
    }
    const rows = [];
    for (let r = best.rowIndex + 1; r < best.allRows.length; r++) {
      const row = best.allRows[r];
      const rowObj = {};
      best.headerRow.forEach((cell, i) => {
        const h = normKey(cell);
        if (h !== '') rowObj[h] = row[i] !== undefined ? row[i] : '';
      });
      rows.push(rowObj);
    }
    return rows;
  }

  // Bouwt de productkaart: Product ID -> afmetingen in mm. Alleen producten
  // die de plausibiliteitstoets halen komen erin. Dat is essentieel: een derde
  // van products.xlsx staat op 1×1×1 mm (placeholder), en zo'n product zou
  // anders een vrijwel oneindige palletcapaciteit krijgen en daarmee de
  // hoogste consolidatiescore van allemaal.
  function buildProductDims(rows) {
    const map = new Map();
    const rejected = { leeg: 0, placeholder: 0, teKlein: 0, teGroot: 0 };
    rows.forEach(row => {
      const key = normKey(row['product id']);
      if (key === '') return;
      const length = Number(row['length']);
      const width = Number(row['width']);
      const height = Number(row['height']);
      if (!length || !width || !height || length < 0 || width < 0 || height < 0) { rejected.leeg++; return; }
      if (length === 1 && width === 1 && height === 1) { rejected.placeholder++; return; }
      if (length * width * height < PRODUCT_MIN_VOLUME_MM3) { rejected.teKlein++; return; }
      if (length > PRODUCT_MAX_SIDE_MM || width > PRODUCT_MAX_SIDE_MM || height > PRODUCT_MAX_SIDE_MM) { rejected.teGroot++; return; }
      map.set(key, { length, width, height });
    });
    return { map, rejected, total: rows.length };
  }

  // Bouwt de locatiekaart: Location -> {hoogte, palletsoort, voetprint}. De
  // Length/Width van de locatie worden bewust NIET gebruikt (die spreken elkaar
  // tegen) — de voetprint komt uit Urn Type. Alleen de hoogte komt uit de
  // locatie zelf, want die bepaalt hoeveel lagen er op kunnen.
  function buildLocationInfo(rows) {
    const map = new Map();
    const rejected = { leeg: 0, hoogteOnwaarschijnlijk: 0, palletsoortOnbekend: 0 };
    rows.forEach(row => {
      const key = normKey(row['location']);
      if (key === '') return;
      const height = Number(row['height']);
      if (!height || height < 0) { rejected.leeg++; return; }
      if (height < LOCATION_MIN_HEIGHT_MM || height > LOCATION_MAX_HEIGHT_MM) { rejected.hoogteOnwaarschijnlijk++; return; }
      const urnType = normKey(row['urn type']);
      const footprint = PALLET_FOOTPRINTS[urnType];
      if (!footprint) { rejected.palletsoortOnbekend++; return; }
      map.set(key, { height, urnType, footprint });
    });
    return { map, rejected, total: rows.length };
  }

  // Laadt de referentiebestanden op (data/reference/, geen upload). Lukt dit
  // niet — bijv. omdat index.html lokaal via dubbelklikken geopend is i.p.v.
  // via een webserver/GitHub Pages — dan werkt de rest van de tool gewoon
  // door, alleen zonder vulgraadberekening.
  async function loadReferenceData() {
    try {
      const [products, locations] = await Promise.all([
        loadRefFile(REF_PRODUCTS_URL, PRODUCT_REF_HEADERS),
        loadRefFile(REF_LOCATIONS_URL, LOCATION_REF_HEADERS),
      ]);
      const prodResult = buildProductDims(products);
      const locResult = buildLocationInfo(locations);
      productDimsById = prodResult.map;
      locationInfoByCode = locResult.map;
      refStats = { products: prodResult, locations: locResult };
      referenceDataReady = true;
      // Zichtbaar in de browserconsole (F12) zodat te zien is hoeveel van de
      // referentiebestanden de plausibiliteitstoets haalt, en waarom niet.
      console.info(
        `Referentiedata geladen. Producten bruikbaar: ${productDimsById.size}/${products.length} ` +
        `(afgekeurd: ${prodResult.rejected.placeholder} placeholder 1x1x1, ${prodResult.rejected.leeg} leeg, ` +
        `${prodResult.rejected.teKlein} te klein, ${prodResult.rejected.teGroot} te groot). ` +
        `Locaties bruikbaar: ${locationInfoByCode.size}/${locations.length} ` +
        `(afgekeurd: ${locResult.rejected.palletsoortOnbekend} palletsoort onbekend, ` +
        `${locResult.rejected.hoogteOnwaarschijnlijk} onwaarschijnlijke hoogte, ${locResult.rejected.leeg} leeg).`
      );
    } catch (e) {
      console.warn('Referentiedata (locatie-/productafmetingen) kon niet geladen worden — vulgraad wordt niet berekend.', e);
      referenceDataReady = false;
    }
    refDataHint.style.display = referenceDataReady ? 'none' : 'block';
    if (baseRows.length) applyFilter(); // alsnog herberekenen als data na de eerste render binnenkomt
  }
  loadReferenceData();

  // Hoeveel stuks passen er naast elkaar in één laag op de palletvoetprint?
  // Het product staat rechtop (hoogte is vast), maar mag in het platte vlak
  // een kwartslag gedraaid worden — we nemen de beste van die twee liggingen.
  // Afronden naar beneden, want een half artikel bestaat niet.
  //
  // Past het niet netjes binnen de voetprint, dan wordt overhang toegestaan
  // (zie OVERHANG_MAX_AREA_FACTOR): 1 stuk per laag, mits het artikel niet
  // buitenmaats is. Geeft terug of die aanname gebruikt is, zodat het aantal
  // zichtbaar te maken is in de statistieken.
  function unitsPerLayer(productLength, productWidth, footprint) {
    const opstelling1 = Math.floor(footprint.a / productLength) * Math.floor(footprint.b / productWidth);
    const opstelling2 = Math.floor(footprint.a / productWidth) * Math.floor(footprint.b / productLength);
    const netjes = Math.max(opstelling1, opstelling2);
    if (netjes > 0) return { perLaag: netjes, overhang: false };

    const productOppervlak = productLength * productWidth;
    const palletOppervlak = footprint.a * footprint.b;
    if (productOppervlak <= palletOppervlak * OVERHANG_MAX_AREA_FACTOR) {
      return { perLaag: 1, overhang: true };
    }
    return { perLaag: 0, overhang: false };
  }

  // Capaciteit van 1 pallet-regel: hoeveel stuks van dit artikel passen er
  // fysiek op deze pallet op deze locatie? Dit vervangt de oude volumeratio.
  // Geeft altijd een reden mee als het niet lukt, zodat in de statistieken
  // zichtbaar is WAAROM iets onbekend is i.p.v. dat er iets geraden wordt.
  function computeCapacity(row, productHeader, locationHeader, quantityHeader, urnHeader, urnTypeHeader) {
    if (!referenceDataReady) return { reason: 'referentiedata-niet-geladen' };

    const qty = Number(row[quantityHeader]);
    if (!qty || qty <= 0) return { reason: 'aantal-ongeldig' };

    // Staan er meerdere artikelen op dezelfde pallet, dan is de capaciteit
    // niet aan één artikel toe te rekenen — dan geven we geen advies.
    if (mixedPalletUrns.has(norm(row[urnHeader]))) return { reason: 'gemengde-pallet' };

    const product = productDimsById.get(normKey(row[productHeader]));
    if (!product) return { reason: 'product-afmetingen-onbetrouwbaar' };

    const location = locationInfoByCode.get(normKey(row[locationHeader]));
    if (!location) return { reason: 'locatie-onbruikbaar' };

    // Palletsoort: de voorraadexport heeft zelf een "Urn Type"-kolom, en die
    // zegt welke pallet er WERKELIJK staat — terwijl locations.xlsx zegt waar
    // de locatie voor bedoeld is. Die twee verschillen in de praktijk (gemeten:
    // 512 van 6310 regels, o.a. blokpallets op europallet-plekken), dus de
    // export gaat voor. Ontbreekt de kolom of is de waarde onbekend, dan
    // vallen we terug op de locatie.
    let footprint = location.footprint;
    if (urnTypeHeader !== undefined) {
      const exportFootprint = PALLET_FOOTPRINTS[normKey(row[urnTypeHeader])];
      if (exportFootprint) footprint = exportFootprint;
    }

    // Stapelhoogte: de laagste van de locatiehoogte en de praktijkgrens van
    // ~2 m, min de hoogte die de pallet zelf inneemt.
    const stapelHoogte = Math.min(location.height, MAX_STAPELHOOGTE_MM) - PALLET_HOOGTE_MM;
    if (stapelHoogte <= 0) return { reason: 'locatie-te-laag' };

    const laag = unitsPerLayer(product.length, product.width, footprint);
    if (laag.perLaag <= 0) return { reason: 'past-niet-op-pallet' };
    if (laag.overhang) overhangAssumedCount++;

    // Bij overhang (het artikel steekt over de rand) nemen we maar 1 laag aan,
    // ongeacht hoeveel er volgens de hoogte-deling zou passen. Overstekende
    // artikelen meerdere lagen hoog stapelen is niet stabiel — bevestigd door
    // de product owner na het hangstoel-voorbeeld (5-6 lagen hoog gestapelde
    // overhangende stoelen is niet realistisch). Dit is dus geen rekenkundige
    // aanname maar een expliciete praktijkregel.
    const lagen = laag.overhang
      ? (product.height <= stapelHoogte ? 1 : 0)
      : Math.floor(stapelHoogte / product.height);
    if (lagen <= 0) return { reason: 'product-te-hoog' };

    const berekend = laag.perLaag * lagen;

    // Zelfcorrectie: staat er méér op dan we berekenden, dan is de berekening
    // fout (of er wordt hoger/ruimer gestapeld dan aangenomen) — niet de
    // werkelijkheid. De waarneming wint, want die staat er fysiek.
    const capacity = Math.max(berekend, qty);
    if (capacity > berekend) observedOverrideCount++;

    return {
      reason: 'ok',
      capacity,
      berekend,
      qty,
      perLaag: laag.perLaag,
      overhang: laag.overhang,
      lagen,
      restruimte: capacity - qty,
      ratio: qty / capacity,
    };
  }

  // Consolidatiepotentieel: hoeveel pallets heeft een artikel minimaal nodig?
  // We vullen de volste pallets eerst en tellen hun capaciteit op tot de
  // totale hoeveelheid erin past. Het aantal pallets dat dan overblijft is het
  // aantal locaties dat daadwerkelijk vrijgemaakt kan worden — de kernmaat van
  // deze tool (zie doel bovenaan: europallet-plekken vrijmaken vóór de piek).
  //
  // Vereist dat __cap al op alle baseRows staat (zie applyFilter). Artikelen
  // waarvan ook maar één pallet een onbekende capaciteit heeft krijgen géén
  // score: dan is niet te zeggen of het echt in minder pallets past.
  function computeConsolidationScores(productHeader, urnHeader) {
    const scores = new Map();
    if (!referenceDataReady) return scores;

    const rowsByProduct = new Map();
    baseRows.forEach(row => {
      const product = norm(row[productHeader]);
      if (!rowsByProduct.has(product)) rowsByProduct.set(product, []);
      rowsByProduct.get(product).push(row);
    });

    rowsByProduct.forEach((rows, product) => {
      if (!rows.every(r => r.__cap && r.__cap.reason === 'ok')) return;

      const totalQty = totalQtyByProduct.get(product) || 0;
      if (totalQty <= 0) return;

      // Per pallet (Urn) rekenen, niet per regel: staat hetzelfde artikel met
      // twee regels op dezelfde pallet, dan is dat één pallet-plek.
      const pallets = groupByUrn(rows, urnHeader);

      // Grootste capaciteiten eerst: zo hebben we de minste pallets nodig.
      const capacities = Array.from(pallets.values()).map(p => p.capacity).sort((a, b) => b - a);
      let verzameld = 0;
      let minPalletsNeeded = 0;
      for (const cap of capacities) {
        if (verzameld >= totalQty) break;
        verzameld += cap;
        minPalletsNeeded++;
      }
      minPalletsNeeded = Math.max(1, minPalletsNeeded);

      const locationsFreed = Math.max(0, pallets.size - minPalletsNeeded);
      scores.set(product, { minPalletsNeeded, locationsFreed, currentPallets: pallets.size });
    });

    return scores;
  }

  // Bundelt regels van hetzelfde artikel per pallet (Urn): hoeveelheid bij
  // elkaar optellen, capaciteit één keer meetellen (die hoort bij de locatie,
  // niet bij de regel).
  function groupByUrn(rows, urnHeader) {
    const pallets = new Map();
    rows.forEach(row => {
      const urn = norm(row[urnHeader]);
      if (!pallets.has(urn)) {
        pallets.set(urn, { urn, qty: 0, capacity: row.__cap.capacity, rows: [] });
      }
      const pallet = pallets.get(urn);
      pallet.qty += row.__cap.qty;
      pallet.rows.push(row);
    });
    return pallets;
  }

  // Actie per pallet: welke pallets moeten leeg (Empty) en welke blijven staan
  // (Keep)? We houden de pallets met het MEESTE erop aan — dat kost de minste
  // ritten — en vullen die aan tot de totale hoeveelheid erin past. Wat dan
  // overblijft kan leeg. Vereist dat __cap en __score al gezet zijn.
  function computeConsolidationActions(rows, productHeader, urnHeader) {
    const rowsByProduct = new Map();
    rows.forEach(row => {
      const product = norm(row[productHeader]);
      if (!rowsByProduct.has(product)) rowsByProduct.set(product, []);
      rowsByProduct.get(product).push(row);
    });

    rowsByProduct.forEach((productRows, product) => {
      const score = scoreByProduct.get(product);
      if (!score) {
        productRows.forEach(row => { row.__action = 'onbekend'; });
        return;
      }
      if (score.locationsFreed <= 0) {
        productRows.forEach(row => { row.__action = null; }); // vol genoeg, geen kans
        return;
      }

      const totalQty = totalQtyByProduct.get(product) || 0;

      // Volste pallets eerst behouden: dat kost de minste ritten, en voor een
      // chauffeur is "de bijna-lege plekken worden geruimd" de logische regel.
      const pallets = Array.from(groupByUrn(productRows, urnHeader).values())
        .sort((a, b) => b.qty - a.qty);
      const teBehouden = new Set();
      let verzameld = 0;
      for (const pallet of pallets) {
        if (verzameld >= totalQty) break;
        verzameld += pallet.capacity;
        teBehouden.add(pallet);
      }
      // Altijd minstens 1 pallet laten staan om de voorraad op te ontvangen.
      if (teBehouden.size === 0 && pallets.length) teBehouden.add(pallets[0]);

      pallets.forEach(pallet => {
        const actie = teBehouden.has(pallet) ? 'behouden' : 'legen';
        pallet.rows.forEach(row => { row.__action = actie; });
      });
    });
  }

  // Van→naar: voor elke "Empty"-pallet (legen) kijken of zijn hele inhoud in
  // één keer bij één "Keep"-pallet (behouden) past — dan is dat een schone,
  // in één beweging uit te voeren instructie. Past een donorpallet nergens in
  // zijn geheel, dan blijft hij gewoon "Empty" zonder concreet doel: op
  // verzoek van de product owner tonen we bewust geen versnipperde
  // instructies (een pallet over meerdere ontvangers verdelen). Proof of
  // concept op de echte export liet zien dat dat voor grote artikelen tot
  // tientallen losse regels van een paar stuks leidt — niet uitvoerbaar.
  //
  // Donoren van groot naar klein verwerkt, en per donor de ontvanger met de
  // KLEINSTE restruimte die nog wel groot genoeg is (best fit) — dat
  // voorkomt dat een kleine donor een grote, later hard nodige ontvanger
  // opsoupeert. Vereist dat __action al gezet is (zie computeConsolidationActions).
  function computeVanNaarMoves(rows, productHeader, urnHeader, locationHeader) {
    const rowsByProduct = new Map();
    rows.forEach(row => {
      row.__moveTo = null; // opnieuw bepalen bij elke herberekening, geen oude waarde laten hangen
      const product = norm(row[productHeader]);
      if (!rowsByProduct.has(product)) rowsByProduct.set(product, []);
      rowsByProduct.get(product).push(row);
    });

    rowsByProduct.forEach((productRows, product) => {
      const receivers = Array.from(groupByUrn(productRows.filter(r => r.__action === 'behouden'), urnHeader).values())
        .map(p => ({ ...p, vrij: p.capacity - p.qty }));
      const donors = Array.from(groupByUrn(productRows.filter(r => r.__action === 'legen'), urnHeader).values())
        .sort((a, b) => b.qty - a.qty);
      if (!donors.length || !receivers.length) return;

      donors.forEach(donor => {
        const bestFit = receivers
          .filter(r => r.vrij >= donor.qty)
          .sort((a, b) => a.vrij - b.vrij)[0];
        if (!bestFit) return; // geen enkele ontvanger heeft in 1 keer plek — blijft "Empty" zonder doel

        bestFit.vrij -= donor.qty;
        const targetLocation = norm(bestFit.rows[0][locationHeader]);
        donor.rows.forEach(row => { row.__moveTo = { location: targetLocation, urn: bestFit.urn }; });
      });
    });
  }

  function resetUI() {
    clearError();
    filterCard.style.display = 'none';
    resultCard.style.display = 'none';
    statsBox.innerHTML = '';
    previewTable.innerHTML = '';
    baseRows = [];
    resultRows = [];
    outputColumns = [];
    urnColumn = null;
    palletCountByProduct = new Map();
    totalQtyByProduct = new Map();
    scoreByProduct = new Map();
    mixedPalletUrns = new Set();
    noiseExcludedRowCount = 0;
    chitaExcludedRowCount = 0;
    observedOverrideCount = 0;
    productGroupHeader = null;
    selectedProductGroups = new Set();
    productGroupSection.style.display = 'none';
    productGroupList.innerHTML = '';
    searchInput.value = '';
  }

  pickFileBtn.addEventListener('click', () => fileInput.click());

  removeFileBtn.addEventListener('click', () => {
    fileInput.value = '';
    fileChip.style.display = 'none';
    resetUI();
  });

  function formatFileSize(bytes) {
    return bytes < 1024 * 1024
      ? `${(bytes / 1024).toFixed(0)} KB`
      : `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }

  fileInput.addEventListener('change', async (ev) => {
    resetUI();
    const file = ev.target.files[0];
    if (!file) { fileChip.style.display = 'none'; return; }

    fileChipName.textContent = `${file.name} (${formatFileSize(file.size)})`;
    fileChip.style.display = 'flex';

    let workbook;
    try {
      const data = await file.arrayBuffer();
      workbook = XLSX.read(data, { type: 'array' });
    } catch (e) {
      showError('Kon dit bestand niet lezen als Excel-bestand. Is het een geldig .xlsx bestand?');
      return;
    }

    const best = findHeaderRow(workbook, STOCK_REQUIRED_HEADERS);
    if (!best || best.matchCount < STOCK_REQUIRED_HEADERS.length) {
      showError(
        'Kon de verwachte kolommen niet vinden (Product, Location Code, Urn, ' +
        'StockLocationTypeName). Controleer of dit een voorraad-export uit het WMS is.'
      );
      return;
    }

    // Koprij en kolom-index bepalen op basis van de gevonden rij.
    originalHeaders = best.headerRow.map(norm).filter(h => h !== '');
    outputColumns = OUTPUT_COLUMNS.map(c => ({
      label: c.label,
      header: originalHeaders.find(h => normKey(h) === c.key), // undefined als kolom ontbreekt in dit bestand
    }));
    urnColumn = {
      label: URN_COLUMN.label,
      header: originalHeaders.find(h => normKey(h) === URN_COLUMN.key),
    };
    const headerRowNormalized = best.headerRow.map(normKey);
    const colIndex = {};
    headerRowNormalized.forEach((h, i) => { if (h !== '') colIndex[h] = i; });

    const idxLocType = colIndex['stocklocationtypename'];
    const idxUrn = colIndex['urn'];
    const idxDescription = colIndex['description'];
    const idxLocationCode = colIndex['location code'];

    // Alle datarijen na de koprij omzetten naar objecten, en meteen filteren
    // op StockLocationTypeName === "Bulk Location". Voorraad op pick-locaties
    // of externe bulklocaties telt niet mee voor consolidatie: die pallets
    // kunnen we niet fysiek samenvoegen met de rest van het magazijn.
    // Ook verpakkingsmateriaal (NOISE_PRODUCT_KEYWORDS) en CHITA-locaties
    // (NOISE_LOCATION_KEYWORD) worden hier al uitgesloten — dat is ruis, geen
    // fysiek te consolideren artikel/locatie.
    baseRows = [];
    noiseExcludedRowCount = 0;
    chitaExcludedRowCount = 0;
    for (let r = best.rowIndex + 1; r < best.allRows.length; r++) {
      const row = best.allRows[r];
      if (normKey(row[idxLocType]) !== PALLET_LOCATION_TYPE) continue;
      if (norm(row[idxUrn]) === '') continue; // geen pallet-ID: kan niet meegeteld worden
      if (idxDescription !== undefined && NOISE_PRODUCT_KEYWORDS.test(norm(row[idxDescription]))) {
        noiseExcludedRowCount++;
        continue;
      }
      if (idxLocationCode !== undefined && NOISE_LOCATION_KEYWORD.test(norm(row[idxLocationCode]))) {
        chitaExcludedRowCount++;
        continue;
      }

      // Bouw het rij-object op basis van kolomvolgorde uit de koprij.
      const rowObj = {};
      best.headerRow.forEach((cell, i) => {
        const h = norm(cell);
        if (h !== '') rowObj[h] = row[i] !== undefined ? row[i] : '';
      });
      baseRows.push(rowObj);
    }

    if (!baseRows.length) {
      showError('Geen regels gevonden met StockLocationTypeName "Bulk Location" in dit bestand.');
      return;
    }

    // Per product het aantal unieke pallets (Urn's) op bulklocaties tellen,
    // plus de totale hoeveelheid — nodig om het consolidatiepotentieel te
    // berekenen. Tegelijk bijhouden welke pallets gemengd zijn (meerdere
    // artikelen op dezelfde Urn): daarvan is de capaciteit niet aan één
    // artikel toe te rekenen, dus die krijgen geen advies.
    const productHeader = originalHeaders.find(h => normKey(h) === 'product');
    const urnHeader = originalHeaders.find(h => normKey(h) === 'urn');
    const quantityHeaderForTotals = originalHeaders.find(h => normKey(h) === 'quantity');
    const palletSetsByProduct = new Map();
    const productsByUrn = new Map();
    totalQtyByProduct = new Map();
    baseRows.forEach(row => {
      const product = norm(row[productHeader]);
      const urn = norm(row[urnHeader]);
      if (!palletSetsByProduct.has(product)) palletSetsByProduct.set(product, new Set());
      palletSetsByProduct.get(product).add(urn);

      if (!productsByUrn.has(urn)) productsByUrn.set(urn, new Set());
      productsByUrn.get(urn).add(product);

      totalQtyByProduct.set(product, (totalQtyByProduct.get(product) || 0) + (Number(row[quantityHeaderForTotals]) || 0));
    });
    palletCountByProduct = new Map();
    palletSetsByProduct.forEach((set, product) => palletCountByProduct.set(product, set.size));
    mixedPalletUrns = new Set();
    productsByUrn.forEach((products, urn) => { if (products.size > 1) mixedPalletUrns.add(urn); });

    // Product Group filter opbouwen: alleen tonen als de kolom aanwezig is in
    // dit bestand. Waarden komen uit de data zelf, dus dit past zich vanzelf
    // aan per export (geen hardcoded lijst van groepen).
    productGroupHeader = originalHeaders.find(h => normKey(h) === 'product group');
    if (productGroupHeader) {
      const groups = Array.from(new Set(baseRows.map(r => norm(r[productGroupHeader])).filter(g => g !== ''))).sort();
      selectedProductGroups = new Set(groups); // standaard: alles aan, dus geen filtering
      renderProductGroupList(groups);
      productGroupSection.style.display = 'block';
    }

    filterCard.style.display = 'block';
    resultCard.style.display = 'block';
    applyFilter();
  });

  function renderProductGroupList(groups) {
    productGroupList.innerHTML = groups.map(g => {
      const id = 'pg_' + g.replace(/[^a-z0-9]/gi, '_');
      return `<label class="flex items-center gap-1.5 text-sm font-medium cursor-pointer">` +
        `<input type="checkbox" class="pgCheckbox w-3.5 h-3.5 rounded border-zinc-300 dark:border-zinc-600 accent-[#eab627]" value="${g}" id="${id}" checked> ${g}</label>`;
    }).join('');
    productGroupList.querySelectorAll('.pgCheckbox').forEach(cb => {
      cb.addEventListener('change', () => {
        if (cb.checked) selectedProductGroups.add(cb.value);
        else selectedProductGroups.delete(cb.value);
        applyFilter();
      });
    });
  }

  selectAllGroups.addEventListener('click', () => {
    productGroupList.querySelectorAll('.pgCheckbox').forEach(cb => { cb.checked = true; selectedProductGroups.add(cb.value); });
    applyFilter();
  });
  selectNoneGroups.addEventListener('click', () => {
    productGroupList.querySelectorAll('.pgCheckbox').forEach(cb => { cb.checked = false; });
    selectedProductGroups.clear();
    applyFilter();
  });

  // Past alle filters toe op de al-op-Bulk-Location gefilterde basisdata
  // (handmatig via de UI, én de automatische ruisreductie), en ververst de
  // weergave.
  function applyFilter() {
    const productHeader = originalHeaders.find(h => normKey(h) === 'product');
    const locationHeader = originalHeaders.find(h => normKey(h) === 'location code');
    const quantityHeader = originalHeaders.find(h => normKey(h) === 'quantity');
    const urnHeader = originalHeaders.find(h => normKey(h) === 'urn');
    const urnTypeHeader = originalHeaders.find(h => normKey(h) === 'urn type'); // optioneel

    // Capaciteit vooraf voor ALLE basisregels berekenen (niet pas na filteren)
    // — de consolidatiescore moet naar alle pallets van een artikel kunnen
    // kijken, los van de huidige UI-filters.
    observedOverrideCount = 0;
    overhangAssumedCount = 0;
    baseRows.forEach(row => {
      row.__cap = computeCapacity(row, productHeader, locationHeader, quantityHeader, urnHeader, urnTypeHeader);
    });

    scoreByProduct = computeConsolidationScores(productHeader, urnHeader);

    resultRows = baseRows.filter(row => {
      const product = norm(row[productHeader]);
      if (hideSinglePallet.checked) {
        if ((palletCountByProduct.get(product) || 0) < 2) return false;
      }
      // Alleen echte kansen tonen: artikelen waar consolideren daadwerkelijk
      // een locatie vrijmaakt. Artikelen met een onbekende capaciteit vallen
      // hier ook weg, want daarvan is geen kans aan te tonen.
      if (onlyOpportunities.checked) {
        const score = scoreByProduct.get(product);
        if (!score || score.locationsFreed <= 0) return false;
      }
      if (productGroupHeader && !selectedProductGroups.has(norm(row[productGroupHeader]))) return false;
      return true;
    });

    // Score en actie per regel bepalen vóór het sorteren, want de actie is een
    // sorteersleutel: in een werklijst horen de pallets die leeg moeten
    // bovenaan te staan, niet verspreid tussen de blijvers.
    resultRows.forEach(row => {
      row.__score = scoreByProduct.get(norm(row[productHeader])) || null;
    });
    computeConsolidationActions(resultRows, productHeader, urnHeader);
    computeVanNaarMoves(resultRows, productHeader, urnHeader, locationHeader);

    // Sorteren op consolidatiewinst (meeste vrij te maken locaties bovenaan),
    // dan op product (zodat alle pallets van hetzelfde artikel bij elkaar
    // staan), dan op actie (Empty vóór Keep — dat is de daadwerkelijke
    // werkvoorraad), en pas daarna op hoeveelheid en locatie voor een
    // voorspelbare volgorde. Artikelen zonder bekende score (-1) staan onderaan.
    const ACTION_ORDER = { legen: 0, behouden: 1, onbekend: 2 };
    resultRows = resultRows.slice().sort((a, b) => {
      const pa = norm(a[productHeader]), pb = norm(b[productHeader]);
      const freedA = scoreByProduct.has(pa) ? scoreByProduct.get(pa).locationsFreed : -1;
      const freedB = scoreByProduct.has(pb) ? scoreByProduct.get(pb).locationsFreed : -1;
      if (freedA !== freedB) return freedB - freedA;
      if (pa !== pb) return pa < pb ? -1 : 1;
      const actA = ACTION_ORDER[a.__action] !== undefined ? ACTION_ORDER[a.__action] : 3;
      const actB = ACTION_ORDER[b.__action] !== undefined ? ACTION_ORDER[b.__action] : 3;
      if (actA !== actB) return actA - actB;
      const qa = a.__cap && a.__cap.reason === 'ok' ? a.__cap.qty : Infinity;
      const qb = b.__cap && b.__cap.reason === 'ok' ? b.__cap.qty : Infinity;
      if (qa !== qb) return qa - qb;
      const la = norm(a[locationHeader]), lb = norm(b[locationHeader]);
      return la < lb ? -1 : la > lb ? 1 : 0;
    });

    // Zoekfilter als laatste stap: zoekt in exact de kolommen die ook in de
    // tabel te zien zijn (Location Code, Product Name, Quantity, Urn). Pas
    // NA de score/actie-berekening toegepast, want die moet naar alle
    // pallets van een artikel kunnen kijken — anders zou zoeken op 1 pallet
    // van een artikel de "vrij te maken locaties"/Legen-Keep-berekening voor
    // dat artikel verstoren.
    const searchTerm = normKey(searchInput.value);
    if (searchTerm) {
      const searchableColumns = urnColumn ? outputColumns.concat([urnColumn]) : outputColumns;
      resultRows = resultRows.filter(row =>
        searchableColumns.some(c => c.header && normKey(row[c.header]).includes(searchTerm))
      );
    }

    renderStats();
    renderPreview();
  }
  hideSinglePallet.addEventListener('change', applyFilter);
  onlyOpportunities.addEventListener('change', applyFilter);
  searchInput.addEventListener('input', applyFilter);

  function renderStats() {
    const productHeader = originalHeaders.find(h => normKey(h) === 'product');
    const uniqueProducts = new Set(baseRows.map(r => norm(r[productHeader]))).size;
    let singlePalletProducts = 0, multiPalletProducts = 0;
    palletCountByProduct.forEach(count => {
      if (count === 1) singlePalletProducts++; else multiPalletProducts++;
    });

    const stats = [
      { label: 'Regels op Bulk Location', num: baseRows.length },
      { label: 'Unieke artikelen op Bulk Location', num: uniqueProducts },
      { label: 'Artikelen op maar 1 pallet', num: singlePalletProducts },
      { label: 'Artikelen op 2+ pallets (consolidatie)', num: multiPalletProducts },
      { label: 'Genegeerd: verpakkingsmateriaal (DOOS/BOX/TOP)', num: `${noiseExcludedRowCount} regels` },
      { label: 'Genegeerd: CHITA-locaties', num: `${chitaExcludedRowCount} regels` },
      { label: 'Regels in huidig resultaat', num: resultRows.length },
    ];

    // Vulgraad-dekking uitsplitsen naar reden, zodat in de statistieken
    // zelf te zien is WAAROM regels "onbekend" zijn — zonder devtools nodig
    // te hebben. Reden 'ok' = vulgraad kon berekend worden.
    if (!referenceDataReady) {
      stats.push({ label: 'Vulgraad', num: 'referentiebestanden niet geladen' });
    } else {
      // Kernmaat bovenaan: hoeveel plekken maakt deze lijst daadwerkelijk vrij?
      // Optellen per uniek artikel, niet per regel — anders telt hetzelfde
      // artikel dubbel mee via zijn meerdere pallets.
      const productsInResult = new Set(resultRows.map(r => norm(r[productHeader])));
      let totalLocationsFreed = 0, productsWithKnownScore = 0;
      productsInResult.forEach(p => {
        const score = scoreByProduct.get(p);
        if (score) { totalLocationsFreed += score.locationsFreed; productsWithKnownScore++; }
      });
      const toEmptyCount = resultRows.filter(r => r.__action === 'legen').length;
      stats.unshift(
        { label: 'Vrij te maken pallet-plekken', num: totalLocationsFreed },
        { label: 'Pallets leeghalen om dat te bereiken', num: toEmptyCount }
      );

      // Van→naar: bij hoeveel van de te legen pallets is er een concrete,
      // schone bestemming gevonden (zie computeVanNaarMoves)? De rest is nog
      // steeds "Empty", maar zonder specifiek doel — bewust, geen datafout.
      if (toEmptyCount) {
        const withTarget = resultRows.filter(r => r.__action === 'legen' && r.__moveTo).length;
        stats.push({
          label: 'Waarvan met concrete "Naar"-locatie',
          num: `${withTarget} / ${toEmptyCount}`,
        });
      }

      stats.push({
        label: 'Artikelen met een aantoonbare kans',
        num: `${productsWithKnownScore} / ${productsInResult.size}`,
      });

      // Capaciteitsdekking uitsplitsen naar reden, zodat in de statistieken
      // zelf te zien is WAAROM regels onbekend zijn — zonder devtools.
      const reasonCounts = {};
      resultRows.forEach(r => {
        reasonCounts[r.__cap.reason] = (reasonCounts[r.__cap.reason] || 0) + 1;
      });
      const known = reasonCounts['ok'] || 0;
      stats.push({ label: 'Capaciteit bekend (van huidig resultaat)', num: `${known} / ${resultRows.length}` });
      if (known) {
        const okRows = resultRows.filter(r => r.__cap.reason === 'ok');
        const avgFill = okRows.reduce((sum, r) => sum + r.__cap.ratio, 0) / known;
        const totalRest = okRows.reduce((sum, r) => sum + r.__cap.restruimte, 0);
        stats.push({ label: 'Gemiddelde vulgraad (bekend)', num: `${Math.round(avgFill * 100)}%` });
        stats.push({ label: 'Totale restruimte in resultaat (stuks)', num: totalRest });
      }
      const REASON_LABELS = {
        'product-afmetingen-onbetrouwbaar': 'Onbekend: productafmetingen ontbreken of zijn placeholder',
        'locatie-onbruikbaar': 'Onbekend: locatie of palletsoort onbruikbaar',
        'gemengde-pallet': 'Onbekend: meerdere artikelen op dezelfde pallet',
        'past-niet-op-pallet': 'Onbekend: product past niet op deze pallet',
        'product-te-hoog': 'Onbekend: product hoger dan de stapelruimte',
        'locatie-te-laag': 'Onbekend: locatie te laag na pallet-aftrek',
        'aantal-ongeldig': 'Onbekend: aantal ontbreekt of is ongeldig',
      };
      Object.keys(REASON_LABELS).forEach(reasonKey => {
        if (reasonCounts[reasonKey]) {
          stats.push({ label: REASON_LABELS[reasonKey], num: reasonCounts[reasonKey] });
        }
      });

      // Zelfcorrectie zichtbaar maken: hoe vaak stond er méér op een pallet
      // dan volgens de afmetingen past? Dat is een signaal over datakwaliteit,
      // niet iets om stil te houden.
      if (observedOverrideCount) {
        stats.push({
          label: 'Capaciteit bijgesteld op wat er werkelijk op staat',
          num: `${observedOverrideCount} pallets`,
        });
      }
      if (overhangAssumedCount) {
        stats.push({
          label: 'Overhang aangenomen (1 stuk per laag)',
          num: `${overhangAssumedCount} pallets`,
        });
      }

      // Kwaliteit van de referentiebestanden zelf — de bron van bijna alle
      // "onbekend" hierboven.
      if (refStats) {
        const p = refStats.products, l = refStats.locations;
        stats.push({
          label: 'products.xlsx bruikbaar na plausibiliteitstoets',
          num: `${productDimsById.size} / ${p.total}`,
        });
        if (p.rejected.placeholder) {
          stats.push({ label: 'Waarvan afgekeurd: placeholder-afmeting 1×1×1 mm', num: p.rejected.placeholder });
        }
        stats.push({
          label: 'locations.xlsx bruikbaar na plausibiliteitstoets',
          num: `${locationInfoByCode.size} / ${l.total}`,
        });
      }
    }

    statsBox.innerHTML = stats.map(s =>
      `<div class="bg-white dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg px-3.5 py-3">` +
      `<div class="text-xl font-semibold">${s.num}</div>` +
      `<div class="text-xs text-zinc-500 dark:text-zinc-400 mt-0.5">${s.label}</div></div>`
    ).join('');
  }

  // Vulgraad als leesbare tekst voor tabel/export: hoeveel procent van de
  // palletcapaciteit is bezet, of "onbekend" als de capaciteit niet te
  // bepalen was.
  function formatFillRatio(cap) {
    if (!cap || cap.reason !== 'ok') return 'onbekend';
    return `${Math.round(cap.ratio * 100)}%`;
  }

  // Restruimte in stuks — het getal waar een reachtruck-chauffeur direct iets
  // mee kan: hoeveel kan er nog bij op deze pallet?
  function formatRestruimte(cap) {
    if (!cap || cap.reason !== 'ok') return 'onbekend';
    return String(cap.restruimte);
  }

  // Vrij te maken locaties als leesbare tekst: "onbekend" als er geen score
  // berekend kon worden (zie computeConsolidationScores).
  function formatLocationsFreed(score) {
    if (!score) return 'onbekend';
    return String(score.locationsFreed);
  }

  // Doellocatie (van→naar) als leesbare tekst: alleen gevuld bij een schone
  // match (zie computeVanNaarMoves) — anders "-", niet "onbekend", want dit is
  // geen ontbrekende data maar een bewuste keuze om geen versnipperde
  // instructie te tonen.
  function formatMoveTo(moveTo) {
    return moveTo ? moveTo.location : '-';
  }

  // Actie per pallet (Fase 4) als leesbare tekst, zie computeConsolidationActions.
  // Op verzoek van de product owner blijven deze twee statuslabels Engels
  // ("Empty"/"Keep") i.p.v. Nederlands — kort en duidelijk genoeg voor de
  // reachtruck-chauffeurs, de rest van de tool blijft Nederlands.
  function formatAction(action) {
    if (action === 'legen') return 'Empty';
    if (action === 'behouden') return 'Keep';
    if (action === 'onbekend') return 'onbekend';
    return '-'; // geen consolidatiewinst voor dit artikel, geen actie nodig
  }

  function renderPreview() {
    statusEl.style.display = 'block';
    const maxPreview = 200;
    statusEl.textContent = resultRows.length > maxPreview
      ? `Preview: eerste ${maxPreview} van ${resultRows.length} regels.`
      : `${resultRows.length} regels.`;

    const thClass = 'sticky top-0 z-10 bg-zinc-50 dark:bg-zinc-800 text-left font-medium text-xs uppercase tracking-wide text-zinc-500 dark:text-zinc-400 px-3 py-2 border-b border-zinc-200 dark:border-zinc-700 whitespace-nowrap';
    const tdClass = 'px-3 py-2 border-b border-zinc-100 dark:border-zinc-800 whitespace-nowrap';
    const trClass = 'hover:bg-zinc-50 dark:hover:bg-zinc-800/60';

    // Kolomvolgorde op verzoek van de product owner: Location Code, Product
    // Name, Quantity, dan de berekende kolommen (Engelse termen), en Urn als
    // allerlaatste kolom.
    const showFillColumn = referenceDataReady;
    const headers = outputColumns.map(c => c.label)
      .concat(showFillColumn ? ['Fill Rate', 'Action', 'To', 'Locations Freed', 'Remaining'] : [])
      .concat(urnColumn ? [urnColumn.label] : []);
    const thead = '<thead><tr>' + headers.map(h => `<th class="${thClass}">${h}</th>`).join('') + '</tr></thead>';
    const bodyRows = resultRows.slice(0, maxPreview).map(row => {
      const cells = outputColumns.map(c => {
        const val = c.header ? row[c.header] : '';
        return `<td class="${tdClass}">${val === undefined || val === null ? '' : String(val)}</td>`;
      });
      if (showFillColumn) {
        cells.push(`<td class="${tdClass}">${formatFillRatio(row.__cap)}</td>`);
        // "Legen"-pallets extra opvallend (geel accent), zodat ze in de
        // preview meteen te herkennen zijn zonder de tekst te moeten lezen.
        const actionClass = row.__action === 'legen'
          ? `${tdClass} font-medium text-[#8a6d1a] dark:text-[#eab627]`
          : tdClass;
        cells.push(`<td class="${actionClass}">${formatAction(row.__action)}</td>`);
        cells.push(`<td class="${tdClass}">${formatMoveTo(row.__moveTo)}</td>`);
        cells.push(`<td class="${tdClass}">${formatLocationsFreed(row.__score)}</td>`);
        cells.push(`<td class="${tdClass}">${formatRestruimte(row.__cap)}</td>`);
      }
      if (urnColumn) {
        const val = urnColumn.header ? row[urnColumn.header] : '';
        cells.push(`<td class="${tdClass}">${val === undefined || val === null ? '' : String(val)}</td>`);
      }
      return `<tr class="${trClass}">` + cells.join('') + '</tr>';
    }).join('');
    previewTable.innerHTML = thead + '<tbody>' + bodyRows + '</tbody>';
  }

  exportBtn.addEventListener('click', async () => {
    if (!resultRows.length) return;
    exportBtn.disabled = true;
    exportBtn.textContent = 'Bezig met exporteren...';

    try {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Consolidatie');

      // Zelfde kolomvolgorde als de preview (zie renderPreview): Location
      // Code, Product Name, Quantity, dan de berekende kolommen (Engelse
      // termen), en Urn als allerlaatste kolom.
      //
      // Urn, Fill Rate en Locations Freed duwden de export over de rand van
      // 1 A4 liggend. Op verzoek van de product owner staan ze standaard
      // NIET in de export (wel altijd in de preview) — via de toggle
      // "exportExtraColumns" zijn ze er alsnog bij te zetten.
      const includeExtra = exportExtraColumns.checked;
      const computedColumns = referenceDataReady ? [
        { label: 'Fill Rate', header: null, isFillColumn: true },
        { label: 'Action', header: null, isActionColumn: true },
        { label: 'To', header: null, isMoveToColumn: true },
        { label: 'Locations Freed', header: null, isScoreColumn: true },
        { label: 'Remaining', header: null, isRestColumn: true },
      ] : [];
      const filteredComputedColumns = includeExtra
        ? computedColumns
        : computedColumns.filter(c => c.label !== 'Fill Rate' && c.label !== 'Locations Freed');
      const exportColumns = outputColumns
        .concat(filteredComputedColumns)
        .concat((urnColumn && includeExtra) ? [urnColumn] : []);
      const valueFor = (row, c) => c.isFillColumn
        ? formatFillRatio(row.__cap)
        : c.isRestColumn
        ? formatRestruimte(row.__cap)
        : c.isScoreColumn
        ? formatLocationsFreed(row.__score)
        : c.isActionColumn
        ? formatAction(row.__action)
        : c.isMoveToColumn
        ? formatMoveTo(row.__moveTo)
        : (c.header ? row[c.header] : '');

      // Kolombreedte: de meeste kolommen krijgen net genoeg breedte voor hun
      // eigen inhoud, Product Name krijgt de rest van de ruimte — dat is de
      // kolom die je wilt kunnen lezen zonder afkapping.
      const WIDTH_CAPS = {
        'Location Code': { min: 12, max: 22 },
        'Quantity': { min: 8, max: 12 },
        'Urn': { min: 12, max: 20 },
        'Product Name': { min: 40, max: 90 },
        'Fill Rate': { min: 10, max: 12 },
        'Remaining': { min: 11, max: 14 },
        'Locations Freed': { min: 12, max: 18 },
        'Action': { min: 8, max: 12 },
        'To': { min: 12, max: 22 },
      };
      sheet.columns = exportColumns.map(c => {
        const cap = WIDTH_CAPS[c.label] || { min: 12, max: 30 };
        let maxLen = c.label.length;
        resultRows.forEach(row => {
          maxLen = Math.max(maxLen, String(valueFor(row, c) ?? '').length);
        });
        const width = Math.min(Math.max(maxLen + 2, cap.min), cap.max);
        return { header: c.label, key: c.label, width };
      });

      resultRows.forEach(row => {
        const rowData = {};
        exportColumns.forEach(c => { rowData[c.label] = valueFor(row, c); });
        sheet.addRow(rowData);
      });

      // Koprij vet + vastgezet, plus filterknoppen, zodat het bestand direct
      // bruikbaar is in Excel zonder verdere opmaak.
      sheet.getRow(1).font = { bold: true };
      sheet.views = [{ state: 'frozen', ySplit: 1 }];
      sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: exportColumns.length } };

      // Printopmaak: A4 liggend, alles op 1 pagina breed (hoogte mag over
      // meerdere pagina's, dat hoeft niet passend gemaakt te worden).
      sheet.pageSetup = {
        paperSize: 9, // 9 = A4
        orientation: 'landscape',
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
        margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 },
      };

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/octet-stream' });
      const today = new Date().toISOString().slice(0, 10);
      const filename = `rapportage_consolidatie_pallets_${today}.xlsx`;

      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } finally {
      exportBtn.disabled = false;
      exportBtn.textContent = 'Download resultaat als .xlsx';
    }
  });
})();
