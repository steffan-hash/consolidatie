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
          (rapportage_consolidatie_pallets_YYYY-MM-DD.xlsx). Het resultaat
          bevat de 4 vaste kolommen (Location Code, Product Name, Quantity,
          Urn) plus een berekende kolom Vulgraad — de rest van de brondata
          is niet nodig om pallets te vinden en fysiek te consolideren.
          Printopmaak staat op A4 liggend, geschaald naar 1 pagina breed,
          met Product Name zo breed als de andere kolommen toelaten.

  Vulgraad (2.0): naast de voorraadexport laadt de tool automatisch twee
          vaste referentiebestanden uit data/reference/ (products.xlsx en
          locations.xlsx, met Length/Width/Height in mm) om te berekenen
          hoeveel % van een locatie daadwerkelijk bezet is. Dit zijn géén
          uploads — de bestanden staan vast in de repo en worden zelden
          bijgewerkt. Ontbreken afmetingen van een product of locatie, dan
          wordt de vulgraad "onbekend" i.p.v. dat er iets geraden wordt.
          Omdozen worden genegeerd (afmetingen daarvan zijn niet bekend);
          om te voorkomen dat de vulgraad daardoor te optimistisch wordt,
          telt een vast percentage van de locatie-inhoud (OMDOOS_MARGE)
          niet mee.

  Consolidatiepotentieel (2.0, Fase 3): voor elk artikel op 2+ pallets wordt
          berekend hoeveel pallets er minimaal nodig zijn als je alles
          consolideert op de grootste locatie die het artikel al gebruikt.
          Het verschil met het huidige aantal pallets ("Vrij te maken
          locaties") is de nieuwe sortering — meeste winst bovenaan, i.p.v.
          alfabetisch op productnaam.

  Werklijst voor de reachers (2.0, Fase 4): naast het totale aantal "vrij te
          maken locaties" per artikel (Fase 3) wijst de tool nu ook per
          pallet-regel een concrete actie aan: "Legen" voor de pallets met de
          laagste vulgraad van dat artikel (zoveel als er vrij te maken zijn),
          "Behouden" voor de rest. Alleen als van ALLE pallets van dat artikel
          de vulgraad bekend is — anders is niet veilig te bepalen welke
          specifieke pallet de laagste is, en toont de kolom "onbekend".

  Zoekfilter: een zoekveld boven de resultaattabel filtert op de zichtbare
          kolommen (Location Code, Product Name, Quantity, Urn). Werkt als
          laatste stap ná de score/actie-berekening, zodat zoeken naar 1
          pallet van een artikel niet de "vrij te maken locaties"/Legen-Keep
          bepaling voor dat artikel verstoort — die blijft naar alle pallets
          van het artikel kijken, ook de pallets die het zoekresultaat niet
          toont.

  Ruisreductie: drie automatische uitsluitingen om het resultaat te beperken
          tot echte consolidatiekansen i.p.v. duizenden regels. (1) Producten
          met "DOOS", "BOX" of "TOP" als los woord in de naam (verpakkings-
          materiaal) tellen nergens in mee. (2) Een artikel op meer dan
          UNIFORM_STACKING_MIN_PALLETS pallets, waarvan alle pallets exact
          dezelfde hoeveelheid én (afgeronde) vulgraad hebben, wordt ook
          genegeerd — dat patroon wijst op een standaard, al-optimale
          stapelwijze. (3) Locaties waarvan de code "CHITA" bevat tellen
          nergens in mee — geen gewone bulklocaties in het magazijnrek. Alle
          drie uitsluitingen zijn zichtbaar in de statistieken (aantal
          genegeerde regels/artikelen), niet stilzwijgend.
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
  const productGroupSection = document.getElementById('productGroupSection');
  const productGroupList = document.getElementById('productGroupList');
  const selectAllGroups = document.getElementById('selectAllGroups');
  const selectNoneGroups = document.getElementById('selectNoneGroups');
  const searchInput = document.getElementById('searchInput');
  const statsBox = document.getElementById('statsBox');
  const statusEl = document.getElementById('status');
  const previewTable = document.getElementById('previewTable');
  const exportBtn = document.getElementById('exportBtn');
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

  // Ruisreductie 2: staat een artikel op meer dan dit aantal pallets, én
  // hebben al die pallets exact dezelfde hoeveelheid én dezelfde (afgeronde)
  // vulgraad, dan gaan we ervan uit dat dit de standaard/al-optimale manier
  // van stapelen is voor dit artikel — niets aan te consolideren. Zie
  // computeUniformStackingProducts().
  const UNIFORM_STACKING_MIN_PALLETS = 10;

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
  const LOCATION_REF_HEADERS = ['location', 'length', 'width', 'height'];

  // De productafmetingen zijn van het kale product, niet van de omdoos
  // waarin het op de pallet ligt (die afmetingen zijn niet bekend). Om te
  // voorkomen dat de vulgraad daardoor te optimistisch wordt ingeschat, telt
  // dit percentage van de locatie-inhoud niet mee als bruikbare ruimte. Dit
  // is een inschatting (geen gemeten waarde) — bijstellen kan door dit getal
  // aan te passen als de praktijk daar aanleiding toe geeft.
  const OMDOOS_MARGE = 0.15;

  // De locatiehoogte uit het referentiebestand is de hoogte van de hele
  // opslaglocatie, maar de producten staan op een europallet die zelf ook
  // hoogte inneemt. Dit aantal mm gaat van de locatiehoogte af voordat de
  // vulgraad berekend wordt, want dat is geen ruimte voor het product zelf.
  const PALLET_HOOGTE_MM = 200;

  // Vaste set kolommen voor het resultaat — niet instelbaar in de UI. De rest
  // van de brondata is ruis voor het consolideren van pallets.
  const OUTPUT_COLUMNS = [
    { key: 'location code', label: 'Location Code' },
    { key: 'description', label: 'Product Name' },
    { key: 'quantity', label: 'Quantity' },
    { key: 'urn', label: 'Urn' },
  ];

  let originalHeaders = [];   // koprij in originele volgorde, zoals in het bestand
  let outputColumns = [];     // OUTPUT_COLUMNS aangevuld met de bijbehorende originele headernaam
  let baseRows = [];          // rijen op een "Bulk Location", als objecten {header: waarde}
  let palletCountByProduct = new Map(); // Product -> aantal unieke pallets (Urn's)
  let resultRows = [];        // rijen die momenteel getoond/geëxporteerd worden
  let productGroupHeader = null;        // originele headernaam van "Product Group", of null als kolom ontbreekt
  let selectedProductGroups = new Set(); // welke Product Group waarden momenteel getoond worden

  let productDimsById = new Map();    // genormaliseerde Product ID -> {length,width,height} in mm
  let locationDimsByCode = new Map(); // genormaliseerde Location Code -> {length,width,height} in mm
  let referenceDataReady = false;     // true zodra beide referentiebestanden geladen en gekoppeld zijn

  let totalQtyByProduct = new Map();    // Product -> totale hoeveelheid over al zijn Bulk Location-pallets
  let locationSetByProduct = new Map(); // Product -> Set van genormaliseerde Location Codes waar het nu op staat
  let scoreByProduct = new Map();       // Product -> {minPalletsNeeded, locationsFreed}, zie computeConsolidationScores

  let noiseExcludedRowCount = 0;         // aantal regels genegeerd door NOISE_PRODUCT_KEYWORDS (verpakkingsmateriaal)
  let chitaExcludedRowCount = 0;         // aantal regels genegeerd door NOISE_LOCATION_KEYWORD (CHITA-locaties)
  let uniformStackingProducts = new Set(); // producten genegeerd door computeUniformStackingProducts

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

  // Bouwt een opzoekkaart keyField -> {length,width,height} in mm. Rijen
  // zonder (volledige) afmetingen worden overgeslagen — die tellen straks
  // gewoon als "onbekend" in plaats van dat er iets geraden wordt.
  function buildDimsMap(rows, keyField) {
    const map = new Map();
    rows.forEach(row => {
      const key = normKey(row[keyField]);
      if (key === '') return;
      const length = Number(row['length']);
      const width = Number(row['width']);
      const height = Number(row['height']);
      if (!length || !width || !height) return;
      map.set(key, { length, width, height });
    });
    return map;
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
      productDimsById = buildDimsMap(products, 'product id');
      locationDimsByCode = buildDimsMap(locations, 'location');
      referenceDataReady = true;
      // Zichtbaar in de browserconsole (F12) zodat te zien is hoeveel van de
      // referentiebestanden daadwerkelijk bruikbare afmetingen bevatten.
      console.info(
        `Referentiedata geladen: producten ${productDimsById.size}/${products.length} met afmetingen, ` +
        `locaties ${locationDimsByCode.size}/${locations.length} met afmetingen.`
      );
    } catch (e) {
      console.warn('Referentiedata (locatie-/productafmetingen) kon niet geladen worden — vulgraad wordt niet berekend.', e);
      referenceDataReady = false;
    }
    refDataHint.style.display = referenceDataReady ? 'none' : 'block';
    if (baseRows.length) applyFilter(); // alsnog herberekenen als data na de eerste render binnenkomt
  }
  loadReferenceData();

  // Vulgraad van 1 regel: hoeveel % van de bruikbare locatie-inhoud is bezet
  // door de hoeveelheid van dit product. Geeft altijd een reden mee als het
  // niet lukt, zodat in de statistieken zichtbaar is WAAROM iets "onbekend"
  // is (referentiedata niet geladen, product/locatie niet in het
  // referentiebestand, geen geldig aantal, of locatie te laag).
  function computeFillRatio(row, productHeader, locationHeader, quantityHeader) {
    if (!referenceDataReady) return { ratio: null, reason: 'referentiedata-niet-geladen' };

    const product = productDimsById.get(normKey(row[productHeader]));
    if (!product) return { ratio: null, reason: 'product-onbekend' };

    const location = locationDimsByCode.get(normKey(row[locationHeader]));
    if (!location) return { ratio: null, reason: 'locatie-onbekend' };

    const qty = Number(row[quantityHeader]);
    if (!qty || qty <= 0) return { ratio: null, reason: 'aantal-ongeldig' };

    const productVolume = product.length * product.width * product.height;
    const usableHeight = Math.max(location.height - PALLET_HOOGTE_MM, 0);
    const usableLocationVolume = location.length * location.width * usableHeight * (1 - OMDOOS_MARGE);
    if (usableLocationVolume <= 0) return { ratio: null, reason: 'locatie-te-laag' };

    return { ratio: (productVolume * qty) / usableLocationVolume, reason: 'ok' };
  }

  // Fase 3 — consolidatiepotentieel: voor elk artikel op 2+ pallets, hoeveel
  // pallets zijn er minimaal nodig als je alles consolideert op de grootste
  // locatie die het artikel al gebruikt? Het verschil met het huidige aantal
  // pallets is het aantal locaties dat écht vrijgemaakt kan worden — dat is
  // de nieuwe prioriteitsscore/sortering (i.p.v. alfabetisch). Geeft geen
  // entry terug (dus "onbekend" bij gebruik) als productafmetingen ontbreken
  // of geen van de gebruikte locaties bruikbare afmetingen heeft.
  function computeConsolidationScores() {
    const scores = new Map();
    if (!referenceDataReady) return scores;

    palletCountByProduct.forEach((currentPallets, product) => {
      const productDims = productDimsById.get(normKey(product));
      if (!productDims) return;

      // Grootste bruikbare inhoud onder de locaties waar dit artikel nu al
      // op staat — consolideren gebeurt op een bestaande locatie, niet op
      // een hypothetische locatie elders in het magazijn.
      let maxUsableVolume = 0;
      (locationSetByProduct.get(product) || new Set()).forEach(locKey => {
        const loc = locationDimsByCode.get(locKey);
        if (!loc) return;
        const usableHeight = Math.max(loc.height - PALLET_HOOGTE_MM, 0);
        const usableVolume = loc.length * loc.width * usableHeight * (1 - OMDOOS_MARGE);
        if (usableVolume > maxUsableVolume) maxUsableVolume = usableVolume;
      });
      if (maxUsableVolume <= 0) return;

      const totalVolume = productDims.length * productDims.width * productDims.height
        * (totalQtyByProduct.get(product) || 0);
      const minPalletsNeeded = Math.max(1, Math.ceil(totalVolume / maxUsableVolume));
      const locationsFreed = Math.max(0, currentPallets - minPalletsNeeded);
      scores.set(product, { minPalletsNeeded, locationsFreed });
    });

    return scores;
  }

  // Fase 4 — actie per pallet: van elk artikel met "vrij te maken locaties"
  // (Fase 3) wijzen we de pallets met de laagste vulgraad aan als "te legen"
  // (voorraad overhevelen naar een andere pallet van hetzelfde artikel), de
  // rest als "te behouden" (ontvangt die overgehevelde voorraad). Dit kan
  // alleen betrouwbaar als van ALLE pallets van dat artikel de vulgraad
  // bekend is — is er twijfel, dan raden we niet welke pallet het is en
  // krijgen alle regels van dat artikel "onbekend". Vereist dat __fillInfo
  // en __score al gezet zijn op de meegegeven rijen (zie applyFilter).
  function computeConsolidationActions(rows, productHeader) {
    const rowsByProduct = new Map();
    rows.forEach(row => {
      const product = norm(row[productHeader]);
      if (!rowsByProduct.has(product)) rowsByProduct.set(product, []);
      rowsByProduct.get(product).push(row);
    });

    rowsByProduct.forEach((productRows, product) => {
      const score = scoreByProduct.get(product);
      if (!score || score.locationsFreed <= 0) {
        productRows.forEach(row => { row.__action = null; }); // geen consolidatiewinst, geen actie nodig
        return;
      }

      const allKnown = productRows.every(row => row.__fillInfo && row.__fillInfo.reason === 'ok');
      if (!allKnown) {
        productRows.forEach(row => { row.__action = 'onbekend'; });
        return;
      }

      // Nooit ALLE pallets van een artikel als "te legen" aanwijzen — er moet
      // altijd minstens 1 pallet overblijven om de voorraad op te ontvangen.
      const toEmptyCount = Math.min(score.locationsFreed, Math.max(productRows.length - 1, 0));
      const sortedByFill = productRows.slice().sort((a, b) => a.__fillInfo.ratio - b.__fillInfo.ratio);
      const toEmpty = new Set(sortedByFill.slice(0, toEmptyCount));
      productRows.forEach(row => { row.__action = toEmpty.has(row) ? 'legen' : 'behouden'; });
    });
  }

  // Ruisreductie 2: producten met meer dan UNIFORM_STACKING_MIN_PALLETS
  // pallets, waarvan alle pallets exact dezelfde hoeveelheid en dezelfde
  // (afgeronde) vulgraad hebben, negeren — dat patroon wijst op een
  // standaard, al-optimale stapelwijze, geen consolidatiekans. Bij twijfel
  // (vulgraad van 1 of meer pallets onbekend) wordt een artikel NIET
  // uitgesloten, want de aanname is dan niet hard te maken. Vereist dat
  // row.__fillInfo al gezet is voor alle rijen van het artikel (zie
  // applyFilter, dat dit voor baseRows doet vóór deze functie aan te roepen).
  function computeUniformStackingProducts(productHeader, quantityHeader) {
    const rowsByProduct = new Map();
    baseRows.forEach(row => {
      const product = norm(row[productHeader]);
      if (!rowsByProduct.has(product)) rowsByProduct.set(product, []);
      rowsByProduct.get(product).push(row);
    });

    const uniform = new Set();
    rowsByProduct.forEach((rows, product) => {
      if ((palletCountByProduct.get(product) || 0) <= UNIFORM_STACKING_MIN_PALLETS) return;

      const firstFill = rows[0].__fillInfo;
      if (!firstFill || firstFill.reason !== 'ok') return;
      const firstQty = norm(rows[0][quantityHeader]);
      const firstPct = Math.round(firstFill.ratio * 100);

      const allSame = rows.every(row => {
        const fill = row.__fillInfo;
        return fill && fill.reason === 'ok'
          && norm(row[quantityHeader]) === firstQty
          && Math.round(fill.ratio * 100) === firstPct;
      });
      if (allSame) uniform.add(product);
    });

    return uniform;
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
    palletCountByProduct = new Map();
    totalQtyByProduct = new Map();
    locationSetByProduct = new Map();
    scoreByProduct = new Map();
    noiseExcludedRowCount = 0;
    chitaExcludedRowCount = 0;
    uniformStackingProducts = new Set();
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
    // plus (t.b.v. Fase 3) de totale hoeveelheid en de set locaties waar het
    // artikel nu op staat — nodig om straks het consolidatiepotentieel te
    // berekenen.
    const productHeader = originalHeaders.find(h => normKey(h) === 'product');
    const urnHeader = originalHeaders.find(h => normKey(h) === 'urn');
    const locationHeaderForTotals = originalHeaders.find(h => normKey(h) === 'location code');
    const quantityHeaderForTotals = originalHeaders.find(h => normKey(h) === 'quantity');
    const palletSetsByProduct = new Map();
    totalQtyByProduct = new Map();
    locationSetByProduct = new Map();
    baseRows.forEach(row => {
      const product = norm(row[productHeader]);
      const urn = norm(row[urnHeader]);
      if (!palletSetsByProduct.has(product)) palletSetsByProduct.set(product, new Set());
      palletSetsByProduct.get(product).add(urn);

      totalQtyByProduct.set(product, (totalQtyByProduct.get(product) || 0) + (Number(row[quantityHeaderForTotals]) || 0));
      if (!locationSetByProduct.has(product)) locationSetByProduct.set(product, new Set());
      locationSetByProduct.get(product).add(normKey(row[locationHeaderForTotals]));
    });
    palletCountByProduct = new Map();
    palletSetsByProduct.forEach((set, product) => palletCountByProduct.set(product, set.size));

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

    // Vulgraad vooraf voor ALLE basisregels berekenen (niet pas na filteren)
    // — computeUniformStackingProducts hieronder moet naar alle pallets van
    // een artikel kunnen kijken, los van de huidige UI-filters.
    baseRows.forEach(row => {
      row.__fillInfo = computeFillRatio(row, productHeader, locationHeader, quantityHeader);
    });

    scoreByProduct = computeConsolidationScores();
    uniformStackingProducts = computeUniformStackingProducts(productHeader, quantityHeader);

    resultRows = baseRows.filter(row => {
      const product = norm(row[productHeader]);
      if (uniformStackingProducts.has(product)) return false; // ruisreductie 2
      if (hideSinglePallet.checked) {
        if ((palletCountByProduct.get(product) || 0) < 2) return false;
      }
      if (productGroupHeader && !selectedProductGroups.has(norm(row[productGroupHeader]))) return false;
      return true;
    });

    // Sorteren op consolidatiewinst (meeste vrij te maken locaties bovenaan —
    // Fase 3), dan op product (zodat alle pallets van hetzelfde artikel bij
    // elkaar staan), en als laatste op vulgraad oplopend — zo staan binnen
    // een artikel de pallets met de laagste vulgraad (de te legen pallets,
    // Fase 4) vanzelf boven de te behouden pallets. Artikelen zonder bekende
    // score (-1) staan onderaan; pallets zonder bekende vulgraad staan als
    // laatste binnen hun artikel.
    resultRows = resultRows.slice().sort((a, b) => {
      const pa = norm(a[productHeader]), pb = norm(b[productHeader]);
      const freedA = scoreByProduct.has(pa) ? scoreByProduct.get(pa).locationsFreed : -1;
      const freedB = scoreByProduct.has(pb) ? scoreByProduct.get(pb).locationsFreed : -1;
      if (freedA !== freedB) return freedB - freedA;
      if (pa !== pb) return pa < pb ? -1 : 1;
      const fa = a.__fillInfo && a.__fillInfo.reason === 'ok' ? a.__fillInfo.ratio : Infinity;
      const fb = b.__fillInfo && b.__fillInfo.reason === 'ok' ? b.__fillInfo.ratio : Infinity;
      if (fa !== fb) return fa - fb;
      const la = norm(a[locationHeader]), lb = norm(b[locationHeader]);
      return la < lb ? -1 : la > lb ? 1 : 0;
    });

    // Consolidatiescore per regel vastleggen voor weergave/export (vulgraad
    // staat al op de rij, zie hierboven), en daarna pas de actie per pallet
    // bepalen (Fase 4) — die heeft zowel __score als __fillInfo nodig.
    resultRows.forEach(row => {
      row.__score = scoreByProduct.get(norm(row[productHeader])) || null;
    });
    computeConsolidationActions(resultRows, productHeader);

    // Zoekfilter als laatste stap: zoekt in exact de kolommen die ook in de
    // tabel te zien zijn (Location Code, Product Name, Quantity, Urn). Pas
    // NA de score/actie-berekening toegepast, want die moet naar alle
    // pallets van een artikel kunnen kijken — anders zou zoeken op 1 pallet
    // van een artikel de "vrij te maken locaties"/Legen-Keep-berekening voor
    // dat artikel verstoren.
    const searchTerm = normKey(searchInput.value);
    if (searchTerm) {
      resultRows = resultRows.filter(row =>
        outputColumns.some(c => c.header && normKey(row[c.header]).includes(searchTerm))
      );
    }

    renderStats();
    renderPreview();
  }
  hideSinglePallet.addEventListener('change', applyFilter);
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
      const reasonCounts = {};
      resultRows.forEach(r => {
        reasonCounts[r.__fillInfo.reason] = (reasonCounts[r.__fillInfo.reason] || 0) + 1;
      });
      const known = reasonCounts['ok'] || 0;
      stats.push({ label: 'Vulgraad bekend (van huidig resultaat)', num: `${known} / ${resultRows.length}` });
      if (known) {
        const avgFill = resultRows
          .filter(r => r.__fillInfo.reason === 'ok')
          .reduce((sum, r) => sum + r.__fillInfo.ratio, 0) / known;
        stats.push({ label: 'Gemiddelde vulgraad (bekend)', num: `${Math.round(avgFill * 100)}%` });
      }
      const REASON_LABELS = {
        'product-onbekend': 'Waarvan: productafmetingen niet in referentiebestand',
        'locatie-onbekend': 'Waarvan: locatieafmetingen niet in referentiebestand',
        'aantal-ongeldig': 'Waarvan: aantal ontbreekt/ongeldig',
        'locatie-te-laag': 'Waarvan: locatie te laag na pallet-aftrek',
      };
      Object.keys(REASON_LABELS).forEach(reasonKey => {
        if (reasonCounts[reasonKey]) {
          stats.push({ label: REASON_LABELS[reasonKey], num: reasonCounts[reasonKey] });
        }
      });

      // Vrij te maken locaties optellen per uniek artikel (niet per regel,
      // anders telt hetzelfde artikel dubbel mee via zijn meerdere pallets).
      const productsInResult = new Set(resultRows.map(r => norm(r[productHeader])));
      let totalLocationsFreed = 0, productsWithKnownScore = 0;
      productsInResult.forEach(p => {
        const score = scoreByProduct.get(p);
        if (score) { totalLocationsFreed += score.locationsFreed; productsWithKnownScore++; }
      });
      stats.push({
        label: 'Vrij te maken locaties (huidig resultaat)',
        num: `${totalLocationsFreed} (${productsWithKnownScore}/${productsInResult.size} artikelen bekend)`,
      });

      // Fase 4: hoeveel pallets zijn concreet aangewezen om te legen, en bij
      // hoeveel artikelen kon dat niet (vulgraad van 1+ pallets onbekend).
      const toEmptyCount = resultRows.filter(r => r.__action === 'legen').length;
      const unknownActionProducts = new Set(
        resultRows.filter(r => r.__action === 'onbekend').map(r => norm(r[productHeader]))
      ).size;
      stats.push({
        label: 'Pallets aangewezen om te legen (Actie)',
        num: unknownActionProducts
          ? `${toEmptyCount} (+${unknownActionProducts} artikelen onbekend)`
          : `${toEmptyCount}`,
      });

      // Ruisreductie 2 zichtbaar maken: hoeveel artikelen (en dus pallets)
      // zijn genegeerd omdat ze al gelijkmatig/optimaal gestapeld staan.
      if (uniformStackingProducts.size) {
        let uniformPalletCount = 0;
        uniformStackingProducts.forEach(p => { uniformPalletCount += palletCountByProduct.get(p) || 0; });
        stats.push({
          label: `Genegeerd: gelijkmatige stapeling (>${UNIFORM_STACKING_MIN_PALLETS} pallets, zelfde aantal/vulgraad)`,
          num: `${uniformStackingProducts.size} artikelen (${uniformPalletCount} pallets)`,
        });
      }
    }

    statsBox.innerHTML = stats.map(s =>
      `<div class="bg-white dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg px-3.5 py-3">` +
      `<div class="text-xl font-semibold">${s.num}</div>` +
      `<div class="text-xs text-zinc-500 dark:text-zinc-400 mt-0.5">${s.label}</div></div>`
    ).join('');
  }

  // Vulgraad als leesbare tekst voor tabel/export: percentage, of "onbekend"
  // als afmetingen van product of locatie ontbreken.
  function formatFillRatio(fillInfo) {
    if (!fillInfo || fillInfo.ratio === null || fillInfo.ratio === undefined) return 'onbekend';
    return `${Math.round(fillInfo.ratio * 100)}%`;
  }

  // Vrij te maken locaties (Fase 3) als leesbare tekst: "onbekend" als er geen
  // score berekend kon worden (zie computeConsolidationScores).
  function formatLocationsFreed(score) {
    if (!score) return 'onbekend';
    return String(score.locationsFreed);
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

    const showFillColumn = referenceDataReady;
    const headers = outputColumns.map(c => c.label).concat(showFillColumn ? ['Vulgraad', 'Vrij te maken locaties', 'Actie'] : []);
    const thead = '<thead><tr>' + headers.map(h => `<th class="${thClass}">${h}</th>`).join('') + '</tr></thead>';
    const bodyRows = resultRows.slice(0, maxPreview).map(row => {
      const cells = outputColumns.map(c => {
        const val = c.header ? row[c.header] : '';
        return `<td class="${tdClass}">${val === undefined || val === null ? '' : String(val)}</td>`;
      });
      if (showFillColumn) {
        cells.push(`<td class="${tdClass}">${formatFillRatio(row.__fillInfo)}</td>`);
        cells.push(`<td class="${tdClass}">${formatLocationsFreed(row.__score)}</td>`);
        // "Legen"-pallets extra opvallend (geel accent), zodat ze in de
        // preview meteen te herkennen zijn zonder de tekst te moeten lezen.
        const actionClass = row.__action === 'legen'
          ? `${tdClass} font-medium text-[#8a6d1a] dark:text-[#eab627]`
          : tdClass;
        cells.push(`<td class="${actionClass}">${formatAction(row.__action)}</td>`);
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

      // De 4 vaste brondkolommen, plus Vulgraad, Vrij te maken locaties
      // (Fase 3) en Actie (Fase 4) als berekende kolommen — maar alleen als
      // de referentiedata geladen kon worden.
      const exportColumns = outputColumns.concat(
        referenceDataReady ? [
          { label: 'Vulgraad', header: null, isFillColumn: true },
          { label: 'Vrij te maken locaties', header: null, isScoreColumn: true },
          { label: 'Actie', header: null, isActionColumn: true },
        ] : []
      );
      const valueFor = (row, c) => c.isFillColumn
        ? formatFillRatio(row.__fillInfo)
        : c.isScoreColumn
        ? formatLocationsFreed(row.__score)
        : c.isActionColumn
        ? formatAction(row.__action)
        : (c.header ? row[c.header] : '');

      // Kolombreedte: Location Code/Quantity/Urn/Vulgraad/Vrij te maken
      // locaties krijgen net genoeg breedte voor hun eigen inhoud, Product
      // Name krijgt de rest van de ruimte — dat is de kolom die je wilt
      // kunnen lezen zonder afkapping.
      const WIDTH_CAPS = {
        'Location Code': { min: 12, max: 22 },
        'Quantity': { min: 8, max: 12 },
        'Urn': { min: 12, max: 20 },
        'Product Name': { min: 40, max: 90 },
        'Vulgraad': { min: 10, max: 12 },
        'Vrij te maken locaties': { min: 12, max: 22 },
        'Actie': { min: 12, max: 40 },
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
