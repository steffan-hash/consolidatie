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
  const statsBox = document.getElementById('statsBox');
  const statusEl = document.getElementById('status');
  const previewTable = document.getElementById('previewTable');
  const exportBtn = document.getElementById('exportBtn');
  const refDataHint = document.getElementById('refDataHint');

  // Kolomnamen die we nodig hebben om het bestand te kunnen interpreteren.
  // Vergelijking gebeurt case-insensitive en na trimmen van spaties, zodat
  // kleine verschillen tussen exports geen probleem zijn.
  const STOCK_REQUIRED_HEADERS = ['product', 'location code', 'urn', 'stocklocationtypename'];
  const PALLET_LOCATION_TYPE = 'bulk location'; // exact deze waarde telt mee, "Bulk Location Extern" dus niet

  // Referentiebestanden voor de vulgraadberekening (2.0) — vaste bestanden in
  // de repo, geen upload. Koppelveld voor producten is "Product ID" (komt
  // overeen met de "Product"-kolom in de voorraadexport), voor locaties is
  // dat "Location" (komt overeen met "Location Code").
  const REF_PRODUCTS_URL = 'data/reference/products.xlsx';
  const REF_LOCATIONS_URL = 'data/reference/locations.xlsx';
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
    productGroupHeader = null;
    selectedProductGroups = new Set();
    productGroupSection.style.display = 'none';
    productGroupList.innerHTML = '';
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

    // Alle datarijen na de koprij omzetten naar objecten, en meteen filteren
    // op StockLocationTypeName === "Bulk Location". Voorraad op pick-locaties
    // of externe bulklocaties telt niet mee voor consolidatie: die pallets
    // kunnen we niet fysiek samenvoegen met de rest van het magazijn.
    baseRows = [];
    for (let r = best.rowIndex + 1; r < best.allRows.length; r++) {
      const row = best.allRows[r];
      if (normKey(row[idxLocType]) !== PALLET_LOCATION_TYPE) continue;
      if (norm(row[idxUrn]) === '') continue; // geen pallet-ID: kan niet meegeteld worden

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

    // Per product het aantal unieke pallets (Urn's) op bulklocaties tellen.
    const productHeader = originalHeaders.find(h => normKey(h) === 'product');
    const urnHeader = originalHeaders.find(h => normKey(h) === 'urn');
    const palletSetsByProduct = new Map();
    baseRows.forEach(row => {
      const product = norm(row[productHeader]);
      const urn = norm(row[urnHeader]);
      if (!palletSetsByProduct.has(product)) palletSetsByProduct.set(product, new Set());
      palletSetsByProduct.get(product).add(urn);
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
      return `<label><input type="checkbox" class="pgCheckbox" value="${g}" id="${id}" checked> ${g}</label>`;
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

  // Past de "verberg 1-pallet artikelen" filter toe op de al-op-Bulk-Location
  // gefilterde basisdata, en ververst de weergave.
  function applyFilter() {
    const productHeader = originalHeaders.find(h => normKey(h) === 'product');
    const locationHeader = originalHeaders.find(h => normKey(h) === 'location code');
    const quantityHeader = originalHeaders.find(h => normKey(h) === 'quantity');

    resultRows = baseRows.filter(row => {
      if (hideSinglePallet.checked) {
        const product = norm(row[productHeader]);
        if ((palletCountByProduct.get(product) || 0) < 2) return false;
      }
      if (productGroupHeader && !selectedProductGroups.has(norm(row[productGroupHeader]))) return false;
      return true;
    });

    // Sorteren op product en daarna locatie, zodat alle pallets van hetzelfde
    // artikel bij elkaar staan — makkelijker om ze fysiek te consolideren.
    resultRows = resultRows.slice().sort((a, b) => {
      const pa = norm(a[productHeader]), pb = norm(b[productHeader]);
      if (pa !== pb) return pa < pb ? -1 : 1;
      const la = norm(a[locationHeader]), lb = norm(b[locationHeader]);
      return la < lb ? -1 : la > lb ? 1 : 0;
    });

    // Vulgraad per regel opnieuw berekenen (bijv. nodig als de referentie-
    // data pas na het inladen van de voorraad binnenkomt).
    resultRows.forEach(row => {
      row.__fillInfo = computeFillRatio(row, productHeader, locationHeader, quantityHeader);
    });

    renderStats();
    renderPreview();
  }
  hideSinglePallet.addEventListener('change', applyFilter);

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
    }

    statsBox.innerHTML = stats.map(s =>
      `<div class="stat"><div class="num">${s.num}</div><div class="label">${s.label}</div></div>`
    ).join('');
  }

  // Vulgraad als leesbare tekst voor tabel/export: percentage, of "onbekend"
  // als afmetingen van product of locatie ontbreken.
  function formatFillRatio(fillInfo) {
    if (!fillInfo || fillInfo.ratio === null || fillInfo.ratio === undefined) return 'onbekend';
    return `${Math.round(fillInfo.ratio * 100)}%`;
  }

  function renderPreview() {
    statusEl.style.display = 'block';
    const maxPreview = 200;
    statusEl.textContent = resultRows.length > maxPreview
      ? `Preview: eerste ${maxPreview} van ${resultRows.length} regels.`
      : `${resultRows.length} regels.`;

    const showFillColumn = referenceDataReady;
    const headers = outputColumns.map(c => c.label).concat(showFillColumn ? ['Vulgraad'] : []);
    const thead = '<thead><tr>' + headers.map(h => `<th>${h}</th>`).join('') + '</tr></thead>';
    const bodyRows = resultRows.slice(0, maxPreview).map(row => {
      const cells = outputColumns.map(c => {
        const val = c.header ? row[c.header] : '';
        return `<td>${val === undefined || val === null ? '' : String(val)}</td>`;
      });
      if (showFillColumn) cells.push(`<td>${formatFillRatio(row.__fillInfo)}</td>`);
      return '<tr>' + cells.join('') + '</tr>';
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

      // De 4 vaste brondkolommen, plus Vulgraad als berekende 5e kolom —
      // maar alleen als de referentiedata geladen kon worden.
      const exportColumns = outputColumns.concat(
        referenceDataReady ? [{ label: 'Vulgraad', header: null, isFillColumn: true }] : []
      );
      const valueFor = (row, c) => c.isFillColumn
        ? formatFillRatio(row.__fillInfo)
        : (c.header ? row[c.header] : '');

      // Kolombreedte: Location Code/Quantity/Urn/Vulgraad krijgen net genoeg
      // breedte voor hun eigen inhoud, Product Name krijgt de rest van de
      // ruimte — dat is de kolom die je wilt kunnen lezen zonder afkapping.
      const WIDTH_CAPS = {
        'Location Code': { min: 12, max: 22 },
        'Quantity': { min: 8, max: 12 },
        'Urn': { min: 12, max: 20 },
        'Product Name': { min: 40, max: 90 },
        'Vulgraad': { min: 10, max: 12 },
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
