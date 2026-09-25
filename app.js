const PRODUCTS = [
  { id: "jus1l", name: "Jus de pomme 1 L", savedKgPerUnit: 1.5, savedItemsPerUnit: 10 },
  { id: "petillant75", name: "Pétillant pomme 75 cl", savedKgPerUnit: 1.2, savedItemsPerUnit: 8 },
  { id: "pomme33", name: "Jus pomme 33 cl", savedKgPerUnit: 0.6, savedItemsPerUnit: 4 },
  { id: "tartinadeBetterave", name: "Tartinade betterave", savedKgPerUnit: 2.5, savedItemsPerUnit: 1 },
  { id: "tartinadeTomateOignon", name: "Tartinade tomate-oignon", savedKgPerUnit: 0.34, savedItemsPerUnit: 3 },
  { id: "confitureTomate", name: "Confiture de tomate", savedKgPerUnit: 0.36, savedItemsPerUnit: 3 },
  { id: "tartinadeButternut", name: "Tartinade butternut", savedKgPerUnit: 1, savedItemsPerUnit: 1 },
  { id: "compotePomme", name: "Compote de pomme", savedKgPerUnit: 0.75, savedItemsPerUnit: 5 },
  { id: "ketchupBetterave", name: "Ketchup de betterave", savedKgPerUnit: 2.5, savedItemsPerUnit: 1 },
  { id: "ketchupTomateVerte", name: "Ketchup de tomate verte", savedKgPerUnit: 0.24, savedItemsPerUnit: 2 }
];

const RATIOS = {
  waterLitersPerKg: 300,
  co2KgPerKgSaved: 1.53,
  mealKgEquivalent: 0.45,
  showerLiters: 60,
  carKgCo2PerKm: 0.192,
  kgPerSupportedProducer: 120
};

const PRODUCT_BY_ID = Object.fromEntries(PRODUCTS.map((p) => [p.id, p]));

const ALIASES = [
  { id: "jus1l", codes: ["251077"], needles: ["jus de pomme 1l", "jus de pomme 1 l", "jus pomme 1l"] },
  { id: "petillant75", codes: ["251078"], needles: ["petillant", "petillant pomme", "jus de pomme petillant"] },
  { id: "pomme33", codes: ["251074", "677072"], needles: ["jus de pomme 33", "jus pomme 33", "pomme rhubarbe 33", "rhubarbe 33", "jus de pomme rhubarbe"] },
  { id: "tartinadeBetterave", codes: ["6398384"], needles: ["tartinade betterave"] },
  { id: "tartinadeTomateOignon", codes: ["251081"], needles: ["tartinade tomate", "tartinade tomates oignon"] },
  { id: "confitureTomate", codes: ["2394793"], needles: ["confiture de tomate"] },
  { id: "tartinadeButternut", codes: ["774785"], needles: ["tartinade butternut"] },
  { id: "compotePomme", codes: ["251079"], needles: ["compote de pomme", "compote pomme"] },
  { id: "ketchupBetterave", codes: [], needles: ["ketchup betterave"] },
  { id: "ketchupTomateVerte", codes: [], needles: ["ketchup tomate verte", "ketchup de tomate verte"] }
];

const HEADER_ALIASES = {
  invoiceId: ["identifiant facture"],
  invoiceNumber: ["numero de facture", "n de facture"],
  date: ["date"],
  client: ["tiers", "client", "nom du client"],
  clientId: ["identifiant tiers"],
  siren: ["siren tiers", "numero siren", "siren"],
  label: ["libelle", "libelle produit", "designation"],
  productName: ["libelle produit"],
  productCode: ["code produit", "reference produit"],
  quantity: ["quantite", "qte"],
  unit: ["unite"]
};

let rawLines = [];
let customMapping = loadMapping();
let currentClients = [];

function normalize(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/œ/g, "oe")
    .replace(/[^a-z0-9]+/g, " ")
    .trim()
    .replace(/\s+/g, " ");
}

function loadMapping() {
  try {
    return JSON.parse(localStorage.getItem("topinamour-product-map") || "{}");
  } catch {
    return {};
  }
}

function saveMapping() {
  localStorage.setItem("topinamour-product-map", JSON.stringify(customMapping));
}

function mappingKey(line) {
  return [normalize(line.productCode), normalize(line.productName || line.label)].join("|");
}

function normalizeCode(value) {
  const text = String(value ?? "").trim();
  if (!text) return "";
  const asNum = Number(text);
  if (Number.isFinite(asNum) && asNum > 0) return String(Math.round(asNum));
  return text;
}

function guessProductId(line) {
  const key = mappingKey(line);
  if (customMapping[key]) return customMapping[key];

  const code = normalizeCode(line.productCode);
  const haystack = normalize([line.label, line.productName, line.productRef].filter(Boolean).join(" "));

  for (const alias of ALIASES) {
    if (code && alias.codes.includes(code)) return alias.id;
  }
  for (const alias of ALIASES) {
    if (alias.needles.some((needle) => haystack.includes(needle))) return alias.id;
  }
  return "";
}

function formatNumber(value) {
  return Math.round(value).toLocaleString("fr-FR");
}

function formatSiren(value) {
  const digits = String(value ?? "").replace(/\D/g, "");
  if (!digits) return "";
  const siren = digits.length >= 9 ? digits.slice(0, 9) : digits;
  if (siren.length === 9) {
    return `${siren.slice(0, 3)} ${siren.slice(3, 6)} ${siren.slice(6)}`;
  }
  return siren;
}

function parseExcelDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
  }
  const text = String(value || "").trim();
  const match = text.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (match) {
    const year = match[3].length === 2 ? Number("20" + match[3]) : Number(match[3]);
    return new Date(year, Number(match[2]) - 1, Number(match[1]));
  }
  const iso = new Date(text);
  return Number.isNaN(iso.getTime()) ? null : iso;
}

function findHeaderIndex(headers, aliases) {
  const normalized = headers.map(normalize);
  const exact = normalized.findIndex((header) => aliases.includes(header));
  if (exact >= 0) return exact;
  return normalized.findIndex((header) =>
    aliases.some((alias) => alias.length >= 8 && header.includes(alias))
  );
}

function inspectSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
  if (!rows.length) {
    return { lines: [], missing: ["Tiers", "Libellé", "Quantité"], empty: true };
  }

  const headers = rows[0].map((cell) => String(cell || ""));
  const indexes = {};
  for (const [key, aliases] of Object.entries(HEADER_ALIASES)) {
    indexes[key] = findHeaderIndex(headers, aliases);
  }

  const missing = [];
  if (indexes.client < 0) missing.push("Tiers");
  if (indexes.label < 0 && indexes.productName < 0) missing.push("Libellé");
  if (indexes.quantity < 0) missing.push("Quantité");
  if (missing.length) return { lines: [], missing, empty: false };

  const lines = rows.slice(1).map((row) => {
    const pick = (key) => (indexes[key] >= 0 ? row[indexes[key]] : "");
    return {
      invoiceId: String(pick("invoiceId") || ""),
      invoiceNumber: String(pick("invoiceNumber") || ""),
      date: parseExcelDate(pick("date")),
      client: String(pick("client") || "").trim(),
      clientId: String(pick("clientId") || ""),
      siren: formatSiren(pick("siren")),
      label: String(pick("label") || pick("productName") || "").trim(),
      productName: String(pick("productName") || pick("label") || "").trim(),
      productCode: String(pick("productCode") || "").trim(),
      productRef: String(pick("productName") || ""),
      quantity: Number(String(pick("quantity")).replace(",", ".")) || 0,
      unit: String(pick("unit") || "")
    };
  }).filter((line) => line.client || line.label);

  return { lines, missing: [], empty: false };
}

function parseWorkbook(workbook) {
  if (!workbook || !workbook.SheetNames || !workbook.SheetNames.length) {
    throw new Error("Le fichier Excel ne contient aucune feuille. Exportez de nouveau les factures depuis Pennylane.");
  }

  const preferredName = workbook.SheetNames.find((name) => normalize(name).includes("ligne"));
  const namesToTry = preferredName
    ? [preferredName, ...workbook.SheetNames.filter((name) => name !== preferredName)]
    : workbook.SheetNames;

  let firstMissing = null;

  for (const name of namesToTry) {
    const inspected = inspectSheet(workbook.Sheets[name]);
    if (inspected.lines.length) return inspected.lines;
    if (!firstMissing && inspected.missing.length && !inspected.empty) {
      firstMissing = { sheet: name, missing: inspected.missing };
    }
  }

  if (firstMissing) {
    throw new Error(
      `Le fichier n’est pas au format attendu (feuille « ${firstMissing.sheet} »). ` +
      `Colonnes manquantes : ${firstMissing.missing.join(", ")}. ` +
      "Importez un export Pennylane de factures avec au minimum Tiers, Libellé et Quantité."
    );
  }

  throw new Error(
    "Aucune ligne de facture n’a été trouvée. Vérifiez que le fichier contient la feuille « Lignes de facture » avec des données."
  );
}

function computeImpact(quantityByProductId, unmatchedQty = 0) {
  let savedKg = 0;
  let savedItems = 0;
  let totalProducts = 0;

  for (const [id, qty] of Object.entries(quantityByProductId)) {
    const product = PRODUCT_BY_ID[id];
    if (!product) continue;
    totalProducts += qty;
    savedKg += qty * product.savedKgPerUnit;
    savedItems += qty * product.savedItemsPerUnit;
  }

  const water = savedKg * RATIOS.waterLitersPerKg;
  const co2 = savedKg * RATIOS.co2KgPerKgSaved;
  const producers = savedKg > 0 ? Math.max(1, Math.ceil(savedKg / RATIOS.kgPerSupportedProducer)) : 0;
  const meals = savedKg / RATIOS.mealKgEquivalent;
  const showers = water / RATIOS.showerLiters;
  const carKm = co2 / RATIOS.carKgCo2PerKm;

  return {
    savedKg, savedItems, totalProducts, unmatchedQty, water, co2, producers, meals, showers, carKm
  };
}

function aggregate(lines) {
  const byClient = new Map();

  for (const line of lines) {
    const key = line.clientId || line.client;
    if (!byClient.has(key)) {
      byClient.set(key, {
        key,
        client: line.client,
        siren: line.siren || "",
        invoices: new Set(),
        quantityByProductId: {},
        unmatchedQty: 0,
        unmatched: [],
        invoiceCount: 0,
        minDate: null,
        maxDate: null
      });
    }
    const bucket = byClient.get(key);
    if (!bucket.siren && line.siren) bucket.siren = line.siren;
    if (line.invoiceNumber) bucket.invoices.add(line.invoiceNumber);
    if (line.date && (!bucket.minDate || line.date < bucket.minDate)) bucket.minDate = line.date;
    if (line.date && (!bucket.maxDate || line.date > bucket.maxDate)) bucket.maxDate = line.date;

    const productId = guessProductId(line);
    if (!productId) {
      bucket.unmatchedQty += line.quantity;
      bucket.unmatched.push(line);
      continue;
    }
    bucket.quantityByProductId[productId] = (bucket.quantityByProductId[productId] || 0) + line.quantity;
  }

  return [...byClient.values()].map((bucket) => ({
    ...bucket,
    invoiceCount: bucket.invoices.size,
    impact: computeImpact(bucket.quantityByProductId, bucket.unmatchedQty)
  })).sort((a, b) => a.client.localeCompare(b.client, "fr"));
}

function uniqueUnmapped(lines) {
  const map = new Map();
  for (const line of lines) {
    if (guessProductId(line)) continue;
    const key = mappingKey(line);
    if (!map.has(key)) {
      map.set(key, {
        key,
        label: line.productName || line.label,
        code: line.productCode,
        quantity: 0
      });
    }
    map.get(key).quantity += line.quantity;
  }
  return [...map.values()];
}

function parseInputDate(value, endOfDay = false) {
  if (!value) return null;
  const [year, month, day] = value.split("-").map(Number);
  return endOfDay
    ? new Date(year, month - 1, day, 23, 59, 59, 999)
    : new Date(year, month - 1, day);
}

function toInputDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function filterLines() {
  const query = normalize(document.getElementById("clientFilter").value);
  const from = parseInputDate(document.getElementById("dateFrom").value);
  const to = parseInputDate(document.getElementById("dateTo").value, true);

  return rawLines.filter((line) => {
    if (query && !normalize(line.client).includes(query)) return false;
    if (from && line.date && line.date < from) return false;
    if (to && line.date && line.date > to) return false;
    return true;
  });
}

function renderKpis(clients) {
  const totals = clients.reduce((acc, client) => {
    for (const key of Object.keys(client.impact)) acc[key] = (acc[key] || 0) + client.impact[key];
    return acc;
  }, {});
  if (totals.savedKg > 0) {
    totals.producers = Math.max(1, Math.ceil(totals.savedKg / RATIOS.kgPerSupportedProducer));
  }

  const items = [
    ["savedKg", "kg sauvés"],
    ["savedItems", "fruits & légumes"],
    ["water", "litres d'eau"],
    ["co2", "kg de CO₂"],
    ["producers", "producteurs"],
    ["meals", "repas sauvés"]
  ];
  document.getElementById("kpis").innerHTML = items.map(([key, label]) => `
    <div class="kpi">
      <div class="value">${formatNumber(totals[key] || 0)}</div>
      <div class="label">${label}</div>
    </div>
  `).join("");
}

function productMixItems(client) {
  return Object.entries(client.quantityByProductId)
    .filter(([, qty]) => qty)
    .map(([id, qty]) => `${PRODUCT_BY_ID[id]?.name || id} × ${formatNumber(qty)}`);
}

function productMix(client) {
  return productMixItems(client).join(" · ");
}

function renderTable(clients) {
  const body = document.getElementById("resultsBody");
  body.innerHTML = clients.map((client, index) => `
    <tr>
      <td><strong>${escapeHtml(client.client)}</strong></td>
      <td>${escapeHtml(client.siren || "—")}</td>
      <td class="mix">${productMixItems(client).map((item) => `<div>${escapeHtml(item)}</div>`).join("") || "—"}</td>
      <td>${formatNumber(client.impact.totalProducts)}</td>
      <td>${client.invoiceCount}</td>
      <td>${formatNumber(client.impact.savedKg)}</td>
      <td>${formatNumber(client.impact.savedItems)}</td>
      <td>${formatNumber(client.impact.water)}</td>
      <td>${formatNumber(client.impact.co2)}</td>
      <td>${formatNumber(client.impact.producers)}</td>
      <td>${formatNumber(client.impact.meals)}</td>
      <td>${formatNumber(client.impact.showers)}</td>
      <td>${formatNumber(client.impact.carKm)}</td>
      <td>${client.impact.unmatchedQty
        ? `<span class="badge warn">${formatNumber(client.impact.unmatchedQty)} non mappés</span>`
        : `<span class="badge ok">OK</span>`}</td>
              <td>
                <button class="light" data-cert="${index}">Voir le certificat</button>
              </td>
    </tr>
  `).join("");

  body.querySelectorAll("[data-cert]").forEach((button) => {
    button.addEventListener("click", () => openCertificate(clients[Number(button.dataset.cert)]));
  });
}

function renderUnmapped(lines) {
  const unmapped = uniqueUnmapped(lines);
  const box = document.getElementById("unmappedBox");
  if (!unmapped.length) {
    box.classList.add("hidden");
    box.innerHTML = "";
    return;
  }
  box.classList.remove("hidden");
  const options = PRODUCTS.map((p) => `<option value="${p.id}">${p.name}</option>`).join("");
  box.innerHTML = `
    <strong>Produits non reconnus</strong>
    <p class="note">Associez-les à un produit du calculateur pour qu’ils comptent dans l’impact.</p>
    ${unmapped.map((item) => `
      <div class="mapping-row">
        <div>${escapeHtml(item.label)} ${item.code ? `(code ${escapeHtml(item.code)})` : ""} — ${formatNumber(item.quantity)} unités</div>
        <select data-map="${escapeHtml(item.key)}">
          <option value="">Ignorer pour l’instant</option>
          ${options}
        </select>
      </div>
    `).join("")}
  `;
  box.querySelectorAll("select").forEach((select) => {
    select.addEventListener("change", () => {
      if (select.value) customMapping[select.dataset.map] = select.value;
      else delete customMapping[select.dataset.map];
      saveMapping();
      refresh();
    });
  });
}

function refresh() {
  const filtered = filterLines();
  currentClients = aggregate(filtered);
  document.getElementById("resultsCard").classList.toggle("hidden", !rawLines.length);
  document.getElementById("emptyState").classList.toggle("hidden", Boolean(rawLines.length));
  renderKpis(currentClients);
  renderTable(currentClients);
  renderUnmapped(filtered);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function formatDate(date) {
  if (!date) return "";
  return date.toLocaleDateString("fr-FR");
}

const CERTIFICATE_TIERS = [
  { id: "debutant" },
  { id: "confirme" },
  { id: "heros" }
];

function certificateTier(savedKg) {
  if (savedKg >= 1500) return CERTIFICATE_TIERS[2];
  if (savedKg >= 500) return CERTIFICATE_TIERS[1];
  return CERTIFICATE_TIERS[0];
}

function certificateHtmlOfficial(client, options) {
  const impact = client.impact;
  return `
    <div class="certificate cert-html ${options.className}" id="certificatePage">
      <div class="content">
        <aside class="badge-frame">
          <img class="badge" src="${options.badge}" alt="${options.badgeAlt}">
          <div class="badge-caption">
            <div class="badge-caption-title">TOPI FAMILY</div>
            <div class="badge-caption-tagline">Chaque commande = un sauvetage</div>
          </div>
        </aside>
        <section class="right">
          <img class="logo" src="logo-topinamour-rose.png" alt="Topinamour">
          <h1>CERTIFICAT</h1>
          <div class="subtitle">${options.subtitle}</div>
          <div class="recipient">Décerné à</div>
          <div class="company">${escapeHtml(client.client || "Client")}</div>
          <div class="commitment">
            pour son engagement anti-gaspi :
            <strong>${formatNumber(impact.savedKg)} kg</strong>
            de fruits et légumes hors normes sauvés avec Topinamour.
          </div>
          <div class="thanks">Topinamour vous remercie pour votre engagement.</div>
          <div class="impact">
            <div class="impact-title">VOTRE IMPACT</div>
            <div class="impact-row"><div class="icon">💧</div><div>Eau économisée : ${formatNumber(impact.water)} litres</div></div>
            <div class="impact-row"><div class="icon">🌍</div><div>CO₂ évité : ${formatNumber(impact.co2)} kg</div></div>
            <div class="impact-row"><div class="icon">🍽️</div><div>Équivalent repas sauvés : ${formatNumber(impact.meals)}</div></div>
            <div class="impact-row"><div class="icon">🚿</div><div>Douches économisées : ${formatNumber(impact.showers)}</div></div>
            <div class="impact-row"><div class="icon">🚗</div><div>Kilomètres en voiture évités : ${formatNumber(impact.carKm)} km</div></div>
          </div>
          <div class="bottom">
            <div class="signature">Date ${formatDate(new Date())}</div>
            <div class="signature">L'équipe Topinamour</div>
          </div>
        </section>
      </div>
    </div>
  `;
}

function certificateHtmlDebutant(client) {
  return certificateHtmlOfficial(client, {
    className: "cert-html-debutant",
    badge: "badge-1-sauveur-debutant.png",
    badgeAlt: "Badge Topi Family - Sauveur débutant",
    subtitle: "SAUVEUR DÉBUTANT"
  });
}

function certificateHtmlConfirme(client) {
  return certificateHtmlOfficial(client, {
    className: "cert-html-confirme",
    badge: "badge-2-sauveur-confirme.png",
    badgeAlt: "Badge Topi Family - Sauveur confirmé",
    subtitle: "SAUVEUR CONFIRMÉ"
  });
}

function certificateHtmlHeros(client) {
  return certificateHtmlOfficial(client, {
    className: "cert-html-heros",
    badge: "badge-3-heros-anti-gaspi_2.png",
    badgeAlt: "Badge Topi Family - Héros anti-gaspi",
    subtitle: "HÉROS ANTI-GASPI"
  });
}

function certificateHtml(client) {
  const tier = certificateTier(client.impact.savedKg);
  if (tier.id === "confirme") return certificateHtmlConfirme(client);
  if (tier.id === "heros") return certificateHtmlHeros(client);
  return certificateHtmlDebutant(client);
}

let previewClient = null;
const CERT_PREVIEW_WIDTH = 809;
const CERT_PREVIEW_HEIGHT = 462;
let previewFitObserver = null;

function fitCertificatePreview() {
  const overlay = document.getElementById("certOverlay");
  const preview = document.getElementById("certPreview");
  const cert = preview && preview.querySelector(".certificate");
  if (!cert || overlay.classList.contains("hidden")) return;

  const card = document.getElementById("certCard");
  const actions = card.querySelector(".cert-actions");
  const cardStyle = getComputedStyle(card);
  const padX = parseFloat(cardStyle.paddingLeft) + parseFloat(cardStyle.paddingRight);
  const padY = parseFloat(cardStyle.paddingTop) + parseFloat(cardStyle.paddingBottom);
  const overlayStyle = getComputedStyle(overlay);
  const overlayPadY = parseFloat(overlayStyle.paddingTop) + parseFloat(overlayStyle.paddingBottom);
  const availableWidth = Math.max(120, card.clientWidth - padX);
  const availableHeight = Math.max(80, overlay.clientHeight - overlayPadY - padY - actions.offsetHeight - 8);
  const scale = Math.min(1, availableWidth / CERT_PREVIEW_WIDTH, availableHeight / CERT_PREVIEW_HEIGHT);

  cert.style.transform = `scale(${scale})`;
  preview.style.width = `${CERT_PREVIEW_WIDTH * scale}px`;
  preview.style.height = `${CERT_PREVIEW_HEIGHT * scale}px`;
}

function openCertificate(client) {
  previewClient = client;
  const overlay = document.getElementById("certOverlay");
  overlay.querySelector("#certPreview").innerHTML = certificateHtml(client);
  overlay.classList.remove("hidden");
  requestAnimationFrame(fitCertificatePreview);
  if (!previewFitObserver) {
    previewFitObserver = new ResizeObserver(fitCertificatePreview);
    previewFitObserver.observe(document.getElementById("certCard"));
  }
}

function certificateFilename(client, extension) {
  const target = client || previewClient || {};
  const tier = certificateTier(target.impact ? target.impact.savedKg : 0);
  const name = slugify(target.client);
  return `certificat-${tier.id}-${name}.${extension}`;
}

function mountHiddenCertificate(client) {
  const root = document.getElementById("printRoot");
  root.innerHTML = certificateHtml(client).replace('id="certificatePage"', 'id="certificatePrintPage"');
  return root.querySelector("#certificatePrintPage");
}

function unmountHiddenCertificate() {
  document.getElementById("printRoot").innerHTML = "";
}

function closeCertificate() {
  document.getElementById("certOverlay").classList.add("hidden");
}

function setCertBusy(busy) {
  document.getElementById("downloadCertPdf").disabled = busy;
  document.getElementById("downloadCertImage").disabled = busy;
}

function waitForImages(root) {
  return Promise.all([...root.querySelectorAll("img")].map((img) => {
    if (img.complete) return Promise.resolve();
    return new Promise((resolve) => {
      img.onload = resolve;
      img.onerror = resolve;
    });
  }));
}

async function captureCertificateImage(type = "png", page = null) {
  const ownsPage = !page;
  if (!page) {
    if (!previewClient) throw new Error("Aucun certificat affiché.");
    page = mountHiddenCertificate(previewClient);
  }
  const root = document.getElementById("printRoot");
  const previousRootStyle = root.getAttribute("style");
  try {
    page.style.maxWidth = "none";
    page.style.margin = "0";
    page.style.transform = "none";
    root.style.cssText = "position:fixed;left:0;top:0;width:auto;height:auto;overflow:visible;clip-path:none;opacity:0.01;pointer-events:none;z-index:-1;";
    await waitForImages(page);
    const width = Math.ceil(page.offsetWidth);
    const height = Math.ceil(page.offsetHeight);
    root.style.width = `${width}px`;
    root.style.height = `${height}px`;
    const canvasOptions = {
      scale: 2,
      useCORS: true,
      allowTaint: true,
      backgroundColor: "#ffffff",
      width,
      height,
      windowWidth: width,
      windowHeight: height,
      x: 0,
      y: 0,
      scrollX: 0,
      scrollY: 0,
      onclone: (doc) => {
        const clonedRoot = doc.getElementById("printRoot");
        if (clonedRoot) {
          clonedRoot.style.cssText = `position:absolute;left:0;top:0;width:${width}px;height:${height}px;opacity:1;overflow:visible;clip-path:none;`;
        }
        const cloned = doc.getElementById(page.id);
        if (cloned) {
          cloned.style.maxWidth = "none";
          cloned.style.width = `${width}px`;
          cloned.style.height = `${height}px`;
          cloned.style.margin = "0 auto";
          cloned.style.position = "relative";
          cloned.style.left = "0";
          cloned.style.top = "0";
          cloned.style.right = "auto";
          cloned.style.transform = "none";
          cloned.style.opacity = "1";
          cloned.style.visibility = "visible";
          cloned.style.overflow = "hidden";
        }
      }
    };
    const mime = type === "jpeg" ? "image/jpeg" : "image/png";
    const quality = type === "jpeg" ? 0.98 : 1;
    if (typeof html2canvas === "function") {
      const canvas = await html2canvas(page, canvasOptions);
      return canvas.toDataURL(mime, quality);
    }
    return html2pdf().set({
      margin: 0,
      image: { type, quality },
      jsPDF: { unit: "px", format: [width, height], orientation: width >= height ? "landscape" : "portrait" },
      html2canvas: canvasOptions
    }).from(page).outputImg("datauristring");
  } finally {
    if (previousRootStyle) root.setAttribute("style", previousRootStyle);
    else root.removeAttribute("style");
    if (ownsPage) unmountHiddenCertificate();
  }
}

function dataUrlToBlob(dataUrl) {
  const [header, data] = dataUrl.split(",");
  const mime = header.match(/data:(.*?);/)[1];
  const bytes = atob(data);
  const buffer = new Uint8Array(bytes.length);
  for (let i = 0; i < bytes.length; i += 1) buffer[i] = bytes.charCodeAt(i);
  return new Blob([buffer], { type: mime });
}

async function pdfFromImage(dataUrl, client) {
  const img = new Image();
  await new Promise((resolve, reject) => {
    img.onload = resolve;
    img.onerror = reject;
    img.src = dataUrl;
  });
  const pageW = 297;
  const pageH = Math.round((pageW * img.height / img.width) * 10) / 10;
  const probe = document.createElement("div");
  probe.style.cssText = "position:fixed;left:-9999px;top:0;width:1px;height:1px;opacity:0;pointer-events:none;";
  document.body.appendChild(probe);
  try {
    const worker = html2pdf().set({
      margin: 0,
      filename: certificateFilename(client, "pdf"),
      jsPDF: { unit: "mm", format: [pageW, pageH], orientation: pageW >= pageH ? "landscape" : "portrait" }
    }).from(probe).toPdf();
    const pdf = await worker.get("pdf");
    pdf.addImage(dataUrl, "JPEG", 0, 0, pageW, pageH);
    const total = pdf.internal.getNumberOfPages();
    for (let page = total; page > 1; page -= 1) pdf.deletePage(page);
    await worker.save();
  } finally {
    probe.remove();
  }
}

async function exportClientCertificatePdf(client) {
  const page = mountHiddenCertificate(client);
  try {
    const dataUrl = await captureCertificateImage("jpeg", page);
    await pdfFromImage(dataUrl, client);
  } finally {
    unmountHiddenCertificate();
  }
}

async function downloadCertificatePdf() {
  if (!previewClient) return;
  setCertBusy(true);
  try {
    const dataUrl = await captureCertificateImage("jpeg");
    await pdfFromImage(dataUrl, previewClient);
  } finally {
    setCertBusy(false);
  }
}

function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function downloadCertificateImage() {
  if (!previewClient) return;
  setCertBusy(true);
  try {
    const dataUrl = await captureCertificateImage();
    triggerDownload(dataUrlToBlob(dataUrl), certificateFilename(previewClient, "png"));
  } finally {
    setCertBusy(false);
  }
}

async function downloadAllCertificates() {
  const button = document.getElementById("allCertsBtn");
  button.disabled = true;
  try {
    for (const client of currentClients) {
      await exportClientCertificatePdf(client);
      await new Promise((resolve) => setTimeout(resolve, 400));
    }
  } finally {
    button.disabled = false;
  }
}

function slugify(value) {
  return normalize(value).replace(/\s+/g, "-").slice(0, 60) || "client";
}

function exportCsv() {
  const headers = [
    "Client", "SIREN", "Détail produits", "Produits", "Factures", "kg sauvés", "Fruits et légumes",
    "Litres d'eau", "kg CO2", "Producteurs", "Repas", "Douches", "km voiture", "Quantités non reconnues"
  ];
  const rows = currentClients.map((c) => [
    c.client, c.siren || "", productMix(c), Math.round(c.impact.totalProducts), c.invoiceCount, Math.round(c.impact.savedKg),
    Math.round(c.impact.savedItems), Math.round(c.impact.water), Math.round(c.impact.co2),
    Math.round(c.impact.producers), Math.round(c.impact.meals), Math.round(c.impact.showers),
    Math.round(c.impact.carKm), Math.round(c.impact.unmatchedQty)
  ]);
  const sheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet, "Impact clients");
  XLSX.writeFile(wb, "impact-topinamour-clients.xlsx");
}

const ALLOWED_EXTENSIONS = [".xlsx", ".xls", ".csv"];

function fileExtension(name) {
  const match = String(name || "").toLowerCase().match(/\.[a-z0-9]+$/);
  return match ? match[0] : "";
}

function isAllowedFile(file) {
  return ALLOWED_EXTENSIONS.includes(fileExtension(file && file.name));
}

function showFileError(message) {
  const box = document.getElementById("fileError");
  const dropzone = document.getElementById("dropzone");
  box.textContent = message;
  box.classList.remove("hidden");
  dropzone.classList.add("has-error");
}

function clearFileError() {
  const box = document.getElementById("fileError");
  const dropzone = document.getElementById("dropzone");
  box.textContent = "";
  box.classList.add("hidden");
  dropzone.classList.remove("has-error");
}

function readErrorMessage(error) {
  const text = String(error && error.message ? error.message : error || "");
  if (/password|encrypt/i.test(text)) {
    return "Ce fichier Excel est protégé par un mot de passe. Exportez-le de nouveau depuis Pennylane sans protection.";
  }
  if (/cfb|ole|unsupported|corrupt|zip|central directory|invalid/i.test(text)) {
    return "Le fichier n’a pas pu être lu. Vérifiez qu’il n’est pas corrompu et qu’il s’agit bien d’un export Excel Pennylane.";
  }
  return text || "Impossible de lire ce fichier.";
}

async function handleFile(file) {
  clearFileError();

  if (!file) {
    showFileError("Aucun fichier n’a été sélectionné.");
    return;
  }
  if (!file.size) {
    showFileError(`Le fichier « ${file.name} » est vide.`);
    return;
  }
  if (!isAllowedFile(file)) {
    const ext = fileExtension(file.name) || "inconnu";
    showFileError(
      `Le type de fichier « ${ext} » n’est pas pris en charge. ` +
      "Importez un export Pennylane au format Excel (.xlsx ou .xls) ou CSV."
    );
    return;
  }

  let workbook;
  try {
    const data = await file.arrayBuffer();
    workbook = XLSX.read(data, { type: "array", cellDates: true });
  } catch (error) {
    showFileError(readErrorMessage(error));
    return;
  }

  try {
    rawLines = parseWorkbook(workbook);
  } catch (error) {
    showFileError(readErrorMessage(error));
    return;
  }

  document.getElementById("fileName").textContent = `${file.name} — ${rawLines.length} lignes`;
  const dates = rawLines.map((l) => l.date).filter(Boolean).sort((a, b) => a - b);
  if (dates.length) {
    document.getElementById("dateFrom").value = toInputDate(dates[0]);
    document.getElementById("dateTo").value = toInputDate(dates[dates.length - 1]);
  }
  refresh();
}

function wireUi() {
  const dropzone = document.getElementById("dropzone");
  const fileInput = document.getElementById("fileInput");

  dropzone.addEventListener("click", () => fileInput.click());
  dropzone.addEventListener("dragover", (event) => {
    event.preventDefault();
    dropzone.classList.add("dragover");
  });
  dropzone.addEventListener("dragleave", () => dropzone.classList.remove("dragover"));
  dropzone.addEventListener("drop", (event) => {
    event.preventDefault();
    dropzone.classList.remove("dragover");
    const files = [...(event.dataTransfer.files || [])];
    if (!files.length) {
      showFileError("Aucun fichier n’a été déposé.");
      return;
    }
    if (files.length > 1) {
      showFileError("Déposez un seul fichier à la fois : l’export Pennylane des factures.");
      return;
    }
    handleFile(files[0]);
  });
  fileInput.addEventListener("change", () => {
    if (fileInput.files[0]) handleFile(fileInput.files[0]);
    fileInput.value = "";
  });

  ["clientFilter", "dateFrom", "dateTo"].forEach((id) => {
    document.getElementById(id).addEventListener("input", refresh);
  });
  document.getElementById("exportBtn").addEventListener("click", exportCsv);
  document.getElementById("allCertsBtn").addEventListener("click", downloadAllCertificates);
  document.getElementById("closeCert").addEventListener("click", closeCertificate);
  document.getElementById("downloadCertPdf").addEventListener("click", downloadCertificatePdf);
  document.getElementById("downloadCertImage").addEventListener("click", downloadCertificateImage);
  document.getElementById("certOverlay").addEventListener("click", (event) => {
    if (event.target.id === "certOverlay") closeCertificate();
  });
  window.addEventListener("resize", fitCertificatePreview);
}

wireUi();
