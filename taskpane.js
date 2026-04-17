/* =============================================
   PowerPoint Icon Scraper Add-in
   Source : Iconify API (gratuite, sans clé)
   https://iconify.design/
   ============================================= */

const ICONIFY_API = "https://api.iconify.design";
const RESULTS_LIMIT = 60;

let selectedIcon = null; // { prefix, name }

// ── Init Office.js ──────────────────────────
Office.onReady(() => {
  document.getElementById("search-btn").addEventListener("click", doSearch);
  document.getElementById("search-input").addEventListener("keydown", (e) => {
    if (e.key === "Enter") doSearch();
  });
  document.getElementById("insert-btn").addEventListener("click", insertIcon);
});

// ── Recherche ───────────────────────────────
async function doSearch() {
  const query = document.getElementById("search-input").value.trim();
  const filterSet = document.getElementById("filter-set").value;

  if (!query) return setStatus("Saisis un mot-clé pour chercher.");

  setStatus("Recherche en cours…");
  clearResults();
  hidePreview();

  try {
    // Construction de l'URL Iconify search
    const params = new URLSearchParams({
      query,
      limit: RESULTS_LIMIT,
    });
    if (filterSet) params.set("prefix", filterSet);

    const res = await fetch(`${ICONIFY_API}/search?${params}`);
    if (!res.ok) throw new Error(`Erreur API : ${res.status}`);

    const data = await res.json();

    if (!data.icons || data.icons.length === 0) {
      return setStatus("Aucun résultat. Essaie un autre mot-clé.");
    }

    setStatus(`${data.icons.length} icône(s) trouvée(s)`);
    renderResults(data.icons);
  } catch (err) {
    setStatus(`Erreur : ${err.message}`, true);
  }
}

// ── Affichage des résultats ──────────────────
function renderResults(icons) {
  const grid = document.getElementById("results");

  icons.forEach((fullName) => {
    // fullName = "mdi:home" ou "logos:github"
    const [prefix, ...rest] = fullName.split(":");
    const name = rest.join(":");

    const card = document.createElement("div");
    card.className = "icon-card";
    card.title = fullName;

    const img = document.createElement("img");
    img.src = `${ICONIFY_API}/${prefix}/${name}.svg`;
    img.alt = name;
    img.loading = "lazy";
    img.onerror = () => { card.style.display = "none"; }; // cache les broken

    const label = document.createElement("span");
    label.className = "icon-label";
    label.textContent = name;

    card.appendChild(img);
    card.appendChild(label);

    card.addEventListener("click", () => selectIcon(card, prefix, name, img.src));
    grid.appendChild(card);
  });
}

// ── Sélection d'une icône ───────────────────
function selectIcon(card, prefix, name, svgUrl) {
  // Désélectionne l'ancienne carte
  document.querySelectorAll(".icon-card.selected").forEach((c) =>
    c.classList.remove("selected")
  );
  card.classList.add("selected");

  selectedIcon = { prefix, name };

  // Met à jour la barre de prévisualisation
  document.getElementById("preview-icon").innerHTML =
    `<img src="${svgUrl}" alt="${name}" />`;
  document.getElementById("preview-name").textContent = name;
  document.getElementById("preview-set").textContent = prefix;
  showPreview();
}

// ── Insertion dans PowerPoint ────────────────
async function insertIcon() {
  if (!selectedIcon) return;

  const size = parseInt(document.getElementById("size-select").value, 10);
  const { prefix, name } = selectedIcon;
  const btn = document.getElementById("insert-btn");

  btn.disabled = true;
  btn.textContent = "Insertion…";

  try {
    // Récupère le SVG en texte
    const res = await fetch(`${ICONIFY_API}/${prefix}/${name}.svg?width=${size}&height=${size}&color=%23000000`);
    if (!res.ok) throw new Error(`Impossible de charger le SVG (${res.status})`);
    const svgText = await res.text();

    // Encode en base64
    const base64 = btoa(unescape(encodeURIComponent(svgText)));

    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle); // placeholder
      await context.sync();
    });

    // Méthode principale : insertion via insertSvg (disponible Office 2021+)
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);

      // Taille et position en points (1 pt = 1/72 pouce)
      const sizeInPoints = (size / 96) * 72;
      const left = 100;
      const top = 100;

      slide.shapes.addSvgImage(
        svgText,
        {
          left,
          top,
          width: sizeInPoints,
          height: sizeInPoints,
        }
      );

      await context.sync();
    });

    setStatus(`✅ "${name}" inséré dans la slide !`);
  } catch (err) {
    // Fallback : insertion via l'API copiePicture si addSvgImage non dispo
    try {
      await insertAsPng(prefix, name, size);
    } catch (fallbackErr) {
      setStatus(`Erreur d'insertion : ${err.message}`, true);
    }
  } finally {
    btn.disabled = false;
    btn.textContent = "Insérer dans la slide";
  }
}

// ── Fallback PNG ─────────────────────────────
async function insertAsPng(prefix, name, size) {
  // Iconify peut aussi retourner un PNG
  const pngUrl = `${ICONIFY_API}/${prefix}/${name}.svg?width=${size}&height=${size}`;

  // On crée un canvas pour convertir SVG → PNG base64
  const img = new Image();
  img.crossOrigin = "anonymous";

  await new Promise((resolve, reject) => {
    img.onload = resolve;
    img.onerror = reject;
    img.src = pngUrl;
  });

  const canvas = document.createElement("canvas");
  canvas.width = size;
  canvas.height = size;
  const ctx = canvas.getContext("2d");
  ctx.drawImage(img, 0, 0, size, size);
  const dataUrl = canvas.toDataURL("image/png");
  const base64 = dataUrl.split(",")[1];

  await PowerPoint.run(async (context) => {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const sizeInPoints = (size / 96) * 72;

    slide.shapes.addImage(base64, {
      left: 100,
      top: 100,
      width: sizeInPoints,
      height: sizeInPoints,
    });

    await context.sync();
  });

  setStatus(`✅ "${name}" inséré (format PNG) dans la slide !`);
}

// ── Helpers UI ───────────────────────────────
function setStatus(msg, isError = false) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = isError ? "status error" : "status";
}

function clearResults() {
  document.getElementById("results").innerHTML = "";
}

function showPreview() {
  document.getElementById("preview-bar").classList.remove("hidden");
}

function hidePreview() {
  document.getElementById("preview-bar").classList.add("hidden");
  selectedIcon = null;
}
