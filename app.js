(() => {
  "use strict";

  const CONFIG = window.FIELD_MAP_CONFIG;
  const COLUMNS = CONFIG.COLUMNS;
  const REQUIRED_HEADERS = [
    COLUMNS.id,
    COLUMNS.taxon,
    COLUMNS.recordType,
    COLUMNS.latitude,
    COLUMNS.longitude
  ];

  const RECORD_TYPE_PRIORITY = {
    type: 50,
    synonymType: 40,
    thisStudy: 30,
    literature: 20,
    unknown: 0
  };

  const DETAILED_MAP_ATTRIBUTION = "Imagery © <a href=\"https://www.esri.com/en-us/home\" target=\"_blank\" rel=\"noopener\">Esri</a>";
  const GSI_TILE_ATTRIBUTION = "<a href='https://maps.gsi.go.jp/development/ichiran.html' target='_blank' rel='noopener'>地理院タイル</a>";

  const TAXON_COLOR_CANDIDATES = [
    { name: "blue", color: "#005AB5" },       // 青
    { name: "red", color: "#DC3220" },        // 赤
    { name: "green", color: "#009E73" },      // 緑
    { name: "purple", color: "#6A3D9A" },     // 紫
    { name: "orange", color: "#E69F00" },     // 橙
    { name: "teal", color: "#00A6A6" },       // 青緑
    { name: "brown", color: "#8B4513" },      // 茶
    { name: "magenta", color: "#E7298A" },    // 赤紫
    { name: "black", color: "#000000" },      // 黒
    { name: "olive", color: "#7A8F00" },      // オリーブ
    { name: "navy", color: "#003F5C" },       // 濃紺
    { name: "pink", color: "#F781BF" },       // 桃
    { name: "mustard", color: "#B8860B" },    // 黄土
    { name: "darkgreen", color: "#006400" },  // 濃緑
    { name: "cyan", color: "#00BFFF" },       // シアン
    { name: "gray", color: "#666666" }        // 灰
  ];

  const TAXON_COLOR_PALETTE = TAXON_COLOR_CANDIDATES.map((entry) => entry.color);

  const state = {
    map: null,
    markers: [],
    rows: [],
    markerGroups: [],
    activePopup: null,
    currentPopupIndex: 0,
    currentPopupRecords: [],
    currentAnchor: null,
    currentShowAbove: null,
    isZooming: false,
    lastZoomRebuildAt: 0,
    hasLoadedOnce: false,
    taxonColors: new Map(),
    datasets: [],
    activeDatasetId: null
  };

  const appEl = document.getElementById("app");
  const statusEl = document.getElementById("status");
  const titleEl = document.getElementById("title");
  const topbarToggleButton = document.getElementById("topbar-toggle-button");
  const topbarOpenButton = document.getElementById("topbar-open-button");
  const reloadButton = document.getElementById("reload-button");
  const taxonLegendEl = document.getElementById("taxon-legend");
  const basemapSelect = document.getElementById("basemap-select");
  const datasetSelect = document.getElementById("dataset-select");
  const datasetAddButton = document.getElementById("dataset-add-button");
  const datasetEditButton = document.getElementById("dataset-edit-button");
  const datasetDeleteButton = document.getElementById("dataset-delete-button");
  const datasetModal = document.getElementById("dataset-modal");
  const datasetForm = document.getElementById("dataset-form");
  const datasetModalTitle = document.getElementById("dataset-modal-title");
  const datasetModalCloseButton = document.getElementById("dataset-modal-close");
  const datasetModalCancelButton = document.getElementById("dataset-modal-cancel");
  const datasetModalSubmitButton = document.getElementById("dataset-modal-submit");
  const datasetNameInput = document.getElementById("dataset-name-input");
  const datasetUrlInput = document.getElementById("dataset-url-input");
  const datasetSheetInput = document.getElementById("dataset-sheet-input");
  const datasetModalError = document.getElementById("dataset-modal-error");
  const DATASETS_STORAGE_KEY = CONFIG.LOCAL_DATASETS_STORAGE_KEY || "fieldMap.localDatasets.v1";
  const ACTIVE_DATASET_STORAGE_KEY = CONFIG.ACTIVE_DATASET_STORAGE_KEY || "fieldMap.activeDatasetId.v1";
  const TOPBAR_COLLAPSED_STORAGE_KEY = CONFIG.TOPBAR_COLLAPSED_STORAGE_KEY || "fieldMap.topbarCollapsed.v1";
  titleEl.textContent = CONFIG.TITLE || "調査用地図";

  function setStatus(message) {
    statusEl.textContent = message;
  }

  function resizeMapSoon() {
    if (!state.map) return;
    window.requestAnimationFrame(() => {
      state.map.resize();
      window.setTimeout(() => {
        if (state.map) state.map.resize();
      }, 160);
    });
  }

  function setTopbarCollapsed(collapsed, persist = true) {
    if (!appEl || !topbarOpenButton) return;
    appEl.classList.toggle("topbar-collapsed", collapsed);
    topbarOpenButton.hidden = !collapsed;
    if (topbarToggleButton) {
      topbarToggleButton.setAttribute("aria-expanded", collapsed ? "false" : "true");
    }
    topbarOpenButton.setAttribute("aria-expanded", collapsed ? "false" : "true");
    if (persist) {
      window.localStorage.setItem(TOPBAR_COLLAPSED_STORAGE_KEY, collapsed ? "1" : "0");
    }
    resizeMapSoon();
  }

  function loadTopbarCollapsed() {
    return window.localStorage.getItem(TOPBAR_COLLAPSED_STORAGE_KEY) === "1";
  }

  function loadStoredDatasets() {
    try {
      const parsed = JSON.parse(window.localStorage.getItem(DATASETS_STORAGE_KEY) || "[]");
      if (!Array.isArray(parsed)) return [];
      return parsed
        .filter((dataset) => dataset && typeof dataset === "object")
        .map((dataset) => ({
          id: normalizeText(dataset.id),
          name: normalizeText(dataset.name),
          sourceUrl: normalizeText(dataset.sourceUrl),
          spreadsheetId: normalizeText(dataset.spreadsheetId),
          sheetName: normalizeText(dataset.sheetName) || CONFIG.DEFAULT_SHEET_NAME || "シート1",
          csvUrl: normalizeText(dataset.csvUrl),
          createdAt: normalizeText(dataset.createdAt),
          updatedAt: normalizeText(dataset.updatedAt)
        }))
        .filter((dataset) => dataset.id && dataset.name && (dataset.spreadsheetId || dataset.csvUrl));
    } catch (error) {
      console.warn("登録済み調査名を読み取れませんでした。", error);
      return [];
    }
  }

  function saveStoredDatasets() {
    window.localStorage.setItem(DATASETS_STORAGE_KEY, JSON.stringify(state.datasets));
  }

  function loadActiveDatasetId() {
    return normalizeText(window.localStorage.getItem(ACTIVE_DATASET_STORAGE_KEY));
  }

  function saveActiveDatasetId(id) {
    if (id) {
      window.localStorage.setItem(ACTIVE_DATASET_STORAGE_KEY, id);
    } else {
      window.localStorage.removeItem(ACTIVE_DATASET_STORAGE_KEY);
    }
  }

  function getActiveDataset() {
    return state.datasets.find((dataset) => dataset.id === state.activeDatasetId) || null;
  }

  function makeDatasetId(name) {
    const base = normalizeText(name)
      .toLowerCase()
      .replace(/[^a-z0-9_-]+/gi, "-")
      .replace(/^-+|-+$/g, "") || "dataset";
    let candidate = `${base}-${Date.now().toString(36)}`;
    let counter = 1;
    while (state.datasets.some((dataset) => dataset.id === candidate)) {
      candidate = `${base}-${Date.now().toString(36)}-${counter}`;
      counter += 1;
    }
    return candidate;
  }

  function extractSpreadsheetId(input) {
    const text = normalizeText(input);
    const urlMatch = text.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    if (urlMatch) return urlMatch[1];
    if (/^[a-zA-Z0-9_-]{20,}$/.test(text)) return text;
    return "";
  }

  function normalizeDatasetInput(nameInput, urlInput, sheetNameInput, existing = null) {
    const name = normalizeText(nameInput);
    const sourceUrl = normalizeText(urlInput);
    const sheetName = normalizeText(sheetNameInput) || CONFIG.DEFAULT_SHEET_NAME || "シート1";

    if (!name) throw new Error("調査名を入力してください。");
    if (!sourceUrl) throw new Error("GoogleスプレッドシートURLを入力してください。");

    const spreadsheetId = extractSpreadsheetId(sourceUrl);
    const isHttpUrl = /^https?:\/\//i.test(sourceUrl);
    if (!spreadsheetId && !isHttpUrl) {
      throw new Error("GoogleスプレッドシートURL、スプレッドシートID、または公開CSV URLを入力してください。");
    }

    const now = new Date().toISOString();
    return {
      id: existing?.id || makeDatasetId(name),
      name,
      sourceUrl,
      spreadsheetId,
      sheetName,
      csvUrl: spreadsheetId ? "" : sourceUrl,
      createdAt: existing?.createdAt || now,
      updatedAt: now
    };
  }

  let datasetModalResolve = null;
  let datasetModalExisting = null;

  function setDatasetModalError(message) {
    if (!datasetModalError) return;
    const text = normalizeText(message);
    datasetModalError.textContent = text;
    datasetModalError.hidden = !text;
  }

  function closeDatasetModal(result = null) {
    if (!datasetModal) return;
    datasetModal.hidden = true;
    document.body.classList.remove("modal-open");
    setDatasetModalError("");

    if (datasetModalResolve) {
      const resolve = datasetModalResolve;
      datasetModalResolve = null;
      datasetModalExisting = null;
      resolve(result);
    }
  }

  function openDatasetModal(existing = null) {
    if (!datasetModal || !datasetForm || !datasetNameInput || !datasetUrlInput || !datasetSheetInput) {
      return Promise.resolve(null);
    }

    datasetModalExisting = existing;
    datasetModalTitle.textContent = existing ? "調査名を編集" : "調査名を追加";
    datasetModalSubmitButton.textContent = existing ? "保存" : "登録";
    datasetNameInput.value = existing?.name || "";
    datasetUrlInput.value = existing?.sourceUrl || existing?.spreadsheetId || existing?.csvUrl || "";
    datasetSheetInput.value = existing?.sheetName || CONFIG.DEFAULT_SHEET_NAME || "シート1";
    setDatasetModalError("");

    datasetModal.hidden = false;
    document.body.classList.add("modal-open");
    window.setTimeout(() => datasetNameInput.focus(), 0);

    return new Promise((resolve) => {
      datasetModalResolve = resolve;
    });
  }

  function submitDatasetModal() {
    try {
      const dataset = normalizeDatasetInput(
        datasetNameInput?.value || "",
        datasetUrlInput?.value || "",
        datasetSheetInput?.value || "",
        datasetModalExisting
      );
      closeDatasetModal(dataset);
    } catch (error) {
      setDatasetModalError(error.message);
    }
  }

  function renderDatasetControls() {
    if (!datasetSelect) return;

    const activeExists = state.datasets.some((dataset) => dataset.id === state.activeDatasetId);
    if (!activeExists) {
      state.activeDatasetId = state.datasets[0]?.id || "";
      saveActiveDatasetId(state.activeDatasetId);
    }

    datasetSelect.innerHTML = "";
    if (state.datasets.length === 0) {
      const option = document.createElement("option");
      option.value = "";
      option.textContent = "未登録";
      datasetSelect.appendChild(option);
      datasetSelect.disabled = true;
      if (reloadButton) reloadButton.disabled = true;
      if (datasetEditButton) datasetEditButton.disabled = true;
      if (datasetDeleteButton) datasetDeleteButton.disabled = true;
      return;
    }

    state.datasets.forEach((dataset) => {
      const option = document.createElement("option");
      option.value = dataset.id;
      option.textContent = dataset.name;
      datasetSelect.appendChild(option);
    });

    datasetSelect.value = state.activeDatasetId;
    datasetSelect.disabled = false;
    if (reloadButton) reloadButton.disabled = false;
    if (datasetEditButton) datasetEditButton.disabled = false;
    if (datasetDeleteButton) datasetDeleteButton.disabled = false;
  }

  function resetMapDataForDatasetChange() {
    closePopup();
    clearMarkers();
    state.rows = [];
    state.markerGroups = [];
    state.taxonColors = new Map();
    if (taxonLegendEl) taxonLegendEl.innerHTML = "";
  }

  async function addDataset() {
    const dataset = await openDatasetModal();
    if (!dataset) return;
    state.datasets.push(dataset);
    state.activeDatasetId = dataset.id;
    saveStoredDatasets();
    saveActiveDatasetId(dataset.id);
    renderDatasetControls();
    resetMapDataForDatasetChange();
    reloadData();
  }

  async function editDataset() {
    const current = getActiveDataset();
    if (!current) return;
    const updated = await openDatasetModal(current);
    if (!updated) return;
    state.datasets = state.datasets.map((dataset) => dataset.id === current.id ? updated : dataset);
    state.activeDatasetId = updated.id;
    saveStoredDatasets();
    saveActiveDatasetId(updated.id);
    renderDatasetControls();
    resetMapDataForDatasetChange();
    reloadData();
  }

  function deleteDataset() {
    const current = getActiveDataset();
    if (!current) return;
    const ok = window.confirm(`「${current.name}」をこの端末の登録一覧から削除しますか？
スプレッドシート本体は削除されません。`);
    if (!ok) return;
    state.datasets = state.datasets.filter((dataset) => dataset.id !== current.id);
    state.activeDatasetId = state.datasets[0]?.id || "";
    saveStoredDatasets();
    saveActiveDatasetId(state.activeDatasetId);
    renderDatasetControls();
    resetMapDataForDatasetChange();
    reloadData();
  }

  function initializeDatasetState() {
    state.datasets = loadStoredDatasets();
    const storedActiveId = loadActiveDatasetId();
    state.activeDatasetId = state.datasets.some((dataset) => dataset.id === storedActiveId)
      ? storedActiveId
      : (state.datasets[0]?.id || "");
    saveActiveDatasetId(state.activeDatasetId);
    renderDatasetControls();
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }


  function isHigherTaxonNotation(text) {
    return /\bord\.|\bfam\.|\bgen\./.test(text);
  }

  function formatScientificNameHTML(scientificName) {
    const raw = normalizeText(scientificName);
    if (!raw) return "—";

    const escaped = escapeHtml(raw)
      .replace(/\(/g, '<span class="non-italic">(</span>')
      .replace(/\)/g, '<span class="non-italic">)</span>');

    if (isHigherTaxonNotation(escaped)) {
      return `<span class="non-italic">${escaped}</span>`;
    }

    if (/ sp\./.test(escaped)) {
      const [beforeSp, afterSp = ""] = escaped.split(/ sp\./, 2);
      const italicPart = beforeSp
        .trim()
        .split(/\s+/)
        .filter(Boolean)
        .map((word) => ["cf.", "aff."].includes(word)
          ? `<span class="non-italic">${word}</span>`
          : `<i>${word}</i>`)
        .join(" ");
      const nonItalicSp = `<span class="non-italic"> sp.${afterSp}</span>`;
      return `${italicPart}${italicPart ? " " : ""}${nonItalicSp}`.trim();
    }

    return escaped
      .split(/\s+/)
      .filter(Boolean)
      .map((part) => ["cf.", "aff."].includes(part)
        ? `<span class="non-italic">${part}</span>`
        : `<i>${part}</i>`)
      .join(" ");
  }

  function formatTaxonHTML(taxon) {
    const raw = normalizeText(taxon);
    if (!raw) return "—";

    if (!raw.includes(" / ")) {
      return formatScientificNameHTML(raw);
    }

    const [japaneseName, ...rest] = raw.split(" / ");
    const scientificName = rest.join(" / ").trim();
    if (!scientificName) return escapeHtml(raw);
    return `${escapeHtml(japaneseName.trim())} / ${formatScientificNameHTML(scientificName)}`;
  }

  function normalizeText(value) {
    return String(value ?? "").trim();
  }

  function displayText(value) {
    const text = normalizeText(value);
    return text || "—";
  }

  function parseNumber(value) {
    if (value === null || value === undefined) return NaN;
    if (typeof value === "number") return value;
    const normalized = String(value).trim().replace(/,/g, "");
    if (!normalized) return NaN;
    return Number(normalized);
  }

  function formatCoordinate(value) {
    const number = Number(value);
    if (!Number.isFinite(number)) return "—";
    return number.toFixed(6);
  }

  function hashString(text) {
    let hash = 2166136261;
    for (let i = 0; i < text.length; i += 1) {
      hash ^= text.charCodeAt(i);
      hash = Math.imul(hash, 16777619);
    }
    return hash >>> 0;
  }

  function taxonKey(value) {
    return normalizeText(value) || "同定未入力";
  }

  function colorForTaxon(taxon) {
    const text = taxonKey(taxon);
    return state.taxonColors.get(text) || fallbackColorForTaxon(text);
  }

  function fallbackColorForTaxon(taxon) {
    return TAXON_COLOR_PALETTE[hashString(taxonKey(taxon)) % TAXON_COLOR_PALETTE.length];
  }

  function assignTaxonColors(records) {
    const taxa = [...new Set(records.map((record) => taxonKey(record.taxon)))]
      .sort((a, b) => a.localeCompare(b, "ja"));

    const adjacency = buildTaxonAdjacency(records, taxa);
    const counts = new Map();
    records.forEach((record) => {
      const taxon = taxonKey(record.taxon);
      counts.set(taxon, (counts.get(taxon) || 0) + 1);
    });

    const order = taxa.sort((a, b) => {
      const degreeA = weightedDegree(adjacency.get(a));
      const degreeB = weightedDegree(adjacency.get(b));
      if (degreeB !== degreeA) return degreeB - degreeA;

      const countA = counts.get(a) || 0;
      const countB = counts.get(b) || 0;
      if (countB !== countA) return countB - countA;

      return a.localeCompare(b, "ja");
    });

    const assigned = new Map();
    const assignedGroups = new Map();
    const groupUsage = new Map();
    const selectedColorGroups = selectDistinctColorGroups(Math.max(order.length, 1));

    order.forEach((taxon) => {
      const neighbors = adjacency.get(taxon) || new Map();
      const paletteGroups = selectedColorGroups;
      const forceUniqueGlobalColors = paletteGroups.length >= order.length;
      let bestGroup = paletteGroups[0] || TAXON_COLOR_CANDIDATES[0];
      let bestScore = Number.NEGATIVE_INFINITY;

      paletteGroups.forEach((candidate, paletteIndex) => {
        let score = 0;

        // 近接関係のある同定で同じ色グループを使うことを強く避けます。
        neighbors.forEach((weight, neighborTaxon) => {
          const neighborColor = assigned.get(neighborTaxon);
          const neighborGroup = assignedGroups.get(neighborTaxon);
          if (!neighborColor || !neighborGroup) return;

          if (neighborGroup === candidate.name) {
            score -= weight * 100000;
          } else {
            score += weight * colorDistance(candidate.color, neighborColor) * 12;
          }
        });

        // 近接していない同定でも、すでに使われている色からできるだけ離します。
        const assignedColors = [...assigned.values()];
        if (assignedColors.length > 0) {
          const globalMinDistance = minColorDistance(candidate.color, assignedColors);
          score += globalMinDistance * 3;
          if (globalMinDistance < 35) score -= (35 - globalMinDistance) * 120;
        } else {
          score += 1000;
        }

        // 同定数が候補色数以下のときは、全体として色の再利用を禁止します。
        // 8種程度で同じ色・ほぼ同じ色が出ることを防ぐためです。
        const usageCount = groupUsage.get(candidate.name) || 0;
        if (forceUniqueGlobalColors && usageCount > 0) {
          score -= 1000000;
        }

        // 色数を超えた場合の再利用は許容しますが、なるべく使用回数が少ない色を選びます。
        score -= usageCount * 300;

        // 同点時に結果が揺れないよう、パレット順をわずかに優先します。
        score -= paletteIndex * 0.001;

        if (score > bestScore) {
          bestScore = score;
          bestGroup = candidate;
        }
      });

      assigned.set(taxon, bestGroup.color);
      assignedGroups.set(taxon, bestGroup.name);
      groupUsage.set(bestGroup.name, (groupUsage.get(bestGroup.name) || 0) + 1);
    });

    state.taxonColors = assigned;
  }

  function selectDistinctColorGroups(requiredCount) {
    const count = Math.max(1, Number(requiredCount) || 1);
    const candidates = [...TAXON_COLOR_CANDIDATES];
    const selected = [];

    // 最初の色は視認性が高く、地図上で埋もれにくい青に固定します。
    const first = candidates.find((entry) => entry.name === "blue") || candidates[0];
    if (first) {
      selected.push(first);
    }

    while (selected.length < count && selected.length < candidates.length) {
      let best = null;
      let bestScore = Number.NEGATIVE_INFINITY;

      candidates.forEach((candidate, index) => {
        if (selected.some((entry) => entry.name === candidate.name)) return;

        const selectedColors = selected.map((entry) => entry.color);
        const minDistance = minColorDistance(candidate.color, selectedColors);
        const averageDistance = averageColorDistance(candidate.color, selectedColors);
        const contrastPenalty = mapVisibilityPenalty(candidate.color);

        // 最小距離を最重視します。これにより「青と空色」「紫と赤紫」のような
        // 近い色が同時に早い段階で選ばれにくくなります。
        const score = (minDistance * 100) + (averageDistance * 8) - contrastPenalty - (index * 0.001);

        if (score > bestScore) {
          bestScore = score;
          best = candidate;
        }
      });

      if (!best) break;
      selected.push(best);
    }

    // 候補数以内なら、ここで選ばれた「互いに離れた色」だけを使います。
    // これにより、8種程度では空色/青、紫/赤紫のような近い候補が混ざりにくくなります。
    if (count <= candidates.length) {
      return selected;
    }

    // 候補数を超える場合のみ、候補をそのまま後ろに足して再利用を許します。
    candidates.forEach((candidate) => {
      if (!selected.some((entry) => entry.name === candidate.name)) {
        selected.push(candidate);
      }
    });

    return selected;
  }

  function averageColorDistance(color, otherColors) {
    if (!otherColors || otherColors.length === 0) return 100;
    let total = 0;
    let count = 0;
    otherColors.forEach((otherColor) => {
      total += colorDistance(color, otherColor);
      count += 1;
    });
    return count > 0 ? total / count : 100;
  }

  function mapVisibilityPenalty(color) {
    const rgb = hexToRgb(color);
    if (!rgb) return 0;
    const luminance = (0.2126 * rgb.r + 0.7152 * rgb.g + 0.0722 * rgb.b) / 255;

    // 淡色地図上で薄すぎる色は、色差があっても見にくいため少し後回しにします。
    if (luminance > 0.82) return 45;
    if (luminance > 0.72) return 18;
    return 0;
  }

  function buildColorCandidates(targetCount) {
    const candidates = [];

    TAXON_COLOR_PALETTE.forEach((color) => {
      addCandidateColor(candidates, color);
    });

    const saturationCycle = [72, 82, 62, 88];
    const lightnessCycle = [42, 55, 34, 66, 48];
    let index = 0;
    let guard = 0;

    while (candidates.length < targetCount && guard < targetCount * 30) {
      const hue = (index * 137.508) % 360;
      const saturation = saturationCycle[index % saturationCycle.length];
      const lightness = lightnessCycle[Math.floor(index / saturationCycle.length) % lightnessCycle.length];
      const color = hslToHex(hue, saturation, lightness);

      const minDistance = minColorDistance(color, candidates);
      if (minDistance >= 24 || candidates.length < TAXON_COLOR_PALETTE.length + 8) {
        addCandidateColor(candidates, color);
      }

      index += 1;
      guard += 1;
    }

    return candidates;
  }

  function addCandidateColor(candidates, color) {
    const normalized = String(color).trim().toUpperCase();
    if (!/^#[0-9A-F]{6}$/.test(normalized)) return;
    if (candidates.includes(normalized)) return;
    candidates.push(normalized);
  }

  function minColorDistance(color, otherColors) {
    if (!otherColors || otherColors.length === 0) return 100;

    let minDistance = Number.POSITIVE_INFINITY;
    otherColors.forEach((otherColor) => {
      const distance = colorDistance(color, otherColor);
      if (distance < minDistance) minDistance = distance;
    });

    return Number.isFinite(minDistance) ? minDistance : 100;
  }

  function hslToHex(hue, saturation, lightness) {
    const s = saturation / 100;
    const l = lightness / 100;
    const c = (1 - Math.abs(2 * l - 1)) * s;
    const h = hue / 60;
    const x = c * (1 - Math.abs((h % 2) - 1));

    let r1 = 0;
    let g1 = 0;
    let b1 = 0;

    if (h >= 0 && h < 1) {
      r1 = c; g1 = x; b1 = 0;
    } else if (h >= 1 && h < 2) {
      r1 = x; g1 = c; b1 = 0;
    } else if (h >= 2 && h < 3) {
      r1 = 0; g1 = c; b1 = x;
    } else if (h >= 3 && h < 4) {
      r1 = 0; g1 = x; b1 = c;
    } else if (h >= 4 && h < 5) {
      r1 = x; g1 = 0; b1 = c;
    } else {
      r1 = c; g1 = 0; b1 = x;
    }

    const m = l - c / 2;
    const toHex = (value) => Math.round((value + m) * 255).toString(16).padStart(2, "0");
    return `#${toHex(r1)}${toHex(g1)}${toHex(b1)}`.toUpperCase();
  }

  function buildTaxonAdjacency(records, taxa) {
    const adjacency = new Map(taxa.map((taxon) => [taxon, new Map()]));
    const nearestCount = Math.max(1, Number(CONFIG.COLOR_NEAREST_DIFFERENT_TAXA) || 5);

    records.forEach((record, index) => {
      const sourceTaxon = taxonKey(record.taxon);
      const nearest = [];

      records.forEach((otherRecord, otherIndex) => {
        if (index === otherIndex) return;

        const targetTaxon = taxonKey(otherRecord.taxon);
        if (sourceTaxon === targetTaxon) return;

        const distance = haversineMeters(record.latitude, record.longitude, otherRecord.latitude, otherRecord.longitude);
        if (!Number.isFinite(distance)) return;

        nearest.push({ taxon: targetTaxon, distance });
      });

      nearest.sort((a, b) => a.distance - b.distance);

      const addedTaxa = new Set();
      for (const neighbor of nearest) {
        if (addedTaxa.has(neighbor.taxon)) continue;

        addedTaxa.add(neighbor.taxon);
        const rank = addedTaxa.size;
        const weight = nearestCount + 1 - rank;
        addAdjacencyWeight(adjacency, sourceTaxon, neighbor.taxon, weight);

        if (addedTaxa.size >= nearestCount) break;
      }
    });

    return adjacency;
  }

  function addAdjacencyWeight(adjacency, taxonA, taxonB, weight) {
    if (taxonA === taxonB) return;

    const mapA = adjacency.get(taxonA);
    const mapB = adjacency.get(taxonB);
    if (!mapA || !mapB) return;

    mapA.set(taxonB, (mapA.get(taxonB) || 0) + weight);
    mapB.set(taxonA, (mapB.get(taxonA) || 0) + weight);
  }

  function weightedDegree(neighborMap) {
    if (!neighborMap) return 0;
    let total = 0;
    neighborMap.forEach((weight) => {
      total += weight;
    });
    return total;
  }

  function haversineMeters(lat1, lon1, lat2, lon2) {
    const toRadians = (degrees) => degrees * Math.PI / 180;
    const earthRadius = 6371008.8;
    const phi1 = toRadians(lat1);
    const phi2 = toRadians(lat2);
    const deltaPhi = toRadians(lat2 - lat1);
    const deltaLambda = toRadians(lon2 - lon1);

    const a = Math.sin(deltaPhi / 2) ** 2
      + Math.cos(phi1) * Math.cos(phi2) * Math.sin(deltaLambda / 2) ** 2;
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    return earthRadius * c;
  }

  function colorDistance(colorA, colorB) {
    const labA = hexToLab(colorA);
    const labB = hexToLab(colorB);
    if (!labA || !labB) return 0;

    const deltaL = labA.l - labB.l;
    const deltaA = labA.a - labB.a;
    const deltaB = labA.b - labB.b;

    return Math.sqrt(deltaL * deltaL + deltaA * deltaA + deltaB * deltaB);
  }

  function hexToLab(hex) {
    const rgb = hexToRgb(hex);
    if (!rgb) return null;

    const srgbToLinear = (value) => {
      const normalized = value / 255;
      if (normalized <= 0.04045) return normalized / 12.92;
      return ((normalized + 0.055) / 1.055) ** 2.4;
    };

    const r = srgbToLinear(rgb.r);
    const g = srgbToLinear(rgb.g);
    const b = srgbToLinear(rgb.b);

    const x = (0.4124564 * r + 0.3575761 * g + 0.1804375 * b) / 0.95047;
    const y = (0.2126729 * r + 0.7151522 * g + 0.0721750 * b) / 1.00000;
    const z = (0.0193339 * r + 0.1191920 * g + 0.9503041 * b) / 1.08883;

    const f = (value) => {
      if (value > 0.008856) return Math.cbrt(value);
      return (7.787 * value) + (16 / 116);
    };

    const fx = f(x);
    const fy = f(y);
    const fz = f(z);

    return {
      l: (116 * fy) - 16,
      a: 500 * (fx - fy),
      b: 200 * (fy - fz)
    };
  }

  function hexToRgb(hex) {
    const match = /^#?([a-fA-F0-9]{6})$/.exec(String(hex).trim());
    if (!match) return null;

    const value = Number.parseInt(match[1], 16);
    return {
      r: (value >> 16) & 255,
      g: (value >> 8) & 255,
      b: value & 255
    };
  }

  function recordKind(recordType) {
    const text = normalizeText(recordType);
    if (text === "タイプ産地") return "type";
    if (text === "シノニマイズされた種のタイプ産地") return "synonymType";
    if (text === "文献記録") return "literature";
    if (text === "本研究") return "thisStudy";
    return "unknown";
  }

  function getMarkerPriority(record) {
    return RECORD_TYPE_PRIORITY[recordKind(record?.recordType)] ?? 0;
  }

  function compareRecordsForRepresentative(a, b) {
    const priorityDiff = getMarkerPriority(b) - getMarkerPriority(a);
    if (priorityDiff !== 0) return priorityDiff;

    const taxonDiff = normalizeText(a.taxon).localeCompare(normalizeText(b.taxon), "ja");
    if (taxonDiff !== 0) return taxonDiff;

    return (a.__index ?? 0) - (b.__index ?? 0);
  }

  function createBaseMapStyle() {
    const initialBaseMap = CONFIG.INITIAL_BASEMAP === "pale" ? "pale" : "detail";

    return {
      version: 8,
      glyphs: "https://demotiles.maplibre.org/font/{fontstack}/{range}.pbf",
      sources: {
        esri: {
          type: "raster",
          tiles: ["https://services.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}"],
          tileSize: 256,
          attribution: DETAILED_MAP_ATTRIBUTION
        },
        gsi_pale: {
          type: "raster",
          tiles: ["https://cyberjapandata.gsi.go.jp/xyz/pale/{z}/{x}/{y}.png"],
          tileSize: 256,
          attribution: GSI_TILE_ATTRIBUTION
        }
      },
      layers: [
        { id: "background", type: "background", paint: { "background-color": "#e8eef2" } },
        {
          id: "detail-layer",
          type: "raster",
          source: "esri",
          layout: { visibility: initialBaseMap === "detail" ? "visible" : "none" },
          minzoom: 0,
          maxzoom: 19
        },
        {
          id: "pale-layer",
          type: "raster",
          source: "gsi_pale",
          layout: { visibility: initialBaseMap === "pale" ? "visible" : "none" },
          minzoom: 0,
          maxzoom: 19
        }
      ]
    };
  }

  function setBaseMap(baseMap) {
    if (!state.map || !state.map.getLayer("detail-layer") || !state.map.getLayer("pale-layer")) return;
    const usePale = baseMap === "pale";
    state.map.setLayoutProperty("detail-layer", "visibility", usePale ? "none" : "visible");
    state.map.setLayoutProperty("pale-layer", "visibility", usePale ? "visible" : "none");
  }

  function initMap() {
    state.map = new maplibregl.Map({
      container: "map",
      style: createBaseMapStyle(),
      center: CONFIG.START_CENTER || [133.0, 33.7],
      zoom: CONFIG.START_ZOOM || 7,
      maxZoom: 18
    });

    state.map.addControl(new maplibregl.NavigationControl({ visualizePitch: true }), "top-right");

    state.map.addControl(new maplibregl.GeolocateControl({
      positionOptions: {
        enableHighAccuracy: true
      },
      fitBoundsOptions: {
        maxZoom: 16
      },
      trackUserLocation: true,
      showUserHeading: true
    }), "top-right");

    if (basemapSelect) {
      basemapSelect.value = CONFIG.INITIAL_BASEMAP === "pale" ? "pale" : "detail";
      basemapSelect.addEventListener("change", () => {
        setBaseMap(basemapSelect.value);
      });
    }

    state.map.on("zoomstart", () => {
      state.isZooming = true;
      setMarkersHidden(true);
      closePopup();
      hideTooltip();
    });

    state.map.on("zoomend", () => {
      window.requestAnimationFrame(() => {
        renderMarkers();
        state.lastZoomRebuildAt = Date.now();
        setMarkersHidden(false);
        state.isZooming = false;
      });
    });

    state.map.on("moveend", () => {
      if (state.isZooming) return;
      if (Date.now() - state.lastZoomRebuildAt < 250) return;
      renderMarkers();
    });
  }

  async function reloadData() {
    const dataset = getActiveDataset();
    if (!dataset) {
      resetMapDataForDatasetChange();
      setStatus("調査名を追加してください。");
      renderDatasetControls();
      return;
    }

    reloadButton.disabled = true;
    setStatus(`「${dataset.name}」を読み込み中...`);

    try {
      const { rows, headers } = await loadSheetRows(dataset);
      validateHeaders(headers);

      const records = normalizeRows(rows);
      state.rows = records;
      assignTaxonColors(records);
      closePopup();
      renderTaxonLegend(records);

      await fitBoundsIfNeeded(records);
      renderMarkers();

      const now = new Date();
      const skipped = rows.length - records.length;
      const skippedText = skipped > 0 ? `、${skipped}件スキップ` : "";
      setStatus(`「${dataset.name}」: ${records.length}件表示${skippedText} / ${now.toLocaleTimeString("ja-JP", { hour: "2-digit", minute: "2-digit" })} 更新`);
      state.hasLoadedOnce = true;
    } catch (error) {
      console.error(error);
      setStatus(`読み込み失敗: ${error.message}`);
      window.alert(`データを読み込めませんでした。

${error.message}`);
    } finally {
      renderDatasetControls();
    }
  }

  function fitBoundsIfNeeded(records) {
    if (!CONFIG.AUTO_FIT_BOUNDS || records.length === 0 || !state.map) {
      return Promise.resolve();
    }

    const bounds = new maplibregl.LngLatBounds();
    records.forEach((record) => {
      bounds.extend([record.longitude, record.latitude]);
    });

    return new Promise((resolve) => {
      const finish = () => resolve();
      state.map.once("moveend", finish);
      state.map.fitBounds(bounds, {
        padding: { top: 90, right: 45, bottom: 90, left: 45 },
        maxZoom: 12,
        duration: state.hasLoadedOnce ? 250 : 600
      });
      window.setTimeout(finish, 900);
    });
  }

  async function loadSheetRows(dataset) {
    if (dataset.csvUrl && dataset.csvUrl.trim()) {
      return loadCsvRows(dataset.csvUrl.trim());
    }
    return loadGvizRows(dataset.spreadsheetId, dataset.sheetName);
  }

  async function loadCsvRows(url) {
    const separator = url.includes("?") ? "&" : "?";
    const response = await fetch(`${url}${separator}_=${Date.now()}`, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`CSVの取得に失敗しました: HTTP ${response.status}`);
    }
    const text = await response.text();
    const table = parseCsv(text);
    if (table.length === 0) {
      throw new Error("CSVに行がありません。");
    }

    const headers = table[0].map(normalizeText);
    const rows = table.slice(1).map((cells) => {
      const row = {};
      headers.forEach((header, index) => {
        row[header] = cells[index] ?? "";
      });
      return row;
    });

    return { headers, rows };
  }

  function loadGvizRows(sheetId, sheetName) {
    return new Promise((resolve, reject) => {
      if (!sheetId) {
        reject(new Error("SHEET_ID が設定されていません。"));
        return;
      }

      const callbackName = `__fieldMapGviz_${Date.now()}_${Math.floor(Math.random() * 100000)}`;
      const script = document.createElement("script");
      const query = new URLSearchParams({
        sheet: sheetName || "",
        headers: "1",
        tqx: `out:json;responseHandler:${callbackName}`,
        _: String(Date.now())
      });

      const timeoutId = window.setTimeout(() => {
        cleanup();
        reject(new Error("Googleスプレッドシートの読み込みがタイムアウトしました。共有設定またはシート名を確認してください。"));
      }, 15000);

      function cleanup() {
        window.clearTimeout(timeoutId);
        delete window[callbackName];
        script.remove();
      }

      window[callbackName] = (response) => {
        try {
          cleanup();
          if (!response || response.status === "error") {
            const message = response?.errors?.map((e) => e.detailed_message || e.message).join(" / ") || "Google Visualization APIの応答がエラーでした。";
            reject(new Error(message));
            return;
          }
          resolve(gvizResponseToRows(response));
        } catch (error) {
          reject(error);
        }
      };

      script.onerror = () => {
        cleanup();
        reject(new Error("Googleスプレッドシートにアクセスできませんでした。リンク共有が有効か確認してください。"));
      };

      script.src = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(sheetId)}/gviz/tq?${query.toString()}`;
      document.head.appendChild(script);
    });
  }

  function gvizResponseToRows(response) {
    const table = response.table;
    if (!table || !Array.isArray(table.cols) || !Array.isArray(table.rows)) {
      throw new Error("スプレッドシートの表形式を読み取れませんでした。");
    }

    const headers = table.cols.map((col) => normalizeText(col.label || col.id));
    const rows = table.rows.map((row) => {
      const obj = {};
      headers.forEach((header, index) => {
        const cell = row.c[index];
        const displayValue = cell ? (cell.f ?? cell.v ?? "") : "";
        const rawValue = cell ? (cell.v ?? "") : "";
        obj[header] = displayValue;
        obj[`__raw_${header}`] = rawValue;
      });
      return obj;
    });

    return { headers, rows };
  }

  function parseCsv(text) {
    const rows = [];
    let row = [];
    let field = "";
    let inQuotes = false;

    for (let i = 0; i < text.length; i += 1) {
      const char = text[i];
      const next = text[i + 1];

      if (inQuotes) {
        if (char === '"' && next === '"') {
          field += '"';
          i += 1;
        } else if (char === '"') {
          inQuotes = false;
        } else {
          field += char;
        }
        continue;
      }

      if (char === '"') {
        inQuotes = true;
      } else if (char === ",") {
        row.push(field);
        field = "";
      } else if (char === "\n") {
        row.push(field);
        rows.push(row);
        row = [];
        field = "";
      } else if (char === "\r") {
        if (next !== "\n") {
          row.push(field);
          rows.push(row);
          row = [];
          field = "";
        }
      } else {
        field += char;
      }
    }

    row.push(field);
    if (row.length > 1 || row[0] !== "") {
      rows.push(row);
    }

    return rows;
  }

  function validateHeaders(headers) {
    const headerSet = new Set(headers.map(normalizeText));
    if (!headerSet.has(COLUMNS.taxon) && headerSet.has("Taxon")) {
      throw new Error("ヘッダーが旧形式です。「Taxon」ではなく「同定」に変更してから、再読み込みしてください。");
    }

    const missing = REQUIRED_HEADERS.filter((header) => !headerSet.has(header));
    if (missing.length > 0) {
      throw new Error(`必須ヘッダーが見つかりません: ${missing.join(", ")}`);
    }
  }

  function normalizeRows(rows) {
    return rows.map((row, index) => {
      const latitude = parseNumber(row[`__raw_${COLUMNS.latitude}`] ?? row[COLUMNS.latitude]);
      const longitude = parseNumber(row[`__raw_${COLUMNS.longitude}`] ?? row[COLUMNS.longitude]);

      return {
        __index: index,
        id: normalizeText(row[COLUMNS.id]),
        taxon: normalizeText(row[COLUMNS.taxon]),
        recordType: normalizeText(row[COLUMNS.recordType]),
        repository: normalizeText(row[COLUMNS.repository]),
        specimenId: normalizeText(row[COLUMNS.specimenId]),
        locality: normalizeText(row[COLUMNS.locality]),
        latitude,
        longitude,
        year: normalizeText(row[COLUMNS.year]),
        month: normalizeText(row[COLUMNS.month]),
        day: normalizeText(row[COLUMNS.day]),
        date: normalizeText(row[COLUMNS.date]),
        remarks: normalizeText(row[COLUMNS.remarks])
      };
    }).filter((record) => {
      if (!Number.isFinite(record.latitude) || !Number.isFinite(record.longitude)) return false;
      if (record.latitude < -90 || record.latitude > 90) return false;
      if (record.longitude < -180 || record.longitude > 180) return false;
      return true;
    });
  }

  function renderMarkers() {
    if (!state.map) {
      return;
    }
    if (!state.rows.length) {
      clearMarkers();
      return;
    }

    clearMarkers();
    const groups = buildMarkerGroups(state.rows);
    state.markerGroups = groups;

    groups.forEach((group) => {
      const element = createMarkerElement(group);
      const marker = new maplibregl.Marker({ element })
        .setLngLat([group.representative.longitude, group.representative.latitude])
        .addTo(state.map);

      element.addEventListener("click", (event) => {
        event.stopPropagation();
        state.currentPopupRecords = group.records;
        state.currentPopupIndex = 0;
        showPopup(0, false);
      });

      state.markers.push(marker);
    });
  }

  function buildMarkerGroups(records) {
    const proximityPixels = Number(CONFIG.PROXIMITY_PIXELS) || 10;
    const sorted = [...records].sort(compareRecordsForRepresentative);
    const groups = [];

    sorted.forEach((record) => {
      const point = state.map.project([record.longitude, record.latitude]);
      let matchedGroup = null;

      for (const group of groups) {
        const dx = point.x - group.pixel.x;
        const dy = point.y - group.pixel.y;
        const distance = Math.sqrt(dx * dx + dy * dy);
        if (distance <= proximityPixels) {
          matchedGroup = group;
          break;
        }
      }

      if (matchedGroup) {
        matchedGroup.records.push(record);
        matchedGroup.records.sort(compareRecordsForRepresentative);
        matchedGroup.representative = matchedGroup.records[0];
        matchedGroup.pixel = state.map.project([matchedGroup.representative.longitude, matchedGroup.representative.latitude]);
      } else {
        groups.push({
          representative: record,
          records: [record],
          pixel: point
        });
      }
    });

    return groups.sort((a, b) => compareRecordsForRepresentative(a.representative, b.representative));
  }

  function createMarkerElement(group) {
    const record = group.representative;
    const kind = recordKind(record.recordType);
    const segments = getMarkerSegments(group.records);
    const wrapper = document.createElement("button");
    wrapper.type = "button";
    wrapper.className = `record-marker kind-${kind}${segments.length > 1 ? " multi-taxon" : ""}`;
    wrapper.setAttribute("aria-label", `${record.taxon || "記録"} ${record.recordType || ""}`.trim());
    wrapper.innerHTML = markerSvg(kind, segments);

    return wrapper;
  }

  function getMarkerSegments(records) {
    const byTaxon = new Map();

    records.forEach((record) => {
      const key = taxonKey(record.taxon);
      if (!byTaxon.has(key)) {
        byTaxon.set(key, []);
      }
      byTaxon.get(key).push(record);
    });

    return [...byTaxon.entries()]
      .map(([taxon, taxonRecords]) => {
        const sortedRecords = [...taxonRecords].sort(compareRecordsForRepresentative);
        const representative = sortedRecords[0];
        return {
          taxon,
          color: colorForTaxon(taxon),
          kind: recordKind(representative.recordType),
          representative
        };
      })
      .sort((a, b) => compareRecordsForRepresentative(a.representative, b.representative));
  }

  function markerSvg(kind, segmentsOrColor) {
    const segments = normalizeMarkerSegments(segmentsOrColor);
    if (segments.length <= 1) {
      const segment = segments[0] || { color: "#666666", kind };
      return singleColorMarkerSvg(segment.kind || kind, segment.color || "#666666");
    }
    return splitColorMarkerSvg(kind, segments);
  }

  function normalizeMarkerSegments(value) {
    if (!Array.isArray(value)) {
      return [{ color: value || "#666666", kind: "other" }];
    }

    return value.map((item) => {
      if (typeof item === "string") {
        return { color: item || "#666666", kind: "other" };
      }
      return {
        color: item?.color || "#666666",
        kind: item?.kind || "other",
        taxon: item?.taxon || "",
        representative: item?.representative || null
      };
    }).filter((item) => item.color);
  }

  function singleColorMarkerSvg(kind, color) {
    const stroke = "rgba(0,0,0,0.72)";
    const white = "rgba(255,255,255,0.94)";

    if (kind === "type") {
      return `
        <svg viewBox="0 0 24 24" role="img" aria-hidden="true">
          <polygon points="12,1.6 14.8,8.4 22.1,8.9 16.5,13.6 18.2,20.8 12,17 5.8,20.8 7.5,13.6 1.9,8.9 9.2,8.4"
            fill="${color}" stroke="${stroke}" stroke-width="1.4" />
        </svg>`;
    }

    if (kind === "synonymType") {
      return `
        <svg viewBox="0 0 24 24" role="img" aria-hidden="true">
          <rect x="5" y="5" width="14" height="14" rx="1.8"
            fill="${color}" stroke="${stroke}" stroke-width="1.7" />
        </svg>`;
    }

    if (kind === "literature") {
      return `
        <svg viewBox="0 0 24 24" role="img" aria-hidden="true">
          <circle cx="12" cy="12" r="7.3"
            fill="${white}" stroke="${color}" stroke-width="3.2" />
          <circle cx="12" cy="12" r="9.1"
            fill="none" stroke="${stroke}" stroke-width="0.9" />
        </svg>`;
    }

    return `
      <svg viewBox="0 0 24 24" role="img" aria-hidden="true">
        <circle cx="12" cy="12" r="7.8"
          fill="${color}" stroke="${stroke}" stroke-width="1.7" />
      </svg>`;
  }

  function splitColorMarkerSvg(kind, segments) {
    const clipId = `markerClip_${Math.random().toString(36).slice(2)}`;
    const innerClipId = `markerInnerClip_${Math.random().toString(36).slice(2)}`;
    const stroke = "rgba(0,0,0,0.82)";
    const innerWhite = "rgba(255,255,255,0.96)";
    const segmentList = normalizeMarkerSegments(segments);
    const slices = radialColorSlices(segmentList, clipId);
    const hollowSlices = radialHollowSlices(segmentList, innerClipId, innerWhite);

    if (kind === "type") {
      const starPoints = "12,1.6 14.8,8.4 22.1,8.9 16.5,13.6 18.2,20.8 12,17 5.8,20.8 7.5,13.6 1.9,8.9 9.2,8.4";
      return `
        <svg viewBox="0 0 24 24" role="img" aria-hidden="true">
          <defs>
            <clipPath id="${clipId}"><polygon points="${starPoints}" /></clipPath>
            <clipPath id="${innerClipId}"><circle cx="12" cy="12" r="5.6" /></clipPath>
          </defs>
          ${slices}
          ${hollowSlices}
          <polygon points="${starPoints}" fill="none" stroke="${stroke}" stroke-width="1.8" />
        </svg>`;
    }

    if (kind === "synonymType") {
      return `
        <svg viewBox="0 0 24 24" role="img" aria-hidden="true">
          <defs>
            <clipPath id="${clipId}"><rect x="5" y="5" width="14" height="14" rx="1.8" /></clipPath>
            <clipPath id="${innerClipId}"><rect x="8" y="8" width="8" height="8" rx="1.1" /></clipPath>
          </defs>
          ${slices}
          ${hollowSlices}
          <rect x="5" y="5" width="14" height="14" rx="1.8" fill="none" stroke="${stroke}" stroke-width="1.9" />
        </svg>`;
    }

    if (kind === "literature") {
      return `
        <svg viewBox="0 0 24 24" role="img" aria-hidden="true">
          <defs>
            <clipPath id="${clipId}"><circle cx="12" cy="12" r="9.2" /></clipPath>
            <clipPath id="${innerClipId}"><circle cx="12" cy="12" r="5.5" /></clipPath>
          </defs>
          ${slices}
          ${radialHollowSlices(segmentList.map((segment) => ({ ...segment, kind: "literature" })), innerClipId, innerWhite)}
          <circle cx="12" cy="12" r="5.5" fill="none" stroke="${stroke}" stroke-width="0.8" />
          <circle cx="12" cy="12" r="9.2" fill="none" stroke="${stroke}" stroke-width="1.1" />
        </svg>`;
    }

    return `
      <svg viewBox="0 0 24 24" role="img" aria-hidden="true">
        <defs>
          <clipPath id="${clipId}"><circle cx="12" cy="12" r="8.2" /></clipPath>
          <clipPath id="${innerClipId}"><circle cx="12" cy="12" r="4.9" /></clipPath>
        </defs>
        ${slices}
        ${hollowSlices}
        <circle cx="12" cy="12" r="4.9" fill="none" stroke="rgba(0,0,0,0.22)" stroke-width="0.45" />
        <circle cx="12" cy="12" r="8.2" fill="none" stroke="${stroke}" stroke-width="1.9" />
      </svg>`;
  }

  function isHollowMarkerSegment(kind) {
    return kind === "literature";
  }

  function radialColorSlices(segments, clipId) {
    const paths = radialSlicePaths(segments.length);
    const slices = segments.map((segment, index) => {
      return `<path d="${paths[index]}" fill="${segment.color}" clip-path="url(#${clipId})" />`;
    });

    const separators = radialSeparators(segments.length, clipId);
    return slices.concat(separators).join("");
  }

  function radialHollowSlices(segments, innerClipId, fillColor) {
    const paths = radialSlicePaths(segments.length);
    return segments.map((segment, index) => {
      if (!isHollowMarkerSegment(segment.kind)) {
        return "";
      }
      return `<path d="${paths[index]}" fill="${fillColor}" clip-path="url(#${innerClipId})" />`;
    }).join("");
  }

  function radialSlicePaths(segmentCount) {
    const centerX = 12;
    const centerY = 12;
    const radius = 19;
    const segmentAngle = (Math.PI * 2) / segmentCount;
    const startOffset = -Math.PI / 2;

    return Array.from({ length: segmentCount }, (_, index) => {
      const startAngle = startOffset + index * segmentAngle;
      const endAngle = startOffset + (index + 1) * segmentAngle;
      const x1 = centerX + radius * Math.cos(startAngle);
      const y1 = centerY + radius * Math.sin(startAngle);
      const x2 = centerX + radius * Math.cos(endAngle);
      const y2 = centerY + radius * Math.sin(endAngle);
      const largeArc = segmentAngle > Math.PI ? 1 : 0;
      return [
        `M ${centerX} ${centerY}`,
        `L ${x1.toFixed(3)} ${y1.toFixed(3)}`,
        `A ${radius} ${radius} 0 ${largeArc} 1 ${x2.toFixed(3)} ${y2.toFixed(3)}`,
        "Z"
      ].join(" ");
    });
  }

  function radialSeparators(segmentCount, clipId) {
    if (segmentCount <= 1) {
      return [];
    }

    const centerX = 12;
    const centerY = 12;
    const radius = 19;
    const segmentAngle = (Math.PI * 2) / segmentCount;
    const startOffset = -Math.PI / 2;

    return Array.from({ length: segmentCount }, (_, index) => {
      const angle = startOffset + index * segmentAngle;
      const x = centerX + radius * Math.cos(angle);
      const y = centerY + radius * Math.sin(angle);
      return `<line x1="${centerX}" y1="${centerY}" x2="${x.toFixed(3)}" y2="${y.toFixed(3)}" stroke="rgba(255,255,255,0.86)" stroke-width="0.65" clip-path="url(#${clipId})" />`;
    });
  }

  function createPopupContent(record) {
    const content = document.createElement("div");
    content.className = "popup-content";

    const title = document.createElement("div");
    title.className = "popup-title";

    const titleChip = document.createElement("span");
    titleChip.className = "popup-title-chip";
    titleChip.style.background = colorForTaxon(record.taxon);

    const titleText = document.createElement("span");
    titleText.innerHTML = formatTaxonHTML(record.taxon);

    title.appendChild(titleChip);
    title.appendChild(titleText);
    content.appendChild(title);

    appendPopupRow(content, "記録の種類", record.recordType);
    appendPopupRow(content, "場所", record.locality);
    appendCoordinateRow(content, record);
    appendPopupRow(content, "日付", record.date || joinDateParts(record));
    appendPopupRow(content, "所蔵", record.repository);
    appendPopupRow(content, "標本ID", record.specimenId);
    appendPopupRow(content, "ID", record.id);
    appendPopupRow(content, "備考", record.remarks);

    return content;
  }

  function appendPopupRow(container, label, value) {
    const row = document.createElement("div");
    row.className = "popup-row";

    const labelEl = document.createElement("div");
    labelEl.className = "popup-label";
    labelEl.textContent = label;

    const valueEl = document.createElement("div");
    valueEl.className = "popup-value";
    valueEl.textContent = displayText(value);

    row.appendChild(labelEl);
    row.appendChild(valueEl);
    container.appendChild(row);
  }

  function appendCoordinateRow(container, record) {
    const row = document.createElement("div");
    row.className = "popup-row";

    const labelEl = document.createElement("div");
    labelEl.className = "popup-label";
    labelEl.textContent = "Latitude, Longitude";

    const valueEl = document.createElement("div");
    valueEl.className = "popup-value";

    const lat = formatCoordinate(record.latitude);
    const lng = formatCoordinate(record.longitude);
    valueEl.appendChild(document.createTextNode(`${lat}, ${lng} `));

    const link = document.createElement("a");
    link.href = `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(`${record.latitude},${record.longitude}`)}`;
    link.target = "_blank";
    link.rel = "noopener noreferrer";
    link.className = "popup-link-pill";
    link.textContent = "Google Map";
    valueEl.appendChild(link);

    row.appendChild(labelEl);
    row.appendChild(valueEl);
    container.appendChild(row);
  }

  function createPopupNav(index, total) {
    const nav = document.createElement("div");
    nav.className = "popup-nav-fixed";

    if (total <= 1) {
      const counterOnly = document.createElement("span");
      counterOnly.className = "popup-counter";
      counterOnly.textContent = "1 / 1";
      nav.appendChild(counterOnly);
      return { nav, prevButton: null, nextButton: null };
    }

    const prevButton = document.createElement("button");
    prevButton.type = "button";
    prevButton.textContent = "前へ";

    const counter = document.createElement("span");
    counter.className = "popup-counter";
    counter.textContent = `${index + 1} / ${total}`;

    const nextButton = document.createElement("button");
    nextButton.type = "button";
    nextButton.textContent = "次へ";

    nav.appendChild(prevButton);
    nav.appendChild(counter);
    nav.appendChild(nextButton);

    return { nav, prevButton, nextButton };
  }

  function showPopup(index, preserveAnchor) {
    const records = state.currentPopupRecords;
    if (!records.length || !state.map) return;

    const normalizedIndex = ((index % records.length) + records.length) % records.length;
    const record = records[normalizedIndex];
    state.currentPopupIndex = normalizedIndex;

    if (state.activePopup) {
      state.activePopup.remove();
      state.activePopup = null;
    }

    const markerPixel = state.map.project([record.longitude, record.latitude]);
    const mapHeight = state.map.getContainer().offsetHeight;
    const margin = 84;
    const distanceFromTop = markerPixel.y;
    const distanceFromBottom = mapHeight - markerPixel.y;

    let anchor = state.currentAnchor;
    let showAbove = state.currentShowAbove;

    if (!(preserveAnchor && anchor && showAbove !== null)) {
      showAbove = distanceFromTop >= distanceFromBottom;
      anchor = showAbove ? "bottom" : "top";
    }

    const maxHeight = showAbove
      ? Math.max(110, distanceFromTop - margin)
      : Math.max(110, distanceFromBottom - margin);

    const wrapper = document.createElement("div");
    wrapper.className = "popup-wrapper";

    const scroll = document.createElement("div");
    scroll.className = "popup-scroll-container";
    scroll.style.maxHeight = `${maxHeight}px`;
    scroll.appendChild(createPopupContent(record));

    const { nav, prevButton, nextButton } = createPopupNav(normalizedIndex, records.length);

    if (!showAbove) {
      wrapper.appendChild(nav);
    }
    wrapper.appendChild(scroll);
    if (showAbove) {
      wrapper.appendChild(nav);
    }

    const popup = new maplibregl.Popup({
      focusAfterOpen: false,
      closeOnClick: false,
      anchor,
      offset: 0,
      maxWidth: "min(360px, calc(100vw - 28px))"
    })
      .setLngLat([record.longitude, record.latitude])
      .setDOMContent(wrapper)
      .addTo(state.map);

    if (prevButton) {
      prevButton.addEventListener("click", (event) => {
        event.stopPropagation();
        showPopup(state.currentPopupIndex - 1, true);
      });
    }

    if (nextButton) {
      nextButton.addEventListener("click", (event) => {
        event.stopPropagation();
        showPopup(state.currentPopupIndex + 1, true);
      });
    }

    state.activePopup = popup;
    state.currentAnchor = anchor;
    state.currentShowAbove = showAbove;
  }

  function closePopup() {
    if (state.activePopup) {
      state.activePopup.remove();
      state.activePopup = null;
    }
    state.currentPopupRecords = [];
    state.currentPopupIndex = 0;
    state.currentAnchor = null;
    state.currentShowAbove = null;
  }

  function clearMarkers() {
    state.markers.forEach((marker) => marker.remove());
    state.markers = [];
    state.markerGroups = [];
  }

  function setMarkersHidden(hidden) {
    if (!state.map) return;
    const container = state.map.getContainer();
    container.classList.toggle("markers-hidden", Boolean(hidden));
  }

  function hideTooltip() {
    const tooltip = document.querySelector(".marker-tooltip");
    if (tooltip) {
      tooltip.style.display = "none";
    }
  }

  function joinDateParts(record) {
    const parts = [record.year, record.month, record.day].map(normalizeText).filter(Boolean);
    return parts.join("-");
  }

  function renderTaxonLegend(records) {
    const taxa = [...new Set(records.map((record) => taxonKey(record.taxon)))]
      .sort((a, b) => a.localeCompare(b, "ja"));

    taxonLegendEl.innerHTML = taxa.map((taxon) => `
      <div class="taxon-row">
        <span class="taxon-chip" style="background: ${colorForTaxon(taxon)}"></span>
        <span>${formatTaxonHTML(taxon)}</span>
      </div>`).join("");
  }

  function setupPopupClose() {
    document.addEventListener("click", (event) => {
      if (!state.activePopup) return;
      const target = event.target instanceof Element ? event.target : event.target?.parentElement;
      if (!target) return;
      const insidePopup = target.closest(".maplibregl-popup");
      const insideMarker = target.closest(".record-marker");
      const insideMapControl = target.closest(".maplibregl-ctrl");
      if (!insidePopup && !insideMarker && !insideMapControl) {
        closePopup();
      }
    }, true);
  }


  if (topbarToggleButton) {
    topbarToggleButton.addEventListener("click", () => setTopbarCollapsed(true));
  }

  if (topbarOpenButton) {
    topbarOpenButton.addEventListener("click", () => setTopbarCollapsed(false));
  }

  if (datasetForm) {
    datasetForm.addEventListener("submit", (event) => {
      event.preventDefault();
      submitDatasetModal();
    });
  }

  if (datasetModalCloseButton) datasetModalCloseButton.addEventListener("click", () => closeDatasetModal(null));
  if (datasetModalCancelButton) datasetModalCancelButton.addEventListener("click", () => closeDatasetModal(null));
  if (datasetModal) {
    datasetModal.addEventListener("click", (event) => {
      if (event.target === datasetModal) closeDatasetModal(null);
    });
  }

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && datasetModal && !datasetModal.hidden) {
      closeDatasetModal(null);
    }
  });

  if (datasetSelect) {
    datasetSelect.addEventListener("change", () => {
      state.activeDatasetId = datasetSelect.value;
      saveActiveDatasetId(state.activeDatasetId);
      resetMapDataForDatasetChange();
      reloadData();
    });
  }

  if (datasetAddButton) datasetAddButton.addEventListener("click", addDataset);
  if (datasetEditButton) datasetEditButton.addEventListener("click", editDataset);
  if (datasetDeleteButton) datasetDeleteButton.addEventListener("click", deleteDataset);

  reloadButton.addEventListener("click", () => {
    reloadData();
  });


  setupPopupClose();
  initializeDatasetState();
  setTopbarCollapsed(loadTopbarCollapsed(), false);
  initMap();
  state.map.on("load", () => {
    reloadData();
  });
})();
