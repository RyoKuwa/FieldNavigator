(() => {
  "use strict";

  const CONFIG = window.FIELD_MAP_CONFIG;
  const COLUMNS = CONFIG.COLUMNS;
  const REQUIRED_HEADERS = [];

  const LEGACY_TAXON_HEADER = "同定";
  const LEGACY_LATITUDE_HEADER = "Latitude";
  const LEGACY_LONGITUDE_HEADER = "Longitude";
  const LEGACY_DATE_HEADER = "表示する日付";
  const REJECTED_TAXON_HEADER = "Taxon";
  const VIRTUAL_COORDINATES_FIELD = "__coordinates";
  const DEFAULT_MARKER_COLOR = "#666666";
  const POPUP_NEARBY_MARKER_PIXELS = Number(CONFIG.POPUP_NEARBY_MARKER_PIXELS) || Number(CONFIG.PROXIMITY_PIXELS) || 10;
  const LOAD_TIMEOUT_MS = Number(CONFIG.LOAD_TIMEOUT_MS) || 30000;
  const LOAD_SLOW_NOTICE_MS = Number(CONFIG.LOAD_SLOW_NOTICE_MS) || 10000;

  const FILTER_FIELD_LABEL_OVERRIDES = {
    [COLUMNS.taxon]: "分類群",
    [LEGACY_TAXON_HEADER]: "分類群",
    [COLUMNS.latitude]: "緯度",
    [LEGACY_LATITUDE_HEADER]: "緯度",
    [COLUMNS.longitude]: "経度",
    [LEGACY_LONGITUDE_HEADER]: "経度",
    [COLUMNS.date]: "日付",
    [LEGACY_DATE_HEADER]: "日付"
  };

  const DEFAULT_POPUP_FIELDS = [
    COLUMNS.recordType,
    COLUMNS.locality,
    VIRTUAL_COORDINATES_FIELD,
    COLUMNS.date,
    COLUMNS.repository,
    COLUMNS.specimenId,
    COLUMNS.id,
    COLUMNS.remarks
  ];

  const DEFAULT_FILTER_FIELDS = [
    COLUMNS.taxon,
    COLUMNS.recordType,
    COLUMNS.repository,
    COLUMNS.year,
    COLUMNS.month,
    COLUMNS.day,
    COLUMNS.date
  ];

  const EMPTY_FILTER_VALUE = "__FIELD_MAP_EMPTY__";

  const RECORD_TYPE_PRIORITY = {
    // 代表記録の選択用。形状を最優先にするため、星 > 四角 > 丸 > ✕ > △ の順にする。
    type: 300,
    synonymType: 200,
    thisStudy: 100,
    literature: 90,
    uncertain: 80,
    unknown: 70
  };

  const RECORD_TYPE_LEGEND_ITEMS = [
    { kind: "type", label: "タイプ産地" },
    { kind: "synonymType", label: "シノニマイズされた種のタイプ産地" },
    { kind: "thisStudy", label: "作成者の記録" },
    { kind: "literature", label: "文献記録" },
    { kind: "uncertain", label: "不確かな記録" },
    { kind: "unknown", label: "その他" }
  ];

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
    filteredRows: [],
    filterOptions: new Map(),
    filters: new Map(),
    filterFields: [],
    headers: [],
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
    temporaryDataset: null,
    activeDatasetId: null,
    sharedDatasetIdFromUrl: null,
    lastLoadedAt: null,
    skippedRows: 0,
    isLoading: false,
    currentLoadJob: null,
    focusButton: null
  };

  const appEl = document.getElementById("app");
  const statusEl = document.getElementById("status");
  const titleEl = document.getElementById("title");
  const topbarToggleButton = document.getElementById("topbar-toggle-button");
  const topbarOpenButton = document.getElementById("topbar-open-button");
  const reloadButton = document.getElementById("reload-button");
  const reloadCancelButton = document.getElementById("reload-cancel-button");
  const taxonLegendEl = document.getElementById("taxon-legend");
  const recordTypeLegendEl = document.getElementById("record-type-legend");
  const basemapSelect = document.getElementById("basemap-select");
  const filterButton = document.getElementById("filter-button");
  const filterSummary = document.getElementById("filter-summary");
  const filterModal = document.getElementById("filter-modal");
  const filterModalCloseButton = document.getElementById("filter-modal-close");
  const filterModalCancelButton = document.getElementById("filter-modal-cancel");
  const filterApplyButton = document.getElementById("filter-apply-button");
  const filterSelectAllButton = document.getElementById("filter-select-all-button");
  const filterSelectNoneButton = document.getElementById("filter-select-none-button");
  const filterResetButton = document.getElementById("filter-reset-button");
  const filterOptionsEl = document.getElementById("filter-options");
  const datasetSelect = document.getElementById("dataset-select");
  const datasetAddButton = document.getElementById("dataset-add-button");
  const datasetEditButton = document.getElementById("dataset-edit-button");
  const datasetDeleteButton = document.getElementById("dataset-delete-button");
  const datasetShareButton = document.getElementById("dataset-share-button");
  const datasetRegisterButton = document.getElementById("dataset-register-button");
  const helpButton = document.getElementById("help-button");
  const helpModal = document.getElementById("help-modal");
  const helpModalCloseButton = document.getElementById("help-modal-close");
  const helpModalCancelButton = document.getElementById("help-modal-cancel");
  const shareModal = document.getElementById("share-modal");
  const shareModalCloseButton = document.getElementById("share-modal-close");
  const shareModalCancelButton = document.getElementById("share-modal-cancel");
  const shareCopyButton = document.getElementById("share-copy-button");
  const shareUrlOutput = document.getElementById("share-url-output");
  const shareSourceLink = document.getElementById("share-source-link");
  const shareModalMessage = document.getElementById("share-modal-message");
  const registerModal = document.getElementById("register-modal");
  const registerModalCloseButton = document.getElementById("register-modal-close");
  const registerModalCancelButton = document.getElementById("register-modal-cancel");
  const registerConfirmButton = document.getElementById("register-confirm-button");
  const registerModalError = document.getElementById("register-modal-error");
  const datasetModal = document.getElementById("dataset-modal");
  const datasetForm = document.getElementById("dataset-form");
  const datasetModalTitle = document.getElementById("dataset-modal-title");
  const datasetModalCloseButton = document.getElementById("dataset-modal-close");
  const datasetModalCancelButton = document.getElementById("dataset-modal-cancel");
  const datasetModalSubmitButton = document.getElementById("dataset-modal-submit");
  const datasetNameInput = document.getElementById("dataset-name-input");
  const datasetUrlInput = document.getElementById("dataset-url-input");
  const datasetSheetInput = document.getElementById("dataset-sheet-input");
  const datasetLoadFieldsButton = document.getElementById("dataset-load-fields-button");
  const datasetFieldSettings = document.getElementById("dataset-field-settings");
  const datasetFieldsStatus = document.getElementById("dataset-fields-status");
  const datasetLatitudeFieldSelect = document.getElementById("dataset-latitude-field-select");
  const datasetLongitudeFieldSelect = document.getElementById("dataset-longitude-field-select");
  const datasetColorFieldSelect = document.getElementById("dataset-color-field-select");
  const datasetPopupTitleFieldSelect = document.getElementById("dataset-popup-title-field-select");
  const datasetPopupFieldsList = document.getElementById("dataset-popup-fields-list");
  const datasetFilterFieldsList = document.getElementById("dataset-filter-fields-list");
  const colorLegendTitleEl = document.getElementById("color-legend-title");
  const datasetModalError = document.getElementById("dataset-modal-error");
  const DATASETS_STORAGE_KEY = CONFIG.LOCAL_DATASETS_STORAGE_KEY || "fieldMap.localDatasets.v1";
  const ACTIVE_DATASET_STORAGE_KEY = CONFIG.ACTIVE_DATASET_STORAGE_KEY || "fieldMap.activeDatasetId.v1";
  const TOPBAR_COLLAPSED_STORAGE_KEY = CONFIG.TOPBAR_COLLAPSED_STORAGE_KEY || "fieldMap.topbarCollapsed.v1";
  const FILTERS_STORAGE_KEY_PREFIX = CONFIG.FILTERS_STORAGE_KEY_PREFIX || "fieldMap.filters.v1";
  titleEl.textContent = CONFIG.TITLE || "調査結果地図";

  function setStatus(message) {
    statusEl.textContent = message;
  }

  function makeAbortError() {
    const error = new Error("読み込みを中止しました。");
    error.name = "AbortError";
    return error;
  }

  function isAbortError(error) {
    return error?.name === "AbortError" || error?.message === "読み込みを中止しました。";
  }

  function createLoadJob() {
    return {
      aborted: false,
      abortController: typeof AbortController !== "undefined" ? new AbortController() : null,
      cleanupHandlers: new Set(),
      timers: new Set()
    };
  }

  function addLoadTimer(job, callback, delay) {
    if (!job) return null;
    const timerId = window.setTimeout(() => {
      job.timers.delete(timerId);
      if (!job.aborted && state.currentLoadJob === job) callback();
    }, delay);
    job.timers.add(timerId);
    return timerId;
  }

  function clearLoadJob(job) {
    if (!job) return;
    job.timers.forEach((timerId) => window.clearTimeout(timerId));
    job.timers.clear();
    job.cleanupHandlers.forEach((handler) => {
      try { handler(); } catch (error) { console.warn(error); }
    });
    job.cleanupHandlers.clear();
  }

  function throwIfLoadCancelled(job) {
    if (job?.aborted || (job && state.currentLoadJob !== job)) {
      throw makeAbortError();
    }
  }

  function setLoadingUi(loading) {
    state.isLoading = Boolean(loading);
    if (reloadButton) reloadButton.disabled = loading || !getActiveDataset();
    if (reloadCancelButton) reloadCancelButton.hidden = !loading;
    if (datasetSelect) datasetSelect.disabled = loading || datasetSelect.options.length === 0;
    if (datasetAddButton) datasetAddButton.disabled = loading;
    if (datasetEditButton) datasetEditButton.disabled = loading || !isDatasetRegistered();
    if (datasetDeleteButton) datasetDeleteButton.disabled = loading || !isDatasetRegistered();
    if (datasetShareButton) datasetShareButton.disabled = loading || !getActiveDataset();
    if (datasetRegisterButton) datasetRegisterButton.disabled = loading || !getActiveDataset() || isDatasetRegistered();
    if (filterButton) filterButton.disabled = loading || state.rows.length === 0;
    updateFocusControlState();
  }

  function cancelActiveLoad(showMessage = true) {
    const job = state.currentLoadJob;
    if (!job) return;
    job.aborted = true;
    if (job.abortController) job.abortController.abort();
    clearLoadJob(job);
    state.currentLoadJob = null;
    setLoadingUi(false);
    if (showMessage) setStatus("読み込みを中止しました。");
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
          updatedAt: normalizeText(dataset.updatedAt),
          latitudeField: normalizeDatasetFieldName(dataset.latitudeField),
          longitudeField: normalizeDatasetFieldName(dataset.longitudeField),
          colorField: normalizeDatasetFieldName(dataset.colorField),
          popupTitleField: normalizeDatasetFieldName(dataset.popupTitleField),
          popupFields: normalizeOptionalFieldList(dataset.popupFields),
          filterFields: normalizeOptionalFieldList(dataset.filterFields)
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

  function datasetIdentityKey(dataset) {
    if (!dataset) return "";
    const sheetName = normalizeText(dataset.sheetName) || CONFIG.DEFAULT_SHEET_NAME || "シート1";
    if (dataset.spreadsheetId) {
      return `sheet:${normalizeText(dataset.spreadsheetId)}:${sheetName}`;
    }
    if (dataset.csvUrl) {
      return `csv:${normalizeText(dataset.csvUrl)}`;
    }
    return "";
  }

  function findDatasetByIdentity(dataset) {
    const key = datasetIdentityKey(dataset);
    if (!key) return null;
    return state.datasets.find((existing) => datasetIdentityKey(existing) === key) || null;
  }

  function datasetFromShareParams() {
    const params = new URLSearchParams(window.location.search);
    const name = normalizeText(params.get("fm_name"));
    const spreadsheetId = normalizeText(params.get("fm_sheet_id"));
    const sheetName = normalizeText(params.get("fm_sheet")) || CONFIG.DEFAULT_SHEET_NAME || "シート1";
    const csvUrl = normalizeText(params.get("fm_csv"));
    if (!name || (!spreadsheetId && !csvUrl)) return null;
    const sourceUrl = spreadsheetId
      ? `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`
      : csvUrl;
    return {
      id: "shared-" + hashString([name, spreadsheetId, sheetName, csvUrl].join("|")).toString(36),
      name,
      sourceUrl,
      spreadsheetId,
      sheetName,
      csvUrl,
      latitudeField: normalizeDatasetFieldName(params.get("fm_lat")),
      longitudeField: normalizeDatasetFieldName(params.get("fm_lng")),
      colorField: normalizeDatasetFieldName(params.get("fm_color")),
      popupTitleField: normalizeDatasetFieldName(params.get("fm_popup_title")),
      popupFields: parseFieldListParam(params.get("fm_popup")),
      filterFields: parseFieldListParam(params.get("fm_filter")),
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };
  }

  function getActiveDataset() {
    return state.datasets.find((dataset) => dataset.id === state.activeDatasetId)
      || (state.temporaryDataset && state.temporaryDataset.id === state.activeDatasetId ? state.temporaryDataset : null)
      || null;
  }

  function isDatasetRegistered(dataset = getActiveDataset()) {
    if (!dataset) return false;
    return state.datasets.some((stored) => stored.id === dataset.id);
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

  function normalizeDatasetFieldName(fieldName) {
    const field = normalizeText(fieldName);
    if (!field) return "";
    if (field === LEGACY_TAXON_HEADER) return COLUMNS.taxon;
    if (field === LEGACY_LATITUDE_HEADER) return COLUMNS.latitude;
    if (field === LEGACY_LONGITUDE_HEADER) return COLUMNS.longitude;
    if (field === LEGACY_DATE_HEADER) return COLUMNS.date;
    return field;
  }

  function uniqueFieldList(fields) {
    const result = [];
    const seen = new Set();
    (Array.isArray(fields) ? fields : []).forEach((field) => {
      const normalized = normalizeDatasetFieldName(field);
      if (!normalized || seen.has(normalized)) return;
      seen.add(normalized);
      result.push(normalized);
    });
    return result;
  }

  function normalizeOptionalFieldList(fields) {
    return Array.isArray(fields) ? uniqueFieldList(fields) : null;
  }

  function parseFieldListParam(value) {
    const text = normalizeText(value);
    if (!text) return null;
    try {
      const parsed = JSON.parse(text);
      return normalizeOptionalFieldList(parsed);
    } catch (error) {
      return normalizeOptionalFieldList(text.split("|"));
    }
  }

  function normalizeFieldsForHeaders(fields, headers, { includeVirtual = false } = {}) {
    const available = new Set((headers || []).map(normalizeText).filter(Boolean));
    const result = [];
    const seen = new Set();
    uniqueFieldList(fields).forEach((field) => {
      let normalized = field;
      if (field === COLUMNS.taxon && !available.has(field) && available.has(LEGACY_TAXON_HEADER)) {
        normalized = LEGACY_TAXON_HEADER;
      } else if (field === LEGACY_TAXON_HEADER && available.has(COLUMNS.taxon)) {
        normalized = COLUMNS.taxon;
      } else if (field === COLUMNS.latitude && !available.has(field) && available.has(LEGACY_LATITUDE_HEADER)) {
        normalized = LEGACY_LATITUDE_HEADER;
      } else if (field === LEGACY_LATITUDE_HEADER && available.has(COLUMNS.latitude)) {
        normalized = COLUMNS.latitude;
      } else if (field === COLUMNS.longitude && !available.has(field) && available.has(LEGACY_LONGITUDE_HEADER)) {
        normalized = LEGACY_LONGITUDE_HEADER;
      } else if (field === LEGACY_LONGITUDE_HEADER && available.has(COLUMNS.longitude)) {
        normalized = COLUMNS.longitude;
      } else if (field === COLUMNS.date && !available.has(field) && available.has(LEGACY_DATE_HEADER)) {
        normalized = LEGACY_DATE_HEADER;
      } else if (field === LEGACY_DATE_HEADER && available.has(COLUMNS.date)) {
        normalized = COLUMNS.date;
      }
      if (normalized === VIRTUAL_COORDINATES_FIELD && includeVirtual) {
        if (!seen.has(normalized)) {
          seen.add(normalized);
          result.push(normalized);
        }
        return;
      }
      if (!available.has(normalized) || seen.has(normalized)) return;
      seen.add(normalized);
      result.push(normalized);
    });
    return result;
  }

  const LATITUDE_FIELD_CANDIDATES = [
    "緯度",
    "decimalLatitude",
    "Latitude",
    "lat",
    "採集緯度"
  ];

  const LONGITUDE_FIELD_CANDIDATES = [
    "経度",
    "decimalLongitude",
    "Longitude",
    "lon",
    "lng",
    "採集経度"
  ];

  function compactHeaderName(value) {
    return String(value ?? "")
      .trim()
      .toLowerCase()
      .replace(/[\s_\-–—/\\（）()［］[\]{}:：,，.．]+/g, "");
  }

  function headerTokens(value) {
    return String(value ?? "")
      .trim()
      .toLowerCase()
      .split(/[\s_\-–—/\\（）()［］[\]{}:：,，.．]+/g)
      .filter(Boolean);
  }

  function findHeaderCandidate(headers, candidates) {
    const available = availableSheetFields(headers);
    const candidateCompacts = candidates.map((candidate) => compactHeaderName(candidate)).filter(Boolean);

    for (const header of available) {
      if (candidateCompacts.includes(compactHeaderName(header))) return header;
    }

    const longCandidates = candidateCompacts.filter((candidate) => candidate.length >= 4 || /[ぁ-んァ-ヶ一-龠]/.test(candidate));
    for (const header of available) {
      const compact = compactHeaderName(header);
      if (longCandidates.some((candidate) => compact.includes(candidate))) return header;
    }

    const shortCandidates = candidateCompacts.filter((candidate) => candidate.length < 4 && !/[ぁ-んァ-ヶ一-龠]/.test(candidate));
    for (const header of available) {
      const tokens = headerTokens(header);
      if (shortCandidates.some((candidate) => tokens.includes(candidate))) return header;
    }

    for (const header of available) {
      const compact = compactHeaderName(header);
      if (shortCandidates.some((candidate) => compact.startsWith(candidate) || compact.endsWith(candidate))) return header;
    }

    return "";
  }
  function inferLatitudeField(headers) {
    return findHeaderCandidate(headers, LATITUDE_FIELD_CANDIDATES);
  }

  function inferLongitudeField(headers) {
    return findHeaderCandidate(headers, LONGITUDE_FIELD_CANDIDATES);
  }

  function configuredLatitudeField(dataset, headers = state.headers) {
    const field = normalizeDatasetFieldName(dataset?.latitudeField);
    if (field) {
      const normalized = normalizeFieldsForHeaders([field], headers)[0];
      if (normalized) return normalized;
    }
    return inferLatitudeField(headers);
  }

  function configuredLongitudeField(dataset, headers = state.headers) {
    const field = normalizeDatasetFieldName(dataset?.longitudeField);
    if (field) {
      const normalized = normalizeFieldsForHeaders([field], headers)[0];
      if (normalized) return normalized;
    }
    return inferLongitudeField(headers);
  }

  function fieldLabel(field) {
    if (field === VIRTUAL_COORDINATES_FIELD) return "緯度, 経度 [Google Map]";
    return FILTER_FIELD_LABEL_OVERRIDES[field] || field;
  }

  function isTaxonField(field) {
    return normalizeDatasetFieldName(field) === COLUMNS.taxon;
  }

  function fieldsReferToSameHeader(a, b) {
    if (a === b) return true;
    return normalizeDatasetFieldName(a) === normalizeDatasetFieldName(b);
  }

  function availablePopupFields(headers) {
    const fields = [VIRTUAL_COORDINATES_FIELD, ...availableSheetFields(headers)];
    return uniqueFieldList(fields);
  }

  function availableSheetFields(headers) {
    return (headers || []).map(normalizeText).filter((header) => header && !header.startsWith("__raw_"));
  }

  function configuredPopupFields(dataset, headers) {
    const source = Array.isArray(dataset?.popupFields) ? dataset.popupFields : DEFAULT_POPUP_FIELDS;
    return normalizeFieldsForHeaders(source, headers, { includeVirtual: true });
  }

  function configuredFilterFields(dataset, headers) {
    const source = Array.isArray(dataset?.filterFields) ? dataset.filterFields : DEFAULT_FILTER_FIELDS;
    return normalizeFieldsForHeaders(source, headers);
  }

  function configuredColorField(dataset, headers = state.headers) {
    const colorField = normalizeDatasetFieldName(dataset?.colorField);
    if (!colorField) return "";
    return normalizeFieldsForHeaders([colorField], headers)[0] || "";
  }

  function configuredPopupTitleField(dataset, headers = state.headers) {
    const titleField = normalizeDatasetFieldName(dataset?.popupTitleField);
    if (!titleField) return "";
    return normalizeFieldsForHeaders([titleField], headers)[0] || "";
  }


  function normalizeDatasetInput(nameInput, urlInput, sheetNameInput, existing = null, settings = {}) {
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
    const hasPopupFields = Array.isArray(settings.popupFields);
    const hasFilterFields = Array.isArray(settings.filterFields);
    const latitudeField = normalizeDatasetFieldName(settings.latitudeField ?? existing?.latitudeField);
    const longitudeField = normalizeDatasetFieldName(settings.longitudeField ?? existing?.longitudeField);
    if (settings.requireLocationFields && (!latitudeField || !longitudeField)) {
      throw new Error("緯度として使う列と経度として使う列を選択してください。");
    }
    return {
      id: existing?.id || makeDatasetId(name),
      name,
      sourceUrl,
      spreadsheetId,
      sheetName,
      csvUrl: spreadsheetId ? "" : sourceUrl,
      latitudeField,
      longitudeField,
      colorField: normalizeDatasetFieldName(settings.colorField ?? existing?.colorField),
      popupTitleField: normalizeDatasetFieldName(settings.popupTitleField ?? existing?.popupTitleField),
      popupFields: hasPopupFields ? uniqueFieldList(settings.popupFields) : normalizeOptionalFieldList(existing?.popupFields),
      filterFields: hasFilterFields ? uniqueFieldList(settings.filterFields) : normalizeOptionalFieldList(existing?.filterFields),
      createdAt: existing?.createdAt || now,
      updatedAt: now
    };
  }

  let datasetModalResolve = null;
  let datasetModalExisting = null;
  let datasetModalHeaders = [];

  function setDatasetModalError(message) {
    if (!datasetModalError) return;
    const text = normalizeText(message);
    datasetModalError.textContent = text;
    datasetModalError.hidden = !text;
  }

  function setDatasetFieldsStatus(message, isError = false) {
    if (!datasetFieldsStatus) return;
    const text = normalizeText(message);
    datasetFieldsStatus.textContent = text;
    datasetFieldsStatus.hidden = !text;
    datasetFieldsStatus.classList.toggle("is-error", Boolean(isError));
  }

  function setDatasetFieldSettingsVisible(visible) {
    if (datasetFieldSettings) datasetFieldSettings.hidden = !visible;
  }

  function orderedFieldsForSettings(available, selected) {
    const availableSet = new Set(available);
    const ordered = [];
    const seen = new Set();
    selected.forEach((field) => {
      if (!availableSet.has(field) || seen.has(field)) return;
      seen.add(field);
      ordered.push(field);
    });
    available.forEach((field) => {
      if (seen.has(field)) return;
      seen.add(field);
      ordered.push(field);
    });
    return ordered;
  }

  function moveDatasetFieldItem(button, direction) {
    const item = button.closest(".dataset-field-item");
    const list = item?.parentElement;
    if (!item || !list) return;
    if (direction < 0 && item.previousElementSibling) {
      list.insertBefore(item, item.previousElementSibling);
    } else if (direction > 0 && item.nextElementSibling) {
      list.insertBefore(item.nextElementSibling, item);
    }
  }

  function renderDatasetCheckboxList(container, fields, selectedFields, { reorder = false } = {}) {
    if (!container) return;
    container.innerHTML = "";
    const selected = new Set(selectedFields);
    fields.forEach((field) => {
      const item = document.createElement("div");
      item.className = "dataset-field-item";
      item.dataset.field = field;

      const label = document.createElement("label");
      label.className = "dataset-field-check";

      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.value = field;
      checkbox.checked = selected.has(field);

      const text = document.createElement("span");
      text.textContent = fieldLabel(field);

      label.append(checkbox, text);
      item.appendChild(label);

      if (reorder) {
        const actions = document.createElement("div");
        actions.className = "dataset-field-order-actions";

        const up = document.createElement("button");
        up.type = "button";
        up.textContent = "↑";
        up.setAttribute("aria-label", fieldLabel(field) + "を上へ移動");
        up.addEventListener("click", () => moveDatasetFieldItem(up, -1));

        const down = document.createElement("button");
        down.type = "button";
        down.textContent = "↓";
        down.setAttribute("aria-label", fieldLabel(field) + "を下へ移動");
        down.addEventListener("click", () => moveDatasetFieldItem(down, 1));

        actions.append(up, down);
        item.appendChild(actions);
      }

      container.appendChild(item);
    });
  }

  function populateDatasetFieldSelect(select, fields, selectedValue, blankLabel = "選択しない") {
    if (!select) return;
    select.innerHTML = "";
    const blank = document.createElement("option");
    blank.value = "";
    blank.textContent = blankLabel;
    select.appendChild(blank);
    fields.forEach((field) => {
      const option = document.createElement("option");
      option.value = field;
      option.textContent = fieldLabel(field);
      select.appendChild(option);
    });
    select.value = selectedValue && fields.includes(selectedValue) ? selectedValue : "";
  }

  function renderDatasetFieldSettings(headers, existing = datasetModalExisting) {
    datasetModalHeaders = availableSheetFields(headers);
    if (!datasetModalHeaders.length) {
      setDatasetFieldSettingsVisible(false);
      setDatasetFieldsStatus("ヘッダーを読み込めませんでした。", true);
      return;
    }

    const latitudeField = configuredLatitudeField(existing, datasetModalHeaders);
    const longitudeField = configuredLongitudeField(existing, datasetModalHeaders);
    const colorField = configuredColorField(existing, datasetModalHeaders);
    const popupTitleField = configuredPopupTitleField(existing, datasetModalHeaders);
    populateDatasetFieldSelect(datasetLatitudeFieldSelect, datasetModalHeaders, latitudeField, "選択してください");
    populateDatasetFieldSelect(datasetLongitudeFieldSelect, datasetModalHeaders, longitudeField, "選択してください");
    populateDatasetFieldSelect(datasetColorFieldSelect, datasetModalHeaders, colorField);
    populateDatasetFieldSelect(datasetPopupTitleFieldSelect, datasetModalHeaders, popupTitleField);

    const popupAvailable = availablePopupFields(datasetModalHeaders);
    const popupSelected = configuredPopupFields(existing, datasetModalHeaders);
    renderDatasetCheckboxList(
      datasetPopupFieldsList,
      orderedFieldsForSettings(popupAvailable, popupSelected),
      popupSelected,
      { reorder: true }
    );

    const filterAvailable = availableSheetFields(datasetModalHeaders);
    const filterSelected = configuredFilterFields(existing, datasetModalHeaders);
    renderDatasetCheckboxList(datasetFilterFieldsList, filterAvailable, filterSelected);

    setDatasetFieldSettingsVisible(true);
    setDatasetFieldsStatus("列を読み込みました。位置情報に使用する列、タイトル項目、表示項目、フィルター項目、色分け項目を変更できます。", false);
  }

  async function loadDatasetFieldsForModal() {
    if (!datasetLoadFieldsButton) return;
    let draftDataset;
    try {
      draftDataset = normalizeDatasetInput(
        datasetNameInput?.value || "一時読み込み",
        datasetUrlInput?.value || "",
        datasetSheetInput?.value || "",
        datasetModalExisting,
        readDatasetModalSettings()
      );
    } catch (error) {
      setDatasetFieldsStatus(error.message, true);
      return;
    }

    datasetLoadFieldsButton.disabled = true;
    setDatasetFieldsStatus("列を読み込み中...", false);
    try {
      const { headers } = await loadSheetHeaders(draftDataset);
      renderDatasetFieldSettings(headers, draftDataset);
    } catch (error) {
      console.error(error);
      setDatasetFieldSettingsVisible(false);
      setDatasetFieldsStatus("列を読み込めませんでした: " + error.message, true);
    } finally {
      datasetLoadFieldsButton.disabled = false;
    }
  }

  function readCheckedDatasetFields(container) {
    if (!container || container.closest("[hidden]")) return null;
    return Array.from(container.querySelectorAll('input[type="checkbox"]'))
      .filter((input) => input.checked)
      .map((input) => input.value);
  }

  function readDatasetModalSettings() {
    const settingsVisible = datasetFieldSettings && !datasetFieldSettings.hidden;
    if (!settingsVisible) {
      return {
        latitudeField: datasetModalExisting?.latitudeField || "",
        longitudeField: datasetModalExisting?.longitudeField || "",
        colorField: datasetModalExisting?.colorField || "",
        popupTitleField: datasetModalExisting?.popupTitleField || "",
        popupFields: Array.isArray(datasetModalExisting?.popupFields) ? datasetModalExisting.popupFields : null,
        filterFields: Array.isArray(datasetModalExisting?.filterFields) ? datasetModalExisting.filterFields : null
      };
    }

    return {
      latitudeField: datasetLatitudeFieldSelect?.value || "",
      longitudeField: datasetLongitudeFieldSelect?.value || "",
      requireLocationFields: true,
      colorField: datasetColorFieldSelect?.value || "",
      popupTitleField: datasetPopupTitleFieldSelect?.value || "",
      popupFields: readCheckedDatasetFields(datasetPopupFieldsList) || [],
      filterFields: readCheckedDatasetFields(datasetFilterFieldsList) || []
    };
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
    datasetModalTitle.textContent = existing ? "調査結果を編集" : "調査結果を追加";
    datasetModalSubmitButton.textContent = existing ? "保存" : "登録";
    datasetNameInput.value = existing?.name || "";
    datasetUrlInput.value = existing?.sourceUrl || existing?.spreadsheetId || existing?.csvUrl || "";
    datasetSheetInput.value = existing?.sheetName || CONFIG.DEFAULT_SHEET_NAME || "シート1";
    datasetModalHeaders = [];
    setDatasetModalError("");
    setDatasetFieldsStatus("");
    setDatasetFieldSettingsVisible(false);

    const activeDataset = getActiveDataset();
    if (existing && activeDataset?.id === existing.id && state.headers.length) {
      renderDatasetFieldSettings(state.headers, existing);
    }

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
        datasetModalExisting,
        readDatasetModalSettings()
      );
      closeDatasetModal(dataset);
    } catch (error) {
      setDatasetModalError(error.message);
    }
  }

  function renderDatasetControls() {
    if (!datasetSelect) return;

    const activeDataset = getActiveDataset();
    const activeExists = !!activeDataset;

    if (!activeExists) {
      state.activeDatasetId = state.datasets[0]?.id || state.temporaryDataset?.id || "";
      if (state.datasets.some((dataset) => dataset.id === state.activeDatasetId)) {
        saveActiveDatasetId(state.activeDatasetId);
      } else {
        saveActiveDatasetId("");
      }
    }

    datasetSelect.innerHTML = "";

    const selectDatasets = [...state.datasets];
    if (state.temporaryDataset && !selectDatasets.some((dataset) => dataset.id === state.temporaryDataset.id)) {
      selectDatasets.push(state.temporaryDataset);
    }

    if (selectDatasets.length === 0) {
      const option = document.createElement("option");
      option.value = "";
      option.textContent = "未登録";
      datasetSelect.appendChild(option);
      datasetSelect.disabled = true;
      if (reloadButton) reloadButton.disabled = true;
      if (reloadCancelButton) reloadCancelButton.hidden = true;
      if (datasetEditButton) datasetEditButton.disabled = true;
      if (datasetDeleteButton) datasetDeleteButton.disabled = true;
      if (datasetShareButton) datasetShareButton.disabled = true;
      if (datasetRegisterButton) {
        datasetRegisterButton.classList.remove("is-registered", "is-registerable", "is-inactive");
        datasetRegisterButton.textContent = "登録";
        datasetRegisterButton.disabled = true;
        datasetRegisterButton.classList.add("is-inactive");
      }
      updateFilterSummary();
      updateFocusControlState();
      return;
    }

    selectDatasets.forEach((dataset) => {
      const option = document.createElement("option");
      option.value = dataset.id;
      option.textContent = isDatasetRegistered(dataset) ? dataset.name : dataset.name + "（未登録）";
      datasetSelect.appendChild(option);
    });

    const current = getActiveDataset();
    datasetSelect.value = state.activeDatasetId;
    datasetSelect.disabled = state.isLoading;

    const registered = current ? isDatasetRegistered(current) : false;
    const hasActive = !!current;

    if (reloadButton) reloadButton.disabled = !hasActive || state.isLoading;
    if (reloadCancelButton) reloadCancelButton.hidden = !state.isLoading;
    if (datasetEditButton) datasetEditButton.disabled = state.isLoading || !registered;
    if (datasetDeleteButton) datasetDeleteButton.disabled = state.isLoading || !registered;
    if (datasetShareButton) datasetShareButton.disabled = state.isLoading || !hasActive;
    if (filterButton) filterButton.disabled = state.isLoading || state.rows.length === 0;
    updateFilterSummary();
    updateFocusControlState();

    if (datasetRegisterButton) {
      datasetRegisterButton.classList.remove("is-registered", "is-registerable", "is-inactive");
      if (!hasActive) {
        datasetRegisterButton.textContent = "登録";
        datasetRegisterButton.disabled = true;
        datasetRegisterButton.classList.add("is-inactive");
      } else if (state.isLoading) {
        datasetRegisterButton.textContent = "登録";
        datasetRegisterButton.disabled = true;
        datasetRegisterButton.classList.add("is-inactive");
      } else if (registered) {
        datasetRegisterButton.textContent = "登録済み";
        datasetRegisterButton.disabled = true;
        datasetRegisterButton.classList.add("is-registered");
      } else {
        datasetRegisterButton.textContent = "登録";
        datasetRegisterButton.disabled = false;
        datasetRegisterButton.classList.add("is-registerable");
      }
    }
  }

  function resetMapDataForDatasetChange() {
    closePopup();
    clearMarkers();
    state.rows = [];
    state.filteredRows = [];
    state.filterOptions = new Map();
    state.filters = new Map();
    state.filterFields = [];
    state.headers = [];
    state.markerGroups = [];
    state.taxonColors = new Map();
    state.lastLoadedAt = null;
    state.skippedRows = 0;
    if (colorLegendTitleEl) colorLegendTitleEl.textContent = "色分け";
    if (taxonLegendEl) taxonLegendEl.innerHTML = "";
    if (recordTypeLegendEl) recordTypeLegendEl.innerHTML = "";
    updateFilterSummary();
  }

  function filterValue(value) {
    const text = normalizeText(value);
    return text || EMPTY_FILTER_VALUE;
  }

  function filterValueLabel(value) {
    return value === EMPTY_FILTER_VALUE ? "（空欄）" : value;
  }

  function getRecordFilterValue(record, fieldKey) {
    return filterValue(record?.__filterValues?.[fieldKey]);
  }

  function currentDatasetFilterStorageKey() {
    const dataset = getActiveDataset();
    if (!dataset) return "";
    const identity = datasetIdentityKey(dataset) || dataset.id || dataset.name;
    return `${FILTERS_STORAGE_KEY_PREFIX}:${hashString(identity).toString(36)}`;
  }

  function buildFilterFields(headers) {
    const dataset = getActiveDataset();
    return configuredFilterFields(dataset, headers)
      .map((header) => ({
        key: header,
        label: fieldLabel(header)
      }));
  }

  function buildFilterOptions(records) {
    const options = new Map();
    state.filterFields.forEach((field) => {
      const counts = new Map();
      records.forEach((record) => {
        const value = getRecordFilterValue(record, field.key);
        counts.set(value, (counts.get(value) || 0) + 1);
      });
      const values = [...counts.entries()]
        .map(([value, count]) => ({ value, count, label: filterValueLabel(value) }))
        .sort((a, b) => {
          if (a.value === EMPTY_FILTER_VALUE) return 1;
          if (b.value === EMPTY_FILTER_VALUE) return -1;
          return a.label.localeCompare(b.label, "ja");
        });
      options.set(field.key, values);
    });
    return options;
  }

  function defaultFiltersFromOptions(options) {
    const filters = new Map();
    state.filterFields.forEach((field) => {
      filters.set(field.key, new Set((options.get(field.key) || []).map((entry) => entry.value)));
    });
    return filters;
  }

  function loadFiltersForActiveDataset(options) {
    const key = currentDatasetFilterStorageKey();
    if (!key) return defaultFiltersFromOptions(options);

    let stored = null;
    try {
      stored = JSON.parse(window.localStorage.getItem(key) || "null");
    } catch (error) {
      console.warn("フィルター設定を読み取れませんでした。", error);
    }

    const filters = new Map();
    state.filterFields.forEach((field) => {
      const available = new Set((options.get(field.key) || []).map((entry) => entry.value));
      const savedValues = Array.isArray(stored?.[field.key]) ? stored[field.key].map(filterValue) : null;
      if (!savedValues) {
        filters.set(field.key, new Set(available));
        return;
      }
      filters.set(field.key, new Set(savedValues.filter((value) => available.has(value))));
    });
    return filters;
  }

  function saveFiltersForActiveDataset() {
    const key = currentDatasetFilterStorageKey();
    if (!key) return;
    const payload = {};
    state.filterFields.forEach((field) => {
      payload[field.key] = [...(state.filters.get(field.key) || new Set())];
    });
    window.localStorage.setItem(key, JSON.stringify(payload));
  }

  function applyFiltersToRows(records) {
    if (!records.length) return [];
    return records.filter((record) => state.filterFields.every((field) => {
      const selected = state.filters.get(field.key);
      if (!selected || selected.size === 0) return false;
      return selected.has(getRecordFilterValue(record, field.key));
    }));
  }

  function hasActiveFilter() {
    return state.filterFields.some((field) => {
      const options = state.filterOptions.get(field.key) || [];
      const selected = state.filters.get(field.key) || new Set();
      return options.length > 0 && selected.size < options.length;
    });
  }

  function updateFilterSummary() {
    const visible = state.filteredRows.length;
    const total = state.rows.length;
    if (filterSummary) {
      filterSummary.textContent = `表示中: ${visible} / ${total}件${hasActiveFilter() ? "（フィルター中）" : ""}`;
    }
    if (filterButton) {
      filterButton.disabled = total === 0;
      filterButton.classList.toggle("is-filtered", hasActiveFilter());
    }
    updateFocusControlState();
  }

  function refreshFilteredRows({ save = false, rerender = true, fit = false } = {}) {
    state.filteredRows = applyFiltersToRows(state.rows);
    assignTaxonColors(state.filteredRows);
    if (save) saveFiltersForActiveDataset();
    updateFilterSummary();

    if (rerender) {
      closePopup();
      renderLegends(state.filteredRows);
      if (fit) {
        fitBoundsIfNeeded(state.filteredRows).then(() => renderMarkers());
      } else {
        renderMarkers();
      }
      updateDataStatus();
    }
  }

  function buildFilterOptionCheckbox(field, entry, selected) {
    const label = document.createElement("label");
    label.className = "filter-option";
    label.dataset.filterLabel = `${entry.label} ${entry.value}`.toLocaleLowerCase("ja-JP");

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.dataset.filterField = field.key;
    checkbox.dataset.filterValue = entry.value;
    checkbox.checked = selected.has(entry.value);

    const text = document.createElement("span");
    text.className = "filter-option-label";
    text.textContent = entry.label;

    const count = document.createElement("span");
    count.className = "filter-option-count";
    count.textContent = String(entry.count);

    label.append(checkbox, text, count);
    return label;
  }

  function filterVisibleOptions(fieldset, query) {
    const normalizedQuery = normalizeText(query).toLocaleLowerCase("ja-JP");
    const options = Array.from(fieldset.querySelectorAll(".filter-option"));
    let visibleCount = 0;

    options.forEach((option) => {
      const targetText = option.dataset.filterLabel || "";
      const shouldShow = !normalizedQuery || targetText.includes(normalizedQuery);
      option.hidden = !shouldShow;
      if (shouldShow) visibleCount += 1;
    });

    const noMatch = fieldset.querySelector(".filter-no-match");
    if (noMatch) noMatch.hidden = visibleCount > 0;
  }

  function renderFilterModalOptions() {
    if (!filterOptionsEl) return;
    filterOptionsEl.innerHTML = "";

    state.filterFields.forEach((field) => {
      const fieldset = document.createElement("section");
      fieldset.className = "filter-fieldset";
      fieldset.dataset.filterKey = field.key;

      const header = document.createElement("div");
      header.className = "filter-fieldset-header";

      const title = document.createElement("h3");
      title.textContent = field.label;

      const fieldActions = document.createElement("div");
      fieldActions.className = "filter-field-actions";

      const selectAll = document.createElement("button");
      selectAll.type = "button";
      selectAll.textContent = "表示項目をすべて選択";
      selectAll.addEventListener("click", () => setFieldCheckboxes(field.key, true));

      const selectNone = document.createElement("button");
      selectNone.type = "button";
      selectNone.textContent = "表示項目を解除";
      selectNone.addEventListener("click", () => setFieldCheckboxes(field.key, false));

      fieldActions.append(selectAll, selectNone);
      header.append(title, fieldActions);
      fieldset.appendChild(header);

      const searchWrap = document.createElement("label");
      searchWrap.className = "filter-search-field";
      const searchInput = document.createElement("input");
      searchInput.type = "search";
      searchInput.autocomplete = "off";
      searchInput.placeholder = `${field.label}を検索`;
      searchInput.setAttribute("aria-label", `${field.label}を検索`);
      searchInput.addEventListener("input", () => filterVisibleOptions(fieldset, searchInput.value));
      searchWrap.appendChild(searchInput);
      fieldset.appendChild(searchWrap);

      const list = document.createElement("div");
      list.className = "filter-option-list";

      const options = state.filterOptions.get(field.key) || [];
      const selected = state.filters.get(field.key) || new Set();
      if (options.length === 0) {
        const empty = document.createElement("p");
        empty.className = "filter-empty";
        empty.textContent = "選択できる値がありません。";
        list.appendChild(empty);
      } else {
        options.forEach((entry) => list.appendChild(buildFilterOptionCheckbox(field, entry, selected)));
        const noMatch = document.createElement("p");
        noMatch.className = "filter-empty filter-no-match";
        noMatch.textContent = "検索に一致する項目がありません。";
        noMatch.hidden = true;
        list.appendChild(noMatch);
      }

      fieldset.appendChild(list);
      filterOptionsEl.appendChild(fieldset);
    });
  }

  function isDisplayedFilterInput(input) {
    const option = input?.closest?.(".filter-option");
    return !option || !option.hidden;
  }

  function setFieldCheckboxes(fieldKey, checked, options = {}) {
    if (!filterOptionsEl) return;
    const visibleOnly = options.visibleOnly !== false;
    filterOptionsEl.querySelectorAll(`input[type="checkbox"][data-filter-field="${fieldKey}"]`).forEach((input) => {
      if (!visibleOnly || isDisplayedFilterInput(input)) {
        input.checked = checked;
      }
    });
  }

  function setAllFilterCheckboxes(checked, options = {}) {
    if (!filterOptionsEl) return;
    const visibleOnly = options.visibleOnly !== false;
    filterOptionsEl.querySelectorAll('input[type="checkbox"][data-filter-field]').forEach((input) => {
      if (!visibleOnly || isDisplayedFilterInput(input)) {
        input.checked = checked;
      }
    });
  }

  function resetFilterModalCheckboxes() {
    setAllFilterCheckboxes(true, { visibleOnly: false });
  }

  function readFilterModalSelections() {
    const filters = new Map();
    state.filterFields.forEach((field) => filters.set(field.key, new Set()));
    if (!filterOptionsEl) return filters;
    filterOptionsEl.querySelectorAll('input[type="checkbox"][data-filter-field]').forEach((input) => {
      const fieldKey = input.dataset.filterField;
      const value = input.dataset.filterValue;
      if (input.checked && filters.has(fieldKey)) {
        filters.get(fieldKey).add(value);
      }
    });
    return filters;
  }

  function openFilterModal() {
    if (!filterModal || !filterOptionsEl || state.rows.length === 0) return;
    renderFilterModalOptions();
    filterModal.hidden = false;
    document.body.classList.add("modal-open");
    window.setTimeout(() => filterModalCloseButton?.focus(), 0);
  }

  function closeFilterModal() {
    if (!filterModal) return;
    filterModal.hidden = true;
    document.body.classList.remove("modal-open");
  }

  function applyFilterModal() {
    state.filters = readFilterModalSelections();
    refreshFilteredRows({ save: true, rerender: true, fit: true });
    closeFilterModal();
  }

  function updateDataStatus() {
    const dataset = getActiveDataset();
    if (!dataset || !state.lastLoadedAt) return;
    const skipped = Number(state.skippedRows) || 0;
    const skippedText = skipped > 0 ? `、${skipped}件スキップ` : "";
    setStatus(`「${dataset.name}」: ${state.filteredRows.length} / ${state.rows.length}件表示${skippedText} / ${state.lastLoadedAt.toLocaleTimeString("ja-JP", { hour: "2-digit", minute: "2-digit" })} 更新`);
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
    state.temporaryDataset = null;

    const sharedDataset = datasetFromShareParams();
    if (sharedDataset) {
      const existing = findDatasetByIdentity(sharedDataset);
      if (existing) {
        state.activeDatasetId = existing.id;
        saveActiveDatasetId(state.activeDatasetId);
      } else {
        state.temporaryDataset = sharedDataset;
        state.activeDatasetId = sharedDataset.id;
        saveActiveDatasetId("");
      }
      state.sharedDatasetIdFromUrl = sharedDataset.id;
    } else {
      const storedActiveId = loadActiveDatasetId();
      state.activeDatasetId = state.datasets.some((dataset) => dataset.id === storedActiveId)
        ? storedActiveId
        : (state.datasets[0]?.id || "");
      saveActiveDatasetId(state.activeDatasetId);
    }

    renderDatasetControls();
  }

  function setRegisterModalError(message) {
    if (!registerModalError) return;
    const text = normalizeText(message);
    registerModalError.textContent = text;
    registerModalError.hidden = !text;
  }

  function openRegisterModal() {
    if (!registerModal || !registerConfirmButton) return;
    const current = getActiveDataset();
    if (!current || isDatasetRegistered(current)) {
      renderDatasetControls();
      return;
    }
    setRegisterModalError("");
    registerModal.hidden = false;
    document.body.classList.add("modal-open");
    window.setTimeout(() => registerConfirmButton.focus(), 0);
  }

  function closeRegisterModal() {
    if (!registerModal) return;
    registerModal.hidden = true;
    document.body.classList.remove("modal-open");
    setRegisterModalError("");
  }

  function registerActiveDataset() {
    const current = getActiveDataset();
    if (!current || isDatasetRegistered(current)) {
      renderDatasetControls();
      closeRegisterModal();
      return;
    }

    const existing = findDatasetByIdentity(current);
    if (existing) {
      state.activeDatasetId = existing.id;
      if (state.temporaryDataset && state.temporaryDataset.id === current.id) {
        state.temporaryDataset = null;
      }
      saveActiveDatasetId(existing.id);
      renderDatasetControls();
      closeRegisterModal();
      return;
    }

    const now = new Date().toISOString();
    const registeredDataset = {
      ...current,
      id: makeDatasetId(current.name),
      createdAt: now,
      updatedAt: now
    };

    state.datasets.push(registeredDataset);
    state.activeDatasetId = registeredDataset.id;
    if (state.temporaryDataset && state.temporaryDataset.id === current.id) {
      state.temporaryDataset = null;
    }
    saveStoredDatasets();
    saveActiveDatasetId(registeredDataset.id);
    renderDatasetControls();
    closeRegisterModal();
  }

  function currentShareUrl() {
    const dataset = getActiveDataset();
    if (!dataset) return "";

    const url = new URL(window.location.href);
    url.search = "";
    url.hash = "";
    url.searchParams.set("fm_name", dataset.name);
    url.searchParams.set("fm_sheet", dataset.sheetName || CONFIG.DEFAULT_SHEET_NAME || "シート1");

    if (dataset.spreadsheetId) {
      url.searchParams.set("fm_sheet_id", dataset.spreadsheetId);
    } else if (dataset.csvUrl) {
      url.searchParams.set("fm_csv", dataset.csvUrl);
    }

    const latitudeField = normalizeDatasetFieldName(dataset.latitudeField || configuredLatitudeField(dataset, state.headers));
    if (latitudeField) url.searchParams.set("fm_lat", latitudeField);
    const longitudeField = normalizeDatasetFieldName(dataset.longitudeField || configuredLongitudeField(dataset, state.headers));
    if (longitudeField) url.searchParams.set("fm_lng", longitudeField);
    const colorField = normalizeDatasetFieldName(dataset.colorField);
    if (colorField) url.searchParams.set("fm_color", colorField);
    const popupTitleField = normalizeDatasetFieldName(dataset.popupTitleField);
    if (popupTitleField) url.searchParams.set("fm_popup_title", popupTitleField);
    if (Array.isArray(dataset.popupFields)) {
      url.searchParams.set("fm_popup", JSON.stringify(uniqueFieldList(dataset.popupFields)));
    }
    if (Array.isArray(dataset.filterFields)) {
      url.searchParams.set("fm_filter", JSON.stringify(uniqueFieldList(dataset.filterFields)));
    }

    const baseMap = basemapSelect?.value || CONFIG.INITIAL_BASEMAP || "detail";
    url.searchParams.set("fm_base", baseMap === "pale" ? "pale" : "detail");

    return url.toString();
  }

  function currentSourceUrl(dataset) {
    if (!dataset) return "";
    if (dataset.sourceUrl) return dataset.sourceUrl;
    if (dataset.spreadsheetId) return "https://docs.google.com/spreadsheets/d/" + dataset.spreadsheetId + "/edit";
    if (dataset.csvUrl) return dataset.csvUrl;
    return "";
  }

  function openHelpModal() {
    if (!helpModal) return;
    helpModal.hidden = false;
    document.body.classList.add("modal-open");
    window.setTimeout(() => {
      helpModalCloseButton?.focus();
    }, 0);
  }

  function closeHelpModal() {
    if (!helpModal) return;
    helpModal.hidden = true;
    document.body.classList.remove("modal-open");
  }

  function openShareModal() {
    const dataset = getActiveDataset();
    if (!dataset || !shareModal || !shareUrlOutput) return;
    shareUrlOutput.value = currentShareUrl();
    const sourceUrl = currentSourceUrl(dataset);
    if (shareSourceLink) {
      shareSourceLink.textContent = sourceUrl || "未設定";
      if (sourceUrl) {
        shareSourceLink.href = sourceUrl;
        shareSourceLink.removeAttribute("aria-disabled");
      } else {
        shareSourceLink.href = "#";
        shareSourceLink.setAttribute("aria-disabled", "true");
      }
    }
    if (shareModalMessage) shareModalMessage.textContent = "";
    shareModal.hidden = false;
    document.body.classList.add("modal-open");
    window.setTimeout(() => {
      shareUrlOutput.focus();
      shareUrlOutput.select();
    }, 0);
  }

  function closeShareModal() {
    if (!shareModal) return;
    shareModal.hidden = true;
    document.body.classList.remove("modal-open");
    if (shareModalMessage) shareModalMessage.textContent = "";
  }

  async function copyShareUrl() {
    const value = shareUrlOutput?.value || currentShareUrl();
    if (!value) return;
    try {
      await navigator.clipboard.writeText(value);
      if (shareModalMessage) shareModalMessage.textContent = "共有リンクをコピーしました。";
    } catch (error) {
      shareUrlOutput?.focus();
      shareUrlOutput?.select();
      if (shareModalMessage) shareModalMessage.textContent = "コピーできませんでした。共有リンクを選択してコピーしてください。";
    }
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  const escapeHTML = escapeHtml;

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
    return normalizeText(value) || "分類群未入力";
  }

  function colorKey(value) {
    return normalizeText(value) || "未設定";
  }

  function activeColorField() {
    return configuredColorField(getActiveDataset(), state.headers);
  }

  function activePopupTitleField() {
    return configuredPopupTitleField(getActiveDataset(), state.headers);
  }

  function getRecordColorValue(record) {
    const field = activeColorField();
    if (!field) return "";
    return colorKey(record?.__values?.[field]);
  }

  function colorForColorValue(value) {
    const text = normalizeText(value);
    if (!text) return DEFAULT_MARKER_COLOR;
    return state.taxonColors.get(text) || fallbackColorForColorValue(text);
  }

  function colorForRecord(record) {
    return colorForColorValue(getRecordColorValue(record));
  }

  function fallbackColorForColorValue(value) {
    return TAXON_COLOR_PALETTE[hashString(colorKey(value)) % TAXON_COLOR_PALETTE.length];
  }

  function assignTaxonColors(records) {
    const colorField = activeColorField();
    if (!colorField) {
      state.taxonColors = new Map();
      return;
    }

    const values = [...new Set(records.map((record) => getRecordColorValue(record)))]
      .sort((a, b) => a.localeCompare(b, "ja"));

    const adjacency = buildTaxonAdjacency(records, values);
    const counts = new Map();
    records.forEach((record) => {
      const value = getRecordColorValue(record);
      counts.set(value, (counts.get(value) || 0) + 1);
    });

    const order = values.sort((a, b) => {
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

    order.forEach((value) => {
      const neighbors = adjacency.get(value) || new Map();
      const paletteGroups = selectedColorGroups;
      const forceUniqueGlobalColors = paletteGroups.length >= order.length;
      let bestGroup = paletteGroups[0] || TAXON_COLOR_CANDIDATES[0];
      let bestScore = Number.NEGATIVE_INFINITY;

      paletteGroups.forEach((candidate, paletteIndex) => {
        let score = 0;

        // 近接関係のある値で同じ色グループを使うことを強く避けます。
        neighbors.forEach((weight, neighborValue) => {
          const neighborColor = assigned.get(neighborValue);
          const neighborGroup = assignedGroups.get(neighborValue);
          if (!neighborColor || !neighborGroup) return;

          if (neighborGroup === candidate.name) {
            score -= weight * 100000;
          } else {
            score += weight * colorDistance(candidate.color, neighborColor) * 12;
          }
        });

        // 近接していない値でも、すでに使われている色からできるだけ離します。
        const assignedColors = [...assigned.values()];
        if (assignedColors.length > 0) {
          const globalMinDistance = minColorDistance(candidate.color, assignedColors);
          score += globalMinDistance * 3;
          if (globalMinDistance < 35) score -= (35 - globalMinDistance) * 120;
        } else {
          score += 1000;
        }

        // 値の数が候補色数以下のときは、全体として色の再利用を禁止します。
        const usageCount = groupUsage.get(candidate.name) || 0;
        if (forceUniqueGlobalColors && usageCount > 0) {
          score -= 1000000;
        }

        score -= usageCount * 300;
        score -= paletteIndex * 0.001;

        if (score > bestScore) {
          bestScore = score;
          bestGroup = candidate;
        }
      });

      assigned.set(value, bestGroup.color);
      assignedGroups.set(value, bestGroup.name);
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

  function buildTaxonAdjacency(records, values) {
    const adjacency = new Map(values.map((value) => [value, new Map()]));
    const nearestCount = Math.max(1, Number(CONFIG.COLOR_NEAREST_DIFFERENT_TAXA) || 5);

    records.forEach((record, index) => {
      const sourceValue = getRecordColorValue(record);
      if (!sourceValue) return;
      const nearest = [];

      records.forEach((otherRecord, otherIndex) => {
        if (index === otherIndex) return;

        const targetValue = getRecordColorValue(otherRecord);
        if (!targetValue || sourceValue === targetValue) return;

        const distance = haversineMeters(record.latitude, record.longitude, otherRecord.latitude, otherRecord.longitude);
        if (!Number.isFinite(distance)) return;

        nearest.push({ value: targetValue, distance });
      });

      nearest.sort((a, b) => a.distance - b.distance);

      const addedValues = new Set();
      for (const neighbor of nearest) {
        if (addedValues.has(neighbor.value)) continue;

        addedValues.add(neighbor.value);
        const rank = addedValues.size;
        const weight = nearestCount + 1 - rank;
        addAdjacencyWeight(adjacency, sourceValue, neighbor.value, weight);

        if (addedValues.size >= nearestCount) break;
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
    if (text === "作成者の記録" || text === "本調査") return "thisStudy";
    if (text === "不確かな記録") return "uncertain";
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

    state.map.addControl(createFocusControl(), "top-right");

    if (basemapSelect) {
      const sharedBaseMap = new URLSearchParams(window.location.search).get("fm_base");
      basemapSelect.value = sharedBaseMap === "pale" ? "pale" : (CONFIG.INITIAL_BASEMAP === "pale" ? "pale" : "detail");
      setBaseMap(basemapSelect.value);
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
      setStatus("調査結果を追加してください。");
      renderDatasetControls();
      return;
    }

    if (state.currentLoadJob) {
      cancelActiveLoad(false);
    }

    const job = createLoadJob();
    state.currentLoadJob = job;
    setLoadingUi(true);
    renderDatasetControls();
    addLoadTimer(job, () => {
      setStatus("読み込みに時間がかかっています。使用しないポップアップ項目・フィルター項目を減らすと、読み込みが速くなる場合があります。中止することもできます。");
    }, LOAD_SLOW_NOTICE_MS);

    try {
      setStatus(`「${dataset.name}」の列情報を取得しています...`);
      const { rows, headers, selectedFields } = await loadDisplayRows(dataset, job);
      throwIfLoadCancelled(job);

      setStatus(`地図用データを作成しています...（読み込み列数: ${selectedFields.length}）`);
      validateHeaders(headers, dataset);
      state.headers = availableSheetFields(headers);
      state.filterFields = buildFilterFields(state.headers);

      const records = normalizeRows(rows, state.headers, dataset);
      state.rows = records;
      state.filterOptions = buildFilterOptions(records);
      state.filters = loadFiltersForActiveDataset(state.filterOptions);
      closePopup();
      state.filteredRows = applyFiltersToRows(records);
      assignTaxonColors(state.filteredRows);
      renderLegends(state.filteredRows);

      setStatus("地図の表示範囲を調整しています...");
      await fitBoundsIfNeeded(state.filteredRows);
      throwIfLoadCancelled(job);

      setStatus("マーカーを描画しています...");
      renderMarkers();

      state.lastLoadedAt = new Date();
      state.skippedRows = rows.length - records.length;
      updateFilterSummary();
      updateDataStatus();
      state.hasLoadedOnce = true;
    } catch (error) {
      if (isAbortError(error)) {
        setStatus("読み込みを中止しました。");
      } else {
        console.error(error);
        setStatus(`読み込み失敗: ${error.message}`);
        window.alert(`データを読み込めませんでした。\n\n${error.message}`);
      }
    } finally {
      if (state.currentLoadJob === job) {
        clearLoadJob(job);
        state.currentLoadJob = null;
      }
      setLoadingUi(false);
      renderDatasetControls();
    }
  }

  function fitToRecords(records, { respectAutoFit = false, duration = 300 } = {}) {
    const targetRecords = Array.isArray(records)
      ? records.filter((record) => Number.isFinite(record?.latitude) && Number.isFinite(record?.longitude))
      : [];
    if (!state.map || targetRecords.length === 0) return Promise.resolve(false);
    if (respectAutoFit && !CONFIG.AUTO_FIT_BOUNDS) return Promise.resolve(false);

    return new Promise((resolve) => {
      let resolved = false;
      const finish = () => {
        if (resolved) return;
        resolved = true;
        resolve(true);
      };
      state.map.once("moveend", finish);

      if (targetRecords.length === 1) {
        const record = targetRecords[0];
        const defaultZoom = Number(CONFIG.FOCUS_SINGLE_MARKER_ZOOM) || 14;
        const maxZoom = Number(CONFIG.FOCUS_SINGLE_MARKER_MAX_ZOOM) || 16;
        state.map.easeTo({
          center: [record.longitude, record.latitude],
          zoom: Math.min(Math.max(state.map.getZoom(), defaultZoom), maxZoom),
          duration
        });
      } else {
        const bounds = new maplibregl.LngLatBounds();
        targetRecords.forEach((record) => {
          bounds.extend([record.longitude, record.latitude]);
        });
        state.map.fitBounds(bounds, {
          padding: { top: 90, right: 45, bottom: 90, left: 45 },
          maxZoom: 12,
          duration
        });
      }
      window.setTimeout(finish, Math.max(900, duration + 300));
    });
  }

  function fitBoundsIfNeeded(records) {
    return fitToRecords(records, {
      respectAutoFit: true,
      duration: state.hasLoadedOnce ? 250 : 600
    });
  }

  async function focusVisibleRecords() {
    if (state.isLoading) return;
    if (!state.filteredRows.length) {
      setStatus("表示中の記録がありません。");
      return;
    }
    closePopup();
    setStatus("表示中の記録範囲へ移動しています...");
    await fitToRecords(state.filteredRows, { respectAutoFit: false, duration: 350 });
    updateDataStatus();
  }

  function updateFocusControlState() {
    if (!state.focusButton) return;
    state.focusButton.disabled = state.isLoading || state.filteredRows.length === 0;
    state.focusButton.setAttribute("aria-disabled", state.focusButton.disabled ? "true" : "false");
  }

  function createFocusControl() {
    let container = null;
    return {
      onAdd() {
        container = document.createElement("div");
        container.className = "maplibregl-ctrl maplibregl-ctrl-group field-map-focus-control";

        const button = document.createElement("button");
        button.type = "button";
        button.className = "field-map-focus-button";
        button.title = "表示中の記録範囲へ移動";
        button.setAttribute("aria-label", "表示中の記録範囲へ移動");
        button.textContent = "◎";
        button.addEventListener("click", (event) => {
          event.preventDefault();
          event.stopPropagation();
          focusVisibleRecords();
        });

        container.appendChild(button);
        state.focusButton = button;
        updateFocusControlState();
        return container;
      },
      onRemove() {
        if (container?.parentNode) container.parentNode.removeChild(container);
        state.focusButton = null;
        container = null;
      }
    };
  }

  async function loadDisplayRows(dataset, job) {
    if (dataset.csvUrl && dataset.csvUrl.trim()) {
      setStatus(`「${dataset.name}」のCSVを読み込んでいます...`);
      const full = await loadCsvRows(dataset.csvUrl.trim(), { job });
      throwIfLoadCancelled(job);
      validateHeaders(full.headers, dataset);
      const selectedFields = buildDataFieldsForDataset(dataset, full.headers);
      setStatus(`使用する ${selectedFields.length} 列だけを地図表示用に抽出しています...`);
      return {
        headers: selectedFields,
        rows: selectRowsByFields(full.rows, selectedFields),
        selectedFields
      };
    }

    const headerResult = await loadGvizHeaders(dataset.spreadsheetId, dataset.sheetName, { job });
    throwIfLoadCancelled(job);
    validateHeaders(headerResult.headers, dataset);
    const selectedFields = buildDataFieldsForDataset(dataset, headerResult.headers);
    setStatus(`「${dataset.name}」の使用する ${selectedFields.length} 列を読み込んでいます...`);
    const rowsResult = await loadGvizRows(dataset.spreadsheetId, dataset.sheetName, {
      job,
      fields: selectedFields,
      sourceHeaders: headerResult.headers
    });
    return {
      headers: rowsResult.headers,
      rows: rowsResult.rows,
      selectedFields
    };
  }

  async function loadSheetHeaders(dataset) {
    if (dataset.csvUrl && dataset.csvUrl.trim()) {
      const { headers } = await loadCsvRows(dataset.csvUrl.trim(), {});
      return { headers };
    }
    return loadGvizHeaders(dataset.spreadsheetId, dataset.sheetName, {});
  }

  async function loadSheetRows(dataset, options = {}) {
    if (dataset.csvUrl && dataset.csvUrl.trim()) {
      return loadCsvRows(dataset.csvUrl.trim(), options);
    }
    return loadGvizRows(dataset.spreadsheetId, dataset.sheetName, options);
  }

  async function loadCsvRows(url, { fields = null, job = null } = {}) {
    const separator = url.includes("?") ? "&" : "?";
    const fetchOptions = { cache: "no-store" };
    let didTimeout = false;
    let timeoutId = null;
    if (job?.abortController) {
      fetchOptions.signal = job.abortController.signal;
      timeoutId = window.setTimeout(() => {
        didTimeout = true;
        job.abortController.abort();
      }, LOAD_TIMEOUT_MS);
    }

    let response;
    try {
      response = await fetch(`${url}${separator}_=${Date.now()}`, fetchOptions);
    } catch (error) {
      if (timeoutId) window.clearTimeout(timeoutId);
      if (job?.aborted) throw makeAbortError();
      if (didTimeout) {
        throw new Error("CSVの読み込みがタイムアウトしました。使用しないポップアップ項目・フィルター項目を減らすか、データを分割すると改善する場合があります。");
      }
      throw error;
    } finally {
      if (timeoutId) window.clearTimeout(timeoutId);
    }

    throwIfLoadCancelled(job);
    if (!response.ok) {
      throw new Error(`CSVの取得に失敗しました: HTTP ${response.status}`);
    }
    const text = await response.text();
    throwIfLoadCancelled(job);
    const table = parseCsv(text);
    if (table.length === 0) {
      throw new Error("CSVに行がありません。");
    }

    const allHeaders = table[0].map(normalizeText);
    const selectedHeaders = fields ? normalizeFieldsForHeaders(fields, allHeaders) : availableSheetFields(allHeaders);
    const selectedSet = new Set(selectedHeaders);
    const rows = table.slice(1).map((cells) => {
      const row = {};
      allHeaders.forEach((header, index) => {
        if (!fields || selectedSet.has(header)) {
          row[header] = cells[index] ?? "";
        }
      });
      return row;
    });

    return { headers: selectedHeaders, rows };
  }

  function loadGvizHeaders(sheetId, sheetName, options = {}) {
    return loadGvizTable(sheetId, sheetName, {
      ...options,
      tq: "select * limit 0",
      timeoutMessage: "Googleスプレッドシートの列情報取得がタイムアウトしました。共有設定またはシート名を確認してください。"
    });
  }

  function loadGvizRows(sheetId, sheetName, { fields = null, sourceHeaders = null, job = null } = {}) {
    const tq = fields && sourceHeaders ? buildGvizSelectQuery(sourceHeaders, fields) : "select *";
    return loadGvizTable(sheetId, sheetName, {
      job,
      tq,
      timeoutMessage: "Googleスプレッドシートのデータ読み込みがタイムアウトしました。使用しないポップアップ項目・フィルター項目を減らすか、シートを分割すると改善する場合があります。"
    });
  }

  function loadGvizTable(sheetId, sheetName, { tq = "select *", timeoutMessage = "Googleスプレッドシートの読み込みがタイムアウトしました。", job = null } = {}) {
    return new Promise((resolve, reject) => {
      if (!sheetId) {
        reject(new Error("SHEET_ID が設定されていません。"));
        return;
      }
      if (job?.aborted) {
        reject(makeAbortError());
        return;
      }

      const callbackName = `__fieldMapGviz_${Date.now()}_${Math.floor(Math.random() * 100000)}`;
      const script = document.createElement("script");
      const query = new URLSearchParams({
        sheet: sheetName || "",
        headers: "1",
        tqx: `out:json;responseHandler:${callbackName}`,
        tq,
        _: String(Date.now())
      });

      let settled = false;
      const timeoutId = window.setTimeout(() => {
        cleanup();
        reject(new Error(timeoutMessage));
      }, LOAD_TIMEOUT_MS);

      function cleanup() {
        if (settled) return;
        settled = true;
        window.clearTimeout(timeoutId);
        delete window[callbackName];
        if (script.parentNode) script.remove();
        if (job) job.cleanupHandlers.delete(cancelHandler);
      }

      function cancelHandler() {
        cleanup();
        reject(makeAbortError());
      }

      if (job) job.cleanupHandlers.add(cancelHandler);

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

  function columnIdForIndex(index) {
    let number = index + 1;
    let id = "";
    while (number > 0) {
      const rem = (number - 1) % 26;
      id = String.fromCharCode(65 + rem) + id;
      number = Math.floor((number - 1) / 26);
    }
    return id;
  }

  function buildGvizSelectQuery(sourceHeaders, fields) {
    const headers = (sourceHeaders || []).map(normalizeText);
    const selected = normalizeFieldsForHeaders(fields, headers);
    const columns = selected
      .map((field) => headers.indexOf(field))
      .filter((index) => index >= 0)
      .map(columnIdForIndex);
    if (!columns.length) return "select *";
    return `select ${columns.join(",")}`;
  }

  function selectRowsByFields(rows, fields) {
    const selected = uniqueFieldList(fields);
    return rows.map((row) => {
      const obj = {};
      selected.forEach((field) => {
        obj[field] = row[field] ?? "";
        const rawKey = `__raw_${field}`;
        if (Object.prototype.hasOwnProperty.call(row, rawKey)) obj[rawKey] = row[rawKey];
      });
      return obj;
    });
  }

  function addResolvedField(result, headers, field) {
    if (!field || field === VIRTUAL_COORDINATES_FIELD) return;
    const resolved = normalizeFieldsForHeaders([field], headers)[0];
    if (resolved && !result.includes(resolved)) result.push(resolved);
  }

  function buildDataFieldsForDataset(dataset, headers) {
    const result = [];
    addResolvedField(result, headers, configuredLatitudeField(dataset, headers));
    addResolvedField(result, headers, configuredLongitudeField(dataset, headers));
    addResolvedField(result, headers, COLUMNS.recordType);

    addResolvedField(result, headers, configuredColorField(dataset, headers));
    addResolvedField(result, headers, configuredPopupTitleField(dataset, headers));

    configuredPopupFields(dataset, headers).forEach((field) => addResolvedField(result, headers, field));
    configuredFilterFields(dataset, headers).forEach((field) => addResolvedField(result, headers, field));

    if (result.includes(COLUMNS.date) || result.includes(LEGACY_DATE_HEADER)) {
      addResolvedField(result, headers, COLUMNS.year);
      addResolvedField(result, headers, COLUMNS.month);
      addResolvedField(result, headers, COLUMNS.day);
    }

    return result;
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

  function validateHeaders(headers, dataset = getActiveDataset()) {
    const headerSet = new Set(headers.map(normalizeText));
    const hasTaxonHeader = headerSet.has(COLUMNS.taxon) || headerSet.has(LEGACY_TAXON_HEADER);
    if (!hasTaxonHeader && headerSet.has(REJECTED_TAXON_HEADER)) {
      throw new Error("ヘッダーが旧形式です。「Taxon」ではなく「分類群」に変更してから、再読み込みしてください。");
    }

    const latitudeField = configuredLatitudeField(dataset, headers);
    const longitudeField = configuredLongitudeField(dataset, headers);
    if (!latitudeField || !longitudeField) {
      throw new Error("位置情報に使用する列を特定できません。データの編集画面で「緯度として使う列」と「経度として使う列」を選択してください。");
    }
  }

  function normalizeRows(rows, headers, dataset = getActiveDataset()) {
    const filterHeaders = (headers || []).map(normalizeText).filter(Boolean);
    const latitudeField = configuredLatitudeField(dataset, headers);
    const longitudeField = configuredLongitudeField(dataset, headers);
    return rows.map((row, index) => {
      const latitude = parseNumber(row[`__raw_${latitudeField}`] ?? row[latitudeField]);
      const longitude = parseNumber(row[`__raw_${longitudeField}`] ?? row[longitudeField]);
      const values = {};
      filterHeaders.forEach((header) => {
        values[header] = normalizeText(row[header]);
      });
      if (!values[COLUMNS.taxon] && values[LEGACY_TAXON_HEADER]) {
        values[COLUMNS.taxon] = values[LEGACY_TAXON_HEADER];
      }
      if (!values[COLUMNS.latitude] && values[LEGACY_LATITUDE_HEADER]) {
        values[COLUMNS.latitude] = values[LEGACY_LATITUDE_HEADER];
      }
      if (!values[COLUMNS.longitude] && values[LEGACY_LONGITUDE_HEADER]) {
        values[COLUMNS.longitude] = values[LEGACY_LONGITUDE_HEADER];
      }
      if (!values[COLUMNS.date] && values[LEGACY_DATE_HEADER]) {
        values[COLUMNS.date] = values[LEGACY_DATE_HEADER];
      }

      return {
        __values: values,
        __filterValues: values,
        __index: index,
        id: values[COLUMNS.id] || "",
        taxon: values[COLUMNS.taxon] || values[LEGACY_TAXON_HEADER] || "",
        recordType: values[COLUMNS.recordType] || "",
        repository: values[COLUMNS.repository] || "",
        specimenId: values[COLUMNS.specimenId] || "",
        locality: values[COLUMNS.locality] || "",
        latitude,
        longitude,
        year: values[COLUMNS.year] || "",
        month: values[COLUMNS.month] || "",
        day: values[COLUMNS.day] || "",
        date: values[COLUMNS.date] || "",
        remarks: values[COLUMNS.remarks] || ""
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
    if (!state.filteredRows.length) {
      clearMarkers();
      return;
    }

    clearMarkers();
    const groups = buildMarkerGroups(state.filteredRows);
    state.markerGroups = groups;

    // 優先度の低いマーカーから描画し、高いマーカーを前面に出す。
    const renderGroups = [...groups].sort((a, b) => getMarkerGroupPriority(a) - getMarkerGroupPriority(b));

    renderGroups.forEach((group) => {
      const element = createMarkerElement(group);
      element.style.zIndex = String(100 + getMarkerGroupPriority(group));
      const marker = new maplibregl.Marker({ element })
        .setLngLat(markerGroupLngLat(group))
        .addTo(state.map);

      element.addEventListener("click", (event) => {
        event.stopPropagation();
        openPopupForMarkerGroup(group);
      });

      state.markers.push(marker);
    });
  }

  function openPopupForMarkerGroup(group) {
    if (!group) return;
    state.currentPopupRecords = popupRecordsForNearbyMarkerGroups(group);
    state.currentPopupIndex = 0;
    showPopup(0, false);
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

    groups.forEach(finalizeMarkerGroup);
    return groups.sort(compareMarkerGroups);
  }

  function createMarkerElement(group) {
    finalizeMarkerGroup(group);
    const record = group.representative;
    const segments = group.segments || getMarkerSegments(group.records);
    const kind = markerKindForGroup(group);
    const wrapper = document.createElement("button");
    wrapper.type = "button";
    wrapper.className = `record-marker kind-${kind}${segments.length > 1 ? " multi-taxon" : ""}`;
    wrapper.setAttribute("aria-label", `${record.taxon || "記録"} ${record.recordType || ""}`.trim());
    wrapper.innerHTML = markerSvg(kind, segments);

    return wrapper;
  }

  function finalizeMarkerGroup(group) {
    if (!group || !Array.isArray(group.records) || !group.records.length) return group;
    group.records.sort(compareRecordsForRepresentative);
    group.segments = getMarkerSegments(group.records);
    const bestSegment = bestMarkerSegment(group.segments);
    group.markerPriority = getMarkerSegmentPriority(bestSegment);
    group.representative = bestSegment?.representative || group.records[0];
    group.longitude = group.representative.longitude;
    group.latitude = group.representative.latitude;
    group.pixel = state.map ? state.map.project(markerGroupLngLat(group)) : group.pixel;
    return group;
  }

  function markerGroupLngLat(group) {
    const longitude = Number(group?.longitude ?? group?.representative?.longitude);
    const latitude = Number(group?.latitude ?? group?.representative?.latitude);
    return [longitude, latitude];
  }

  function markerKindForGroup(group) {
    const segment = bestMarkerSegment(group?.segments || getMarkerSegments(group?.records || []));
    return segment?.kind || recordKind(group?.representative?.recordType) || "unknown";
  }

  function bestMarkerSegment(segments) {
    return [...(segments || [])].sort(compareMarkerSegments)[0] || null;
  }

  function markerShapePriority(kind) {
    if (kind === "type") return 3;
    if (kind === "synonymType") return 2;
    if (kind === "thisStudy" || kind === "literature") return 1;
    return 0;
  }

  function getMarkerSegmentPriority(segment) {
    if (!segment) return 0;
    const shapePriority = markerShapePriority(segment.kind);
    const fillPriority = segment.hollow ? 0 : 1;
    return shapePriority * 10 + fillPriority;
  }

  function compareMarkerSegments(a, b) {
    const priorityDiff = getMarkerSegmentPriority(b) - getMarkerSegmentPriority(a);
    if (priorityDiff !== 0) return priorityDiff;
    return compareRecordsForRepresentative(a?.representative || {}, b?.representative || {});
  }

  function getMarkerGroupPriority(group) {
    if (!group) return 0;
    if (Number.isFinite(group.markerPriority)) return group.markerPriority;
    finalizeMarkerGroup(group);
    return group.markerPriority || 0;
  }

  function compareMarkerGroups(a, b) {
    const priorityDiff = getMarkerGroupPriority(b) - getMarkerGroupPriority(a);
    if (priorityDiff !== 0) return priorityDiff;
    const aLngLat = markerGroupLngLat(a);
    const bLngLat = markerGroupLngLat(b);
    const latDiff = aLngLat[1] - bLngLat[1];
    if (latDiff !== 0) return latDiff;
    return aLngLat[0] - bLngLat[0];
  }

  function popupRecordsForNearbyMarkerGroups(clickedGroup) {
    const groups = nearbyMarkerGroups(clickedGroup);
    const entries = [];
    groups.forEach((group) => {
      finalizeMarkerGroup(group);
      const lngLat = markerGroupLngLat(group);
      [...group.records].sort(compareRecordsForRepresentative).forEach((record) => {
        entries.push({
          ...record,
          __popupLongitude: lngLat[0],
          __popupLatitude: lngLat[1],
          __popupMarkerPriority: getMarkerGroupPriority(group)
        });
      });
    });
    return entries;
  }

  function nearbyMarkerGroups(clickedGroup) {
    if (!state.map || !clickedGroup) return clickedGroup ? [clickedGroup] : [];
    const clickedLngLat = markerGroupLngLat(clickedGroup);
    const clickedPixel = state.map.project(clickedLngLat);
    const candidates = state.markerGroups
      .map((group) => {
        finalizeMarkerGroup(group);
        const pixel = state.map.project(markerGroupLngLat(group));
        const dx = pixel.x - clickedPixel.x;
        const dy = pixel.y - clickedPixel.y;
        return { group, distance: Math.sqrt(dx * dx + dy * dy) };
      })
      .filter((entry) => entry.group === clickedGroup || entry.distance <= POPUP_NEARBY_MARKER_PIXELS);

    candidates.sort((a, b) => {
      if (a.group === clickedGroup) return -1;
      if (b.group === clickedGroup) return 1;
      const priorityDiff = getMarkerGroupPriority(b.group) - getMarkerGroupPriority(a.group);
      if (priorityDiff !== 0) return priorityDiff;
      return a.distance - b.distance;
    });

    return candidates.map((entry) => entry.group);
  }

  function getMarkerSegments(records) {
    const colorField = activeColorField();
    if (!colorField) {
      const sortedRecords = [...records].sort(compareRecordsForRepresentative);
      const representative = sortedRecords[0];
      const kind = recordKind(representative.recordType);
      const hasThisStudy = records.some((record) => recordKind(record.recordType) === "thisStudy");
      return [{
        value: "",
        color: DEFAULT_MARKER_COLOR,
        kind,
        hollow: shouldUseHollowMarkerInterior(kind, hasThisStudy),
        representative
      }];
    }

    const byValue = new Map();
    records.forEach((record) => {
      const key = getRecordColorValue(record);
      if (!byValue.has(key)) {
        byValue.set(key, []);
      }
      byValue.get(key).push(record);
    });

    return [...byValue.entries()]
      .map(([value, valueRecords]) => {
        const sortedRecords = [...valueRecords].sort(compareRecordsForRepresentative);
        const representative = sortedRecords[0];
        const kind = recordKind(representative.recordType);
        const hasThisStudyForSameValue = valueRecords.some((record) => recordKind(record.recordType) === "thisStudy");
        return {
          value,
          color: colorForColorValue(value),
          kind,
          hollow: shouldUseHollowMarkerInterior(kind, hasThisStudyForSameValue),
          representative
        };
      })
      .sort(compareMarkerSegments);
  }

  function shouldUseHollowMarkerInterior(kind, hasThisStudyForSameTaxon) {
    if (kind === "literature") return true;
    if (kind === "type" || kind === "synonymType") return !hasThisStudyForSameTaxon;
    return false;
  }

  function markerSvg(kind, segmentsOrColor) {
    const segments = normalizeMarkerSegments(segmentsOrColor);
    if (segments.length <= 1) {
      const segment = segments[0] || { color: "#666666", kind, hollow: false };
      return singleColorMarkerSvg(segment.kind || kind, segment.color || "#666666", Boolean(segment.hollow));
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
        hollow: Boolean(item?.hollow),
        taxon: item?.taxon || "",
        representative: item?.representative || null
      };
    }).filter((item) => item.color);
  }

  function singleColorMarkerSvg(kind, color, hollow = false) {
    const stroke = "rgba(0,0,0,0.72)";
    const white = "rgba(255,255,255,0.94)";
    const innerStroke = "rgba(0,0,0,0.18)";

    if (kind === "type") {
      const starPoints = "12,1.6 14.8,8.4 22.1,8.9 16.5,13.6 18.2,20.8 12,17 5.8,20.8 7.5,13.6 1.9,8.9 9.2,8.4";
      const innerStarTransform = "translate(12 12) scale(0.50) translate(-12 -12)";
      if (hollow) {
        return `
          <svg viewBox="0 0 24 24" role="img" aria-hidden="true">
            <polygon points="${starPoints}" fill="${color}" stroke="${stroke}" stroke-width="1.4" />
            <polygon points="${starPoints}" transform="${innerStarTransform}"
              fill="${white}" stroke="${innerStroke}" stroke-width="0.45" />
          </svg>`;
      }
      return `
        <svg viewBox="0 0 24 24" role="img" aria-hidden="true">
          <polygon points="${starPoints}"
            fill="${color}" stroke="${stroke}" stroke-width="1.4" />
        </svg>`;
    }

    if (kind === "synonymType") {
      if (hollow) {
        return `
          <svg viewBox="0 0 24 24" role="img" aria-hidden="true">
            <rect x="5" y="5" width="14" height="14" rx="1.8"
              fill="${color}" stroke="${stroke}" stroke-width="1.7" />
            <rect x="8.3" y="8.3" width="7.4" height="7.4" rx="1.0"
              fill="${white}" stroke="${innerStroke}" stroke-width="0.45" />
          </svg>`;
      }
      return `
        <svg viewBox="0 0 24 24" role="img" aria-hidden="true">
          <rect x="5" y="5" width="14" height="14" rx="1.8"
            fill="${color}" stroke="${stroke}" stroke-width="1.7" />
        </svg>`;
    }

    if (kind === "literature") {
      return `
        <svg viewBox="0 0 24 24" role="img" aria-hidden="true">
          <circle cx="12" cy="12" r="8.2"
            fill="${color}" stroke="${stroke}" stroke-width="1.7" />
          <circle cx="12" cy="12" r="4.5"
            fill="${white}" stroke="${innerStroke}" stroke-width="0.45" />
        </svg>`;
    }

    if (kind === "uncertain") {
      return `
        <svg viewBox="0 0 24 24" role="img" aria-hidden="true">
          <line x1="5.2" y1="5.2" x2="18.8" y2="18.8"
            stroke="${stroke}" stroke-width="6.0" stroke-linecap="round" />
          <line x1="18.8" y1="5.2" x2="5.2" y2="18.8"
            stroke="${stroke}" stroke-width="6.0" stroke-linecap="round" />
          <line x1="5.2" y1="5.2" x2="18.8" y2="18.8"
            stroke="${color}" stroke-width="4.1" stroke-linecap="round" />
          <line x1="18.8" y1="5.2" x2="5.2" y2="18.8"
            stroke="${color}" stroke-width="4.1" stroke-linecap="round" />
        </svg>`;
    }

    if (kind === "unknown") {
      return `
        <svg viewBox="0 0 24 24" role="img" aria-hidden="true">
          <polygon points="12,3.8 21,19.2 3,19.2"
            fill="${color}" stroke="${stroke}" stroke-width="1.7" stroke-linejoin="round" />
        </svg>`;
    }

    return `
      <svg viewBox="0 0 24 24" role="img" aria-hidden="true">
        <circle cx="12" cy="12" r="8.2"
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
      const innerStarTransform = "translate(12 12) scale(0.50) translate(-12 -12)";
      return `
        <svg viewBox="0 0 24 24" role="img" aria-hidden="true">
          <defs>
            <clipPath id="${clipId}"><polygon points="${starPoints}" /></clipPath>
            <clipPath id="${innerClipId}"><polygon points="${starPoints}" transform="${innerStarTransform}" /></clipPath>
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
            <clipPath id="${innerClipId}"><rect x="8.3" y="8.3" width="7.4" height="7.4" rx="1.0" /></clipPath>
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
            <clipPath id="${clipId}"><circle cx="12" cy="12" r="8.2" /></clipPath>
            <clipPath id="${innerClipId}"><circle cx="12" cy="12" r="4.5" /></clipPath>
          </defs>
          ${slices}
          ${radialHollowSlices(segmentList.map((segment) => ({ ...segment, kind: "literature" })), innerClipId, innerWhite)}
          <circle cx="12" cy="12" r="4.5" fill="none" stroke="${stroke}" stroke-width="0.45" />
          <circle cx="12" cy="12" r="8.2" fill="none" stroke="${stroke}" stroke-width="1.9" />
        </svg>`;
    }

    if (kind === "uncertain") {
      return `
        <svg viewBox="0 0 24 24" role="img" aria-hidden="true">
          <defs>
            <clipPath id="${clipId}">
              <g transform="rotate(45 12 12)"><rect x="9.9" y="3.0" width="4.2" height="18" rx="1.6" /></g>
              <g transform="rotate(-45 12 12)"><rect x="9.9" y="3.0" width="4.2" height="18" rx="1.6" /></g>
            </clipPath>
          </defs>
          <line x1="5.2" y1="5.2" x2="18.8" y2="18.8" stroke="${stroke}" stroke-width="6.0" stroke-linecap="round" />
          <line x1="18.8" y1="5.2" x2="5.2" y2="18.8" stroke="${stroke}" stroke-width="6.0" stroke-linecap="round" />
          ${slices}
        </svg>`;
    }

    if (kind === "unknown") {
      return `
        <svg viewBox="0 0 24 24" role="img" aria-hidden="true">
          <defs>
            <clipPath id="${clipId}"><polygon points="12,3.8 21,19.2 3,19.2" /></clipPath>
          </defs>
          ${slices}
          <polygon points="12,3.8 21,19.2 3,19.2" fill="none" stroke="${stroke}" stroke-width="1.9" stroke-linejoin="round" />
        </svg>`;
    }

    return `
      <svg viewBox="0 0 24 24" role="img" aria-hidden="true">
        <defs>
          <clipPath id="${clipId}"><circle cx="12" cy="12" r="8.2" /></clipPath>
          <clipPath id="${innerClipId}"><circle cx="12" cy="12" r="4.5" /></clipPath>
        </defs>
        ${slices}
        ${hollowSlices}
        <circle cx="12" cy="12" r="4.5" fill="none" stroke="rgba(0,0,0,0.22)" stroke-width="0.45" />
        <circle cx="12" cy="12" r="8.2" fill="none" stroke="${stroke}" stroke-width="1.9" />
      </svg>`;
  }

  function isHollowMarkerSegment(segment) {
    if (segment && typeof segment === "object") return Boolean(segment.hollow);
    return segment === "literature";
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
      if (!isHollowMarkerSegment(segment)) {
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

  function popupMarkerSvg(record) {
    const kind = recordKind(record?.recordType);
    const color = activeColorField() ? colorForRecord(record) : DEFAULT_MARKER_COLOR;
    const relatedRecords = recordsForPopupMarker(record);
    const hasThisStudyForSameColorValue = relatedRecords.some((candidate) => recordKind(candidate.recordType) === "thisStudy");
    const hollow = shouldUseHollowMarkerInterior(kind, hasThisStudyForSameColorValue);
    return markerSvg(kind, [{
      color,
      kind,
      hollow,
      representative: record
    }]);
  }

  function recordsForPopupMarker(record) {
    const colorField = activeColorField();
    if (!colorField) {
      return [record];
    }

    const colorValue = getRecordColorValue(record);
    return state.currentPopupRecords.filter((candidate) => getRecordColorValue(candidate) === colorValue);
  }

  function createPopupContent(record) {
    const content = document.createElement("div");
    content.className = "popup-content";

    const titleField = activePopupTitleField();
    if (titleField) {
      const rawTitleValue = titleField === COLUMNS.date
        ? (record.date || joinDateParts(record))
        : record?.__values?.[titleField];
      const titleValue = displayText(rawTitleValue);
      const title = document.createElement("div");
      title.className = "popup-title";
      const titleMarker = document.createElement("span");
      titleMarker.className = "popup-title-marker";
      titleMarker.innerHTML = popupMarkerSvg(record);
      title.appendChild(titleMarker);

      const titleText = document.createElement("span");
      titleText.innerHTML = isTaxonField(titleField) && normalizeText(rawTitleValue)
        ? formatTaxonHTML(titleValue)
        : escapeHtml(titleValue);

      title.appendChild(titleText);
      content.appendChild(title);
    }

    configuredPopupFields(getActiveDataset(), state.headers).forEach((field) => {
      if (field === VIRTUAL_COORDINATES_FIELD) {
        appendCoordinateRow(content, record);
        return;
      }
      const value = field === COLUMNS.date
        ? (record.date || joinDateParts(record))
        : record.__values?.[field];
      appendPopupRow(content, fieldLabel(field), value);
    });

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
    labelEl.textContent = "緯度, 経度";

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

    const popupLngLat = [
      Number(record.__popupLongitude ?? record.longitude),
      Number(record.__popupLatitude ?? record.latitude)
    ];
    const markerPixel = state.map.project(popupLngLat);
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
      .setLngLat(popupLngLat)
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

  function renderLegends(records) {
    renderRecordTypeLegend(records);
    renderTaxonLegend(records);
  }

  function renderRecordTypeLegend(records) {
    if (!recordTypeLegendEl) return;

    const visibleKinds = new Set(records.map((record) => recordKind(record.recordType)));
    const items = RECORD_TYPE_LEGEND_ITEMS.filter((item) => visibleKinds.has(item.kind));

    if (!items.length) {
      recordTypeLegendEl.innerHTML = '<div class="legend-row legend-empty">表示中の記録はありません</div>';
      return;
    }

    recordTypeLegendEl.innerHTML = items.map((item) => `
      <div class="legend-row">
        <span class="legend-symbol ${legendSymbolClassForKind(item.kind)}"></span>${escapeHTML(item.label)}
      </div>`).join("");
  }

  function legendSymbolClassForKind(kind) {
    if (kind === "type") return "star";
    if (kind === "synonymType") return "square";
    if (kind === "literature") return "hollow";
    if (kind === "thisStudy") return "filled";
    if (kind === "uncertain") return "x";
    return "triangle unknown";
  }

  function renderTaxonLegend(records) {
    if (!taxonLegendEl) return;

    const colorField = activeColorField();
    if (colorLegendTitleEl) {
      colorLegendTitleEl.textContent = colorField ? `色分け：${fieldLabel(colorField)}` : "色分け";
    }

    if (!colorField) {
      taxonLegendEl.innerHTML = '<div class="taxon-row legend-empty">色分け項目が未選択です</div>';
      return;
    }

    const values = [...new Set(records.map((record) => getRecordColorValue(record)))]
      .sort((a, b) => a.localeCompare(b, "ja"));

    if (!values.length) {
      taxonLegendEl.innerHTML = '<div class="taxon-row legend-empty">表示中の記録はありません</div>';
      return;
    }

    taxonLegendEl.innerHTML = values.map((value) => {
      const labelHtml = colorField === COLUMNS.taxon || colorField === LEGACY_TAXON_HEADER
        ? formatTaxonHTML(value)
        : escapeHTML(value);
      return `
        <div class="taxon-row">
          <span class="taxon-chip" style="background: ${colorForColorValue(value)}"></span>
          <span>${labelHtml}</span>
        </div>`;
    }).join("");
  }

  function setupPopupClose() {
    // ポップアップがマーカーより前面にある場合でも、
    // ポップアップに遮られて直接クリックできなかったマーカーは推定して開かない。
    // マーカー要素を直接クリックできた場合のみ、マーカー側の click ハンドラで開く。
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

  if (datasetLoadFieldsButton) datasetLoadFieldsButton.addEventListener("click", loadDatasetFieldsForModal);
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
    if (event.key === "Escape" && shareModal && !shareModal.hidden) {
      closeShareModal();
    }
    if (event.key === "Escape" && helpModal && !helpModal.hidden) {
      closeHelpModal();
    }
    if (event.key === "Escape" && registerModal && !registerModal.hidden) {
      closeRegisterModal();
    }
    if (event.key === "Escape" && filterModal && !filterModal.hidden) {
      closeFilterModal();
    }
  });

  if (datasetSelect) {
    datasetSelect.addEventListener("change", () => {
      state.activeDatasetId = datasetSelect.value;
      if (state.datasets.some((dataset) => dataset.id === state.activeDatasetId)) {
        saveActiveDatasetId(state.activeDatasetId);
      } else {
        saveActiveDatasetId("");
      }
      renderDatasetControls();
      resetMapDataForDatasetChange();
      reloadData();
    });
  }

  if (datasetAddButton) datasetAddButton.addEventListener("click", addDataset);
  if (datasetEditButton) datasetEditButton.addEventListener("click", editDataset);
  if (datasetDeleteButton) datasetDeleteButton.addEventListener("click", deleteDataset);
  if (datasetShareButton) datasetShareButton.addEventListener("click", openShareModal);
  if (datasetRegisterButton) datasetRegisterButton.addEventListener("click", openRegisterModal);
  if (registerModalCloseButton) registerModalCloseButton.addEventListener("click", closeRegisterModal);
  if (registerModalCancelButton) registerModalCancelButton.addEventListener("click", closeRegisterModal);
  if (registerConfirmButton) registerConfirmButton.addEventListener("click", registerActiveDataset);
  if (registerModal) {
    registerModal.addEventListener("click", (event) => {
      if (event.target === registerModal) closeRegisterModal();
    });
  }
  if (filterButton) filterButton.addEventListener("click", openFilterModal);
  if (filterModalCloseButton) filterModalCloseButton.addEventListener("click", closeFilterModal);
  if (filterModalCancelButton) filterModalCancelButton.addEventListener("click", closeFilterModal);
  if (filterApplyButton) filterApplyButton.addEventListener("click", applyFilterModal);
  if (filterSelectAllButton) filterSelectAllButton.addEventListener("click", () => setAllFilterCheckboxes(true));
  if (filterSelectNoneButton) filterSelectNoneButton.addEventListener("click", () => setAllFilterCheckboxes(false));
  if (filterResetButton) filterResetButton.addEventListener("click", resetFilterModalCheckboxes);
  if (filterModal) {
    filterModal.addEventListener("click", (event) => {
      if (event.target === filterModal) closeFilterModal();
    });
  }
  if (reloadCancelButton) reloadCancelButton.addEventListener("click", () => cancelActiveLoad(true));
  if (helpButton) helpButton.addEventListener("click", openHelpModal);
  if (helpModalCloseButton) helpModalCloseButton.addEventListener("click", closeHelpModal);
  if (helpModalCancelButton) helpModalCancelButton.addEventListener("click", closeHelpModal);
  if (helpModal) {
    helpModal.addEventListener("click", (event) => {
      if (event.target === helpModal) closeHelpModal();
    });
  }
  if (shareModalCloseButton) shareModalCloseButton.addEventListener("click", closeShareModal);
  if (shareModalCancelButton) shareModalCancelButton.addEventListener("click", closeShareModal);
  if (shareCopyButton) shareCopyButton.addEventListener("click", copyShareUrl);
  if (shareModal) {
    shareModal.addEventListener("click", (event) => {
      if (event.target === shareModal) closeShareModal();
    });
  }

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
