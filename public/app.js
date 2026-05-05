"use strict";

const PAPER_SIZES = {
  A4: { width: 210, height: 297 },
  Letter: { width: 216, height: 279 }
};

const MM_TO_PX = 3.78;
const MAX_PREVIEW_SCALE = 0.56;
const KATEX_CSS_URL = "https://cdn.jsdelivr.net/npm/katex@0.16.10/dist/katex.min.css";
const KATEX_JS_URL = "https://cdn.jsdelivr.net/npm/katex@0.16.10/dist/katex.min.js";
const MERMAID_JS_URL = "https://cdn.jsdelivr.net/npm/mermaid@10.9.3/dist/mermaid.min.js";
const TOKEN_STORAGE_KEY = "notionPdfToken";

const state = {
  token: localStorage.getItem(TOKEN_STORAGE_KEY) || sessionStorage.getItem(TOKEN_STORAGE_KEY) || "",
  source: null,
  schema: {},
  pages: [],
  selected: new Set(),
  bundles: [],
  filters: [],
  sorts: [],
  dbSettings: {
    hiddenProperties: []
  },
  previewSettingsCollapsed: false,
  filtersCollapsed: false,
  tablePage: 1,
  tablePageSize: 10,
  previewVersion: 0,
  savedDatabases: [],
  exportHistory: [],
  settings: {
    paper: "A4",
    orientation: "portrait",
    marginTop: 18,
    marginRight: 16,
    marginBottom: 18,
    marginLeft: 16,
    scale: 96,
    showHeader: true,
    showFooter: true,
    startEachPageOnNewSheet: true,
    showProperties: true,
    exportMode: "single"
  },
  loading: false
};

const els = {};
let enhancementRun = 0;
let mermaidReady = false;
let renderPreviewFrame = 0;
let renderTableFrame = 0;
let resizeFrame = 0;
let katexPromise = null;
let mermaidPromise = null;
let katexCssPromise = null;
let katexRefreshQueued = false;
let tokenSaveTimer = 0;
const previewCache = {
  key: "",
  previewHtml: "",
  printHtml: "",
  sheetCount: 0
};

document.addEventListener("DOMContentLoaded", init);

async function init() {
  bindElements();
  bindEvents();
  applySettingsToControls();
  await loadState();
  els.tokenInput.value = state.token;
  renderSidebarToggleState();
  renderFiltersCollapsedState();
  renderFilters();
  renderSortControls();
  renderDbSettings();
  renderPagesTable();
  renderPreview();
  updateConnection();
}

function bindElements() {
  for (const id of [
    "connectionLabel",
    "collapseSidebarButton",
    "tokenInput",
    "sourceInput",
    "syncButton",
    "savedDatabases",
    "saveDbButton",
    "quickPreviewGuideButton",
    "sourceMode",
    "searchInput",
    "selectAllButton",
    "dbSettingsButton",
    "loadPreviewButton",
    "exportWordButton",
    "exportGoogleDocsButton",
    "exportButton",
    "filterSummary",
    "toggleFiltersButton",
    "applyQueryButton",
    "addFilterButton",
    "clearFiltersButton",
    "filterBody",
    "filterRows",
    "sortPropertyInput",
    "sortDirectionInput",
    "dbSettingsPanel",
    "dbSettingsSummary",
    "propertySettingsList",
    "closeDbSettingsButton",
    "bulkCheckbox",
    "clearSelectionButton",
    "invertSelectionButton",
    "selectionCount",
    "pageCount",
    "tableHead",
    "pageRows",
    "tablePager",
    "pageSizeInput",
    "sheetCount",
    "togglePreviewSettingsButton",
    "previewSettingsPanel",
    "paperInput",
    "orientationInput",
    "exportModeInput",
    "scaleInput",
    "marginTopInput",
    "marginRightInput",
    "marginBottomInput",
    "marginLeftInput",
    "headerToggle",
    "footerToggle",
    "propertiesToggle",
    "newSheetToggle",
    "status",
    "previewCanvas",
    "readyExportSummary",
    "readyExportMode",
    "readyExportFormatInput",
    "readyExportButton",
    "printDocument",
    "printStyle"
  ]) {
    els[id] = document.getElementById(id);
  }
}

function bindEvents() {
  els.tokenInput.addEventListener("input", () => {
    state.token = els.tokenInput.value.trim();
    rememberTokenLocally(state.token);
    queueSaveToken();
    updateConnection();
  });
  els.tokenInput.addEventListener("change", () => saveTokenNow());
  els.tokenInput.addEventListener("blur", () => saveTokenNow());
  window.addEventListener("beforeunload", saveTokenBeforeClose);
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "hidden") saveTokenBeforeClose();
  });
  els.collapseSidebarButton.addEventListener("click", () => {
    document.body.classList.toggle("sidebar-collapsed");
    renderSidebarToggleState();
    scheduleRenderPreview(true);
  });
  els.syncButton.addEventListener("click", syncNotion);
  els.saveDbButton.addEventListener("click", saveCurrentDb);
  els.quickPreviewGuideButton.addEventListener("click", loadPreview);
  els.searchInput.addEventListener("input", () => {
    state.tablePage = 1;
    schedulePagesTable();
  });
  els.selectAllButton.addEventListener("click", toggleSelectAll);
  els.bulkCheckbox.addEventListener("change", () => toggleCurrentPageSelection(els.bulkCheckbox.checked));
  els.clearSelectionButton.addEventListener("click", clearSelection);
  els.invertSelectionButton.addEventListener("click", invertVisibleSelection);
  els.pageSizeInput.addEventListener("input", () => {
    state.tablePageSize = Number(els.pageSizeInput.value) || 10;
    state.tablePage = 1;
    renderPagesTable();
  });
  els.dbSettingsButton.addEventListener("click", () => {
    els.dbSettingsPanel.classList.toggle("hidden");
  });
  els.closeDbSettingsButton.addEventListener("click", () => {
    els.dbSettingsPanel.classList.add("hidden");
  });
  els.loadPreviewButton.addEventListener("click", loadPreview);
  els.exportWordButton.addEventListener("click", () => exportWord({ googleDocs: false }));
  els.exportGoogleDocsButton.addEventListener("click", () => exportWord({ googleDocs: true }));
  els.exportButton.addEventListener("click", exportPdf);
  els.readyExportButton.addEventListener("click", exportFromReadyCard);
  els.applyQueryButton.addEventListener("click", syncNotion);
  els.addFilterButton.addEventListener("click", () => {
    state.filters.push(defaultFilter());
    renderFilters();
  });
  els.clearFiltersButton.addEventListener("click", () => {
    state.filters = [];
    state.tablePage = 1;
    renderFilters();
    syncNotion();
  });
  els.toggleFiltersButton.addEventListener("click", () => {
    state.filtersCollapsed = !state.filtersCollapsed;
    renderFiltersCollapsedState();
  });
  window.addEventListener("resize", () => {
    cancelAnimationFrame(resizeFrame);
    resizeFrame = requestAnimationFrame(() => {
      if (state.bundles.length) scheduleRenderPreview(true);
    });
  });
  els.sortPropertyInput.addEventListener("input", () => {
    setSortFromControls();
  });
  els.sortDirectionInput.addEventListener("input", () => {
    setSortFromControls();
  });
  els.togglePreviewSettingsButton.addEventListener("click", () => {
    state.previewSettingsCollapsed = !state.previewSettingsCollapsed;
    renderPreviewSettingsState();
  });

  const settingBindings = [
    ["paperInput", "paper", "value"],
    ["orientationInput", "orientation", "value"],
    ["exportModeInput", "exportMode", "value"],
    ["scaleInput", "scale", "number"],
    ["marginTopInput", "marginTop", "number"],
    ["marginRightInput", "marginRight", "number"],
    ["marginBottomInput", "marginBottom", "number"],
    ["marginLeftInput", "marginLeft", "number"],
    ["headerToggle", "showHeader", "checked"],
    ["footerToggle", "showFooter", "checked"],
    ["propertiesToggle", "showProperties", "checked"],
    ["newSheetToggle", "startEachPageOnNewSheet", "checked"]
  ];

  for (const [elementId, key, mode] of settingBindings) {
    els[elementId].addEventListener("input", () => {
      state.settings[key] = mode === "number" ? Number(els[elementId].value) : els[elementId][mode];
      scheduleRenderPreview(true);
    });
  }
}

async function loadState() {
  try {
    const payload = await api("/api/state");
    if (payload.state.notionToken || !state.token) {
      state.token = payload.state.notionToken || state.token;
      rememberTokenLocally(state.token);
    }
    state.savedDatabases = payload.state.savedDatabases || [];
    state.savedDatabases = state.savedDatabases.filter((saved) => saved.sourceId !== "demo-notion-workspace" && saved.mode !== "demo");
    state.exportHistory = payload.state.exportHistory || [];
    state.settings = { ...state.settings, ...(payload.defaults || {}) };
    applySettingsToControls();
    renderPreviewSettingsState();
    renderSavedDatabases();
    renderHistory();
  } catch (error) {
    setStatus(error.message, true);
  }
}

function rememberTokenLocally(token) {
  if (token) {
    localStorage.setItem(TOKEN_STORAGE_KEY, token);
    sessionStorage.removeItem(TOKEN_STORAGE_KEY);
  } else {
    localStorage.removeItem(TOKEN_STORAGE_KEY);
    sessionStorage.removeItem(TOKEN_STORAGE_KEY);
  }
}

function queueSaveToken() {
  clearTimeout(tokenSaveTimer);
  tokenSaveTimer = setTimeout(() => {
    saveTokenNow();
  }, 350);
}

async function saveTokenNow() {
  clearTimeout(tokenSaveTimer);
  try {
    await api("/api/state/token", { token: state.token });
  } catch {
    // The token is still kept in localStorage; sync/preview will surface API issues.
  }
}

function saveTokenBeforeClose() {
  clearTimeout(tokenSaveTimer);
  try {
    const body = JSON.stringify({ token: state.token });
    const blob = new Blob([body], { type: "application/json" });
    navigator.sendBeacon("/api/state/token", blob);
  } catch {
    rememberTokenLocally(state.token);
  }
}

function markPreviewDirty() {
  state.previewVersion += 1;
  previewCache.key = "";
  if (els.previewCanvas) els.previewCanvas.dataset.previewKey = "";
}

function scheduleRenderPreview(force = false) {
  cancelAnimationFrame(renderPreviewFrame);
  renderPreviewFrame = requestAnimationFrame(() => renderPreview(force));
}

function schedulePagesTable() {
  cancelAnimationFrame(renderTableFrame);
  renderTableFrame = requestAnimationFrame(renderPagesTable);
}

async function syncNotion() {
  const sourceId = els.sourceInput.value.trim();
  if (!sourceId) {
    setStatus("Pega primero el ID o URL de una base de datos/data source de Notion.", true);
    return;
  }
  if (!state.token) {
    setStatus("Pega primero un token de integración de Notion.", true);
    return;
  }

  setLoading(true, "Sincronizando páginas de Notion...");
  try {
    await saveTokenNow();
    const payload = await api("/api/notion/query", {
      token: state.token,
      sourceId,
      filters: state.filters,
      sorts: state.sorts,
      maxPages: 300
    });
    state.source = payload.source;
    state.schema = payload.source.schema || {};
    state.pages = payload.pages || [];
    state.selected = new Set(state.pages.map((page) => page.id));
    state.tablePage = 1;
    state.bundles = [];
    markPreviewDirty();
    state.dbSettings = normalizeDbSettings(state.dbSettings);
    updateConnection();
    renderFilters();
    renderSortControls();
    renderDbSettings();
    renderPagesTable();
    renderPreview(true);
    setStatus(`Sincronizadas ${state.pages.length} página${state.pages.length === 1 ? "" : "s"}.`);
  } catch (error) {
    setStatus(error.message, true);
  } finally {
    setLoading(false);
  }
}

async function saveCurrentDb() {
  if (!state.source) {
    setStatus("Sincroniza o carga una DB antes de guardarla.", true);
    return;
  }
  setLoading(true, "Guardando preset de DB...");
  try {
    const payload = await api("/api/state/save-db", {
      token: state.token,
      sourceId: state.source.sourceId,
      name: state.source.name,
      filters: state.filters,
      sorts: state.sorts,
      dbSettings: state.dbSettings,
      settings: state.settings,
      lastSyncedAt: new Date().toISOString()
    });
    state.savedDatabases = payload.state.savedDatabases || [];
    renderSavedDatabases();
    setStatus("Preset de DB guardado localmente.");
  } catch (error) {
    setStatus(error.message, true);
  } finally {
    setLoading(false);
  }
}

async function loadPreview() {
  const pageIds = Array.from(state.selected);
  if (!pageIds.length) {
    state.bundles = [];
    markPreviewDirty();
    renderPreview();
    setStatus("Selecciona al menos una página para la vista previa.", true);
    return;
  }
  setLoading(true, "Descargando bloques de las páginas seleccionadas...");
  try {
    const payload = await api("/api/notion/pages", {
      token: state.token,
      pageIds
    });
    state.bundles = payload.pages || [];
    markPreviewDirty();
    renderPreview();
    setStatus(`Vista previa cargada para ${state.bundles.length} página${state.bundles.length === 1 ? "" : "s"}.`);
  } catch (error) {
    setStatus(error.message, true);
  } finally {
    setLoading(false);
  }
}

async function exportPdf() {
  if (!state.bundles.length) {
    await loadPreview();
  }
  if (!state.bundles.length) return;
  await ensurePreviewLibraries();
  renderPreview(true);
  await renderEnhancements();
  if (state.settings.exportMode === "separate" && state.bundles.length > 1) {
    printSeparateBundles(0);
    return;
  }
  setStatus("Abriendo impresión. Elige 'Guardar como PDF' en el navegador.");
  setTimeout(() => window.print(), 120);
}

async function exportWord({ googleDocs = false } = {}) {
  if (!state.bundles.length) {
    await loadPreview();
  }
  if (!state.bundles.length) return;

  setLoading(true, googleDocs ? "Preparando documento Word para Google Docs..." : "Generando documento Word...");
  try {
    await ensurePreviewLibraries();
    renderPreview(true);
    await renderEnhancements();
    const mermaidAssets = await collectMermaidAssets();
    const response = await downloadFromApi("/api/export/docx", {
      bundles: state.bundles,
      settings: state.settings,
      hiddenProperties: state.dbSettings.hiddenProperties || [],
      mermaidAssets,
      sourceName: state.source?.name || state.bundles[0]?.title || "Notion export"
    });
    downloadBlob(response.blob, response.filename);
    if (googleDocs) {
      setStatus("DOCX generado. Súbelo a Google Drive y ábrelo con Google Docs.");
      window.open("https://drive.google.com/drive/my-drive", "_blank", "noopener,noreferrer");
    } else {
      setStatus(`Documento Word generado: ${response.filename}`);
    }
  } catch (error) {
    setStatus(error.message, true);
  } finally {
    setLoading(false);
  }
}

function printSeparateBundles(index) {
  if (index >= state.bundles.length) {
    renderPreview(true);
    setStatus("Exportación separada terminada.");
    return;
  }
  const originalBundles = state.bundles;
  state.bundles = [originalBundles[index]];
  renderPreview(true);
  setStatus(`Imprimiendo PDF separado ${index + 1} de ${originalBundles.length}. Guarda el archivo y cierra el diálogo para continuar.`);
  renderEnhancements().finally(() => {
    window.print();
    state.bundles = originalBundles;
    setTimeout(() => printSeparateBundles(index + 1), 250);
  });
}

function renderSavedDatabases() {
  els.savedDatabases.innerHTML = "";
  const savedDatabases = state.savedDatabases.filter((saved) => saved.sourceId !== "demo-notion-workspace" && saved.mode !== "demo");
  if (!savedDatabases.length) {
    els.savedDatabases.innerHTML = `<div class="muted">No hay DB guardadas.</div>`;
    return;
  }
  for (const saved of savedDatabases) {
    const shell = document.createElement("div");
    shell.className = `saved-item-shell ${state.source?.sourceId === saved.sourceId ? "active" : ""}`;
    const button = document.createElement("button");
    button.className = `saved-item ${state.source?.sourceId === saved.sourceId ? "active" : ""}`;
    button.type = "button";
    button.innerHTML = `
      <span class="saved-db-icon" aria-hidden="true"></span>
      <span class="saved-copy">
        <strong>${escapeHtml(saved.name || "DB sin título")}</strong>
        <span>${escapeHtml(saved.sourceId || "")}</span>
        <small>Última sincronización: ${saved.lastSyncedAt ? formatDate(saved.lastSyncedAt) : "sin registro"}</small>
      </span>
    `;
    button.addEventListener("click", () => {
      els.sourceInput.value = saved.sourceId;
      state.filters = saved.filters || [];
      state.sorts = saved.sorts || [];
      state.dbSettings = normalizeDbSettings(saved.dbSettings);
      state.settings = { ...state.settings, ...(saved.settings || {}) };
      applySettingsToControls();
      renderFilters();
      renderSortControls();
      renderDbSettings();
      syncNotion();
    });
    const settingsButton = document.createElement("button");
    settingsButton.className = "saved-settings icon-button subtle-icon";
    settingsButton.type = "button";
    settingsButton.title = "Ajustes de propiedades";
    settingsButton.setAttribute("aria-label", `Ajustes de propiedades de ${saved.name || "DB guardada"}`);
    settingsButton.innerHTML = `<span aria-hidden="true">⚙</span>`;
    settingsButton.addEventListener("click", (event) => {
      event.stopPropagation();
      els.sourceInput.value = saved.sourceId;
      state.dbSettings = normalizeDbSettings(saved.dbSettings);
      renderDbSettings();
      els.dbSettingsPanel.classList.remove("hidden");
    });
    shell.append(button, settingsButton);
    els.savedDatabases.appendChild(shell);
  }
}

function renderHistory() {
  if (!els.historyList) return;
  els.historyList.innerHTML = "";
  if (!state.exportHistory.length) {
    els.historyList.innerHTML = `<div class="muted">Sin exportaciones.</div>`;
    return;
  }
  for (const item of state.exportHistory.slice(0, 6)) {
    const div = document.createElement("div");
    div.className = "history-item";
    div.innerHTML = `<strong>${escapeHtml(item.sourceName || "Exportación")}</strong><span>${item.pageCount} página${item.pageCount === 1 ? "" : "s"} · ${formatDate(item.at)}</span>`;
    els.historyList.appendChild(div);
  }
}

function renderFilters() {
  els.filterRows.innerHTML = "";
  const properties = Object.entries(state.schema || {}).filter(([, prop]) => supportedFilterTypes().has(prop.type));
  if (!state.filters.length) {
    els.filterRows.innerHTML = `<div class="muted">Añade filtros por propiedad y pulsa Aplicar.</div>`;
  }
  state.filters.forEach((filter, index) => {
    const row = document.createElement("div");
    row.className = "filter-row";

    const propertySelect = document.createElement("select");
    propertySelect.innerHTML = properties
      .map(([name, prop]) => `<option value="${escapeAttr(name)}">${escapeHtml(name)} · ${escapeHtml(prop.type)}</option>`)
      .join("");
    propertySelect.value = filter.property || properties[0]?.[0] || "";
    propertySelect.addEventListener("change", () => {
      state.filters[index].property = propertySelect.value;
      state.filters[index].operator = defaultOperator(propertyType(propertySelect.value));
      renderFilters();
    });

    const operatorSelect = document.createElement("select");
    const type = propertyType(propertySelect.value);
    operatorSelect.innerHTML = operatorsForType(type)
      .map((op) => `<option value="${op.value}">${escapeHtml(op.label)}</option>`)
      .join("");
    operatorSelect.value = filter.operator || defaultOperator(type);
    operatorSelect.addEventListener("change", () => {
      state.filters[index].operator = operatorSelect.value;
      renderFilters();
    });

    const valueInput = inputForFilter(type, filter.value, propertySelect.value, operatorSelect.value);
    const updateFilterValue = () => {
      state.filters[index].value = valueInput.type === "checkbox" ? valueInput.checked : valueInput.value;
    };
    valueInput.addEventListener("input", updateFilterValue);
    valueInput.addEventListener("change", updateFilterValue);

    const remove = document.createElement("button");
    remove.className = "icon-button";
    remove.type = "button";
    remove.title = "Quitar filtro";
    remove.textContent = "×";
    remove.addEventListener("click", () => {
      state.filters.splice(index, 1);
      renderFilters();
    });

    state.filters[index] = {
      property: propertySelect.value,
      operator: operatorSelect.value,
      value: valueInput.type === "checkbox" ? valueInput.checked : valueInput.value
    };

    row.append(propertySelect, operatorSelect, valueInput, remove);
    els.filterRows.appendChild(row);
  });

  els.filterSummary.textContent = state.filters.length
    ? `${state.filters.length} filtro${state.filters.length === 1 ? "" : "s"} activo${state.filters.length === 1 ? "" : "s"}`
    : "Sin filtros activos";
}

function renderSortControls() {
  const sortable = sortableProperties();
  els.sortPropertyInput.innerHTML = [
    `<option value="">Sin orden</option>`,
    ...sortable.map((item) => `<option value="${escapeAttr(item.value)}">${escapeHtml(item.label)}</option>`)
  ].join("");
  const active = state.sorts[0] || {};
  els.sortPropertyInput.value = active.property || "";
  els.sortDirectionInput.value = active.direction || "ascending";
}

function sortableProperties() {
  const schemaOptions = Object.entries(state.schema || {}).map(([name, prop]) => ({
    value: name,
    label: `${name} · ${prop.type}`
  }));
  return [
    { value: "__last_edited_time", label: "Editada · timestamp" },
    { value: "__created_time", label: "Creada · timestamp" },
    ...schemaOptions
  ];
}

function setSortFromControls() {
  const property = els.sortPropertyInput.value;
  state.sorts = property ? [{ property, direction: els.sortDirectionInput.value || "ascending" }] : [];
  state.tablePage = 1;
  renderPagesTable();
}

function renderDbSettings() {
  const names = Object.keys(state.schema || {}).filter((name) => state.schema[name].type !== "title");
  const hidden = new Set(state.dbSettings.hiddenProperties || []);
  els.propertySettingsList.innerHTML = names.length
    ? names.map((name) => {
        const prop = state.schema[name];
        return `
          <label class="property-toggle">
            <input type="checkbox" data-property-name="${escapeAttr(name)}" ${hidden.has(name) ? "" : "checked"} />
            <span>
              <strong>${escapeHtml(name)}</strong>
              <small>${escapeHtml(prop.type)}</small>
            </span>
          </label>
        `;
      }).join("")
    : `<div class="muted">Todavía no hay propiedades cargadas.</div>`;

  els.propertySettingsList.querySelectorAll("input[data-property-name]").forEach((box) => {
    box.addEventListener("change", () => {
      const name = box.dataset.propertyName;
      const nextHidden = new Set(state.dbSettings.hiddenProperties || []);
      if (box.checked) nextHidden.delete(name);
      else nextHidden.add(name);
      state.dbSettings.hiddenProperties = Array.from(nextHidden);
      renderDbSettingsSummary();
      renderPagesTable();
      scheduleRenderPreview(true);
    });
  });
  renderDbSettingsSummary();
}

function renderDbSettingsSummary() {
  const total = Object.keys(state.schema || {}).filter((name) => state.schema[name].type !== "title").length;
  const hidden = state.dbSettings.hiddenProperties?.length || 0;
  els.dbSettingsSummary.textContent = `${Math.max(0, total - hidden)} visibles · ${hidden} ocultas`;
}

function normalizeDbSettings(settings = {}) {
  return {
    hiddenProperties: Array.isArray(settings.hiddenProperties) ? settings.hiddenProperties : []
  };
}

function isPropertyHidden(name) {
  return Boolean(state.dbSettings.hiddenProperties?.includes(name));
}

function inputForFilter(type, value) {
  const operator = arguments[3];
  if (operator === "is_empty" || operator === "is_not_empty") {
    const input = document.createElement("input");
    input.type = "text";
    input.value = "";
    input.placeholder = "No hace falta valor";
    input.disabled = true;
    return input;
  }

  const propertyName = arguments[2];
  if (type === "select" || type === "status" || type === "multi_select") {
    const select = document.createElement("select");
    const options = propertyOptions(propertyName);
    select.innerHTML = [
      `<option value="">Elige opción</option>`,
      ...options.map((option) => `<option value="${escapeAttr(option.value)}">${escapeHtml(option.label)}</option>`)
    ].join("");
    select.value = value || "";
    return select;
  }

  if (type === "relation" || type === "people") {
    const select = document.createElement("select");
    const options = relationOrPeopleOptions(propertyName);
    select.innerHTML = [
      `<option value="">Elige ${type === "relation" ? "relación" : "persona"}</option>`,
      ...options.map((option) => `<option value="${escapeAttr(option.value)}">${escapeHtml(option.label)}</option>`)
    ].join("");
    select.value = value || "";
    return select;
  }

  const input = document.createElement("input");
  if (type === "checkbox") {
    const select = document.createElement("select");
    select.innerHTML = `<option value="true">Marcado</option><option value="false">Sin marcar</option>`;
    select.value = String(value === false || value === "false" ? false : true);
    return select;
  }
  if (type === "date" || type === "created_time" || type === "last_edited_time") {
    input.type = "date";
    input.value = value || "";
    return input;
  }
  if (type === "number") {
    input.type = "number";
    input.value = value ?? "";
    return input;
  }
  input.type = "text";
  input.placeholder = "Valor";
  input.value = value ?? "";
  return input;
}

function propertyOptions(name) {
  const prop = state.schema?.[name] || {};
  const source = prop[prop.type]?.options || [];
  const fromSchema = source.map((option) => ({ value: option.name, label: option.name }));
  if (fromSchema.length) return fromSchema;
  return uniqueOptionsFromPages(name).map((label) => ({ value: label, label }));
}

function relationOrPeopleOptions(name) {
  const seen = new Map();
  for (const page of state.pages) {
    const value = page.properties?.[name];
    if (value?.kind === "relation") {
      for (const item of value.items || []) {
        if (item.id) seen.set(item.id, item.title || item.id);
      }
    } else if (Array.isArray(value)) {
      for (const item of value) {
        if (item && typeof item === "object") {
          const id = item.id || item.name || item.title;
          if (id) seen.set(id, item.name || item.title || id);
        } else if (item) {
          seen.set(String(item), String(item));
        }
      }
    }
  }
  return Array.from(seen, ([value, label]) => ({ value, label }));
}

function uniqueOptionsFromPages(name) {
  const values = new Set();
  for (const page of state.pages) {
    const value = page.properties?.[name];
    if (Array.isArray(value)) value.forEach((item) => values.add(propertyText(item)));
    else if (value != null && value !== "") values.add(propertyText(value));
  }
  return Array.from(values).filter(Boolean).sort((a, b) => a.localeCompare(b));
}

function renderPagesTable() {
  const allVisible = filteredVisiblePages();
  const pageSize = Math.max(1, Number(state.tablePageSize || 10));
  const totalPages = Math.max(1, Math.ceil(allVisible.length / pageSize));
  state.tablePage = Math.min(Math.max(1, Number(state.tablePage || 1)), totalPages);
  const start = (state.tablePage - 1) * pageSize;
  const visible = allVisible.slice(start, start + pageSize);
  const propNames = preferredPropertyColumns();
  els.tableHead.innerHTML = [
    `<th style="width:34px"><input id="headCheckbox" type="checkbox" ${visible.length && visible.every((page) => state.selected.has(page.id)) ? "checked" : ""}></th>`,
    `<th style="width:34%">Nombre</th>`,
    ...propNames.map((name) => `<th>${escapeHtml(name)}</th>`),
    `<th style="width:128px">Editada</th>`
  ].join("");

  const checkbox = document.getElementById("headCheckbox");
  checkbox?.addEventListener("change", () => {
    toggleCurrentPageSelection(checkbox.checked);
  });

  els.pageRows.innerHTML = "";
  if (!allVisible.length) {
    els.pageRows.innerHTML = `<tr><td colspan="${propNames.length + 3}" class="muted">No hay páginas que coincidan con la búsqueda.</td></tr>`;
  }
  for (const page of visible) {
    const tr = document.createElement("tr");
    tr.innerHTML = [
      `<td><input type="checkbox" data-page-id="${escapeAttr(page.id)}" ${state.selected.has(page.id) ? "checked" : ""}></td>`,
      `<td><div class="row-title"><span class="page-icon">${pageIcon(page)}</span><strong>${escapeHtml(page.title)}</strong></div></td>`,
      ...propNames.map((name) => `<td>${propertyCell(page.properties?.[name])}</td>`),
      `<td>${formatDate(page.last_edited_time)}</td>`
    ].join("");
    els.pageRows.appendChild(tr);
  }

  els.pageRows.querySelectorAll("input[type='checkbox'][data-page-id]").forEach((box) => {
    box.addEventListener("change", () => {
      if (box.checked) state.selected.add(box.dataset.pageId);
      else state.selected.delete(box.dataset.pageId);
      renderSelectionMeta();
    });
  });
  renderTablePager(allVisible.length, pageSize);
  renderSelectionMeta();
}

function filteredVisiblePages() {
  const query = els.searchInput.value.trim().toLowerCase();
  const pages = !query ? state.pages : state.pages.filter((page) => {
    const haystack = [
      page.title,
      ...Object.values(page.properties || {}).map(propertyText)
    ].join(" ").toLowerCase();
    return haystack.includes(query);
  });
  return applyClientSort(pages);
}

function applyClientSort(pages) {
  const sort = state.sorts[0];
  if (!sort?.property) return pages;
  const direction = sort.direction === "ascending" ? 1 : -1;
  return [...pages].sort((a, b) => {
    const left = sort.property === "__last_edited_time"
      ? a.last_edited_time
      : sort.property === "__created_time"
        ? a.created_time
        : propertyText(a.properties?.[sort.property]);
    const right = sort.property === "__last_edited_time"
      ? b.last_edited_time
      : sort.property === "__created_time"
        ? b.created_time
        : propertyText(b.properties?.[sort.property]);
    return String(left || "").localeCompare(String(right || ""), undefined, { numeric: true, sensitivity: "base" }) * direction;
  });
}

function preferredPropertyColumns() {
  const names = Object.keys(state.schema || {}).filter((name) => state.schema[name].type !== "title" && !isPropertyHidden(name));
  const preferred = ["Status", "Estado", "Tags", "Etiquetas", "Owner", "Responsable", "Updated", "Actualizada", "Approved", "Aprobada"].filter((name) => names.includes(name));
  return [...preferred, ...names.filter((name) => !preferred.includes(name))].slice(0, 5);
}

function renderSelectionMeta() {
  const visible = filteredVisiblePages();
  const pageSize = Math.max(1, Number(state.tablePageSize || 10));
  const totalPages = Math.max(1, Math.ceil(visible.length / pageSize));
  state.tablePage = Math.min(Math.max(1, Number(state.tablePage || 1)), totalPages);
  const start = visible.length ? (state.tablePage - 1) * pageSize + 1 : 0;
  const end = Math.min(visible.length, state.tablePage * pageSize);
  const currentPageRows = visible.slice((state.tablePage - 1) * pageSize, (state.tablePage - 1) * pageSize + pageSize);
  els.selectionCount.textContent = `${state.selected.size} seleccionada${state.selected.size === 1 ? "" : "s"}`;
  els.pageCount.textContent = visible.length
    ? `${start}-${end} de ${visible.length} página${visible.length === 1 ? "" : "s"}`
    : `0 de ${state.pages.length} página${state.pages.length === 1 ? "" : "s"}`;
  els.bulkCheckbox.checked = Boolean(currentPageRows.length && currentPageRows.every((page) => state.selected.has(page.id)));
  els.bulkCheckbox.indeterminate = Boolean(currentPageRows.some((page) => state.selected.has(page.id)) && !els.bulkCheckbox.checked);
  updateReadyExportCard();
}

function toggleSelectAll() {
  const visible = filteredVisiblePages();
  const allSelected = visible.length && visible.every((page) => state.selected.has(page.id));
  if (allSelected) visible.forEach((page) => state.selected.delete(page.id));
  else visible.forEach((page) => state.selected.add(page.id));
  renderPagesTable();
}

function toggleCurrentPageSelection(checked) {
  const visible = filteredVisiblePages();
  const pageSize = Math.max(1, Number(state.tablePageSize || 10));
  const current = visible.slice((state.tablePage - 1) * pageSize, state.tablePage * pageSize);
  if (checked) current.forEach((page) => state.selected.add(page.id));
  else current.forEach((page) => state.selected.delete(page.id));
  renderPagesTable();
}

function clearSelection() {
  state.selected.clear();
  renderPagesTable();
}

function invertVisibleSelection() {
  for (const page of filteredVisiblePages()) {
    if (state.selected.has(page.id)) state.selected.delete(page.id);
    else state.selected.add(page.id);
  }
  renderPagesTable();
}

function renderTablePager(totalRows, pageSize) {
  const totalPages = Math.max(1, Math.ceil(totalRows / pageSize));
  const buttons = [
    `<button class="pager-button" type="button" data-page-action="prev" ${state.tablePage <= 1 ? "disabled" : ""}>‹</button>`
  ];
  const windowStart = Math.max(1, Math.min(state.tablePage - 1, totalPages - 2));
  const windowEnd = Math.min(totalPages, windowStart + 2);
  for (let page = windowStart; page <= windowEnd; page += 1) {
    buttons.push(`<button class="pager-button ${page === state.tablePage ? "active" : ""}" type="button" data-page="${page}">${page}</button>`);
  }
  buttons.push(`<button class="pager-button" type="button" data-page-action="next" ${state.tablePage >= totalPages ? "disabled" : ""}>›</button>`);
  els.tablePager.innerHTML = buttons.join("");
  els.tablePager.querySelectorAll("button[data-page], button[data-page-action]").forEach((button) => {
    button.addEventListener("click", () => {
      if (button.dataset.page) state.tablePage = Number(button.dataset.page);
      else if (button.dataset.pageAction === "prev") state.tablePage -= 1;
      else if (button.dataset.pageAction === "next") state.tablePage += 1;
      renderPagesTable();
    });
  });
}

async function exportFromReadyCard() {
  const format = els.readyExportFormatInput.value;
  if (format === "word") return exportWord({ googleDocs: false });
  if (format === "google") return exportWord({ googleDocs: true });
  return exportPdf();
}

function updateReadyExportCard() {
  const selectedCount = state.selected.size;
  const sheetCount = previewCache.sheetCount || 0;
  els.readyExportSummary.textContent = `${selectedCount} página${selectedCount === 1 ? "" : "s"} seleccionada${selectedCount === 1 ? "" : "s"}`;
  els.readyExportMode.textContent = sheetCount
    ? `${sheetCount} hoja${sheetCount === 1 ? "" : "s"} en vista previa. Se exportará con los ajustes actuales.`
    : "Carga la vista previa para confirmar saltos y márgenes antes de exportar.";
  els.readyExportButton.disabled = state.loading || !selectedCount;
}

function renderPreview(force = false) {
  const key = previewCacheKey();
  if (!force && previewCache.key === key && els.previewCanvas.dataset.previewKey === key) {
    updatePrintStyle();
    return;
  }

  if (force || previewCache.key !== key) {
    const sheets = paginateBundles(state.bundles);
    previewCache.key = key;
    previewCache.sheetCount = sheets.length;
    previewCache.previewHtml = sheets.length
      ? sheets.map((sheet, index) => renderSheet(sheet, index, sheets.length)).join("")
      : `<div class="preview-empty"><div><strong>Sin vista previa todavía.</strong><br>Selecciona páginas y pulsa Vista previa.</div></div>`;
    previewCache.printHtml = sheets.map((sheet, index) => renderSheet(sheet, index, sheets.length, true)).join("");
  }

  els.previewCanvas.innerHTML = previewCache.previewHtml;
  els.printDocument.innerHTML = previewCache.printHtml;
  els.previewCanvas.dataset.previewKey = key;
  els.sheetCount.textContent = `${previewCache.sheetCount} hoja${previewCache.sheetCount === 1 ? "" : "s"}`;
  updateReadyExportCard();
  updatePrintStyle();
  if (hasMathContent() && !window.katex) requestKatexRefresh();
  queueEnhancements();
}

function queueEnhancements() {
  const run = ++enhancementRun;
  requestAnimationFrame(() => {
    renderEnhancements(run);
  });
}

async function renderEnhancements(run = ++enhancementRun) {
  if (run !== enhancementRun) return;
  await renderMermaidDiagrams(els.previewCanvas, "preview");
  if (run !== enhancementRun) return;
  await renderMermaidDiagrams(els.printDocument, "print");
}

function previewCacheKey() {
  const settings = state.settings;
  return JSON.stringify({
    version: state.previewVersion,
    bundleIds: state.bundles.map((bundle) => bundle.id),
    hiddenProperties: state.dbSettings.hiddenProperties || [],
    previewWidth: Math.round(els.previewCanvas?.clientWidth || 0),
    paper: settings.paper,
    orientation: settings.orientation,
    scale: settings.scale,
    marginTop: settings.marginTop,
    marginRight: settings.marginRight,
    marginBottom: settings.marginBottom,
    marginLeft: settings.marginLeft,
    showHeader: settings.showHeader,
    showFooter: settings.showFooter,
    showProperties: settings.showProperties,
    startEachPageOnNewSheet: settings.startEachPageOnNewSheet
  });
}

function paginateBundles(bundles) {
  const settings = state.settings;
  const dimensions = paginationDimensions(settings);
  const measureHost = createMeasureHost(dimensions.bodyWidth, settings.scale);
  const pages = [];
  let current = null;
  const usableBodyHeight = dimensions.bodyHeight - dimensions.safetyGap;

  try {
    for (const bundle of bundles) {
      if (!current || settings.startEachPageOnNewSheet) current = newSheet(bundle.title);
      const items = buildBundleItems(bundle);

      for (let index = 0; index < items.length; index += 1) {
        const item = items[index];
        const measured = measurePreviewItem(measureHost, item);
        const nextMeasured = items[index + 1] ? measurePreviewItem(measureHost, items[index + 1]) : null;
        const shouldKeepWithNext = isSectionStart(item) && nextMeasured;
        const wouldOverflow = current.used + measured.height > usableBodyHeight;
        const orphanHeading = shouldKeepWithNext && current.used + measured.height + Math.min(nextMeasured.height, 110) > usableBodyHeight;

        if (current.items.length && (wouldOverflow || orphanHeading)) {
          pages.push(current);
          current = newSheet(bundle.title);
        }
        current.items.push(item);
        current.used += measured.height;
      }

      if (settings.startEachPageOnNewSheet && current.items.length) {
        pages.push(current);
        current = null;
      }
    }
  } finally {
    measureHost.remove();
  }

  if (current?.items.length) pages.push(current);
  return pages;
}

function buildBundleItems(bundle) {
  const items = [{ kind: "title", bundle }];
  if (state.settings.showProperties) items.push({ kind: "properties", bundle });
  for (const block of flattenBlocks(bundle.blocks || [])) {
    items.push({ kind: "block", block, bundle });
  }
  return items;
}

function paginationDimensions(settings) {
  const paper = paperSize(settings);
  const header = settings.showHeader ? 26 : 0;
  const footer = settings.showFooter ? 26 : 0;
  const marginWidth = (paper.width - settings.marginLeft - settings.marginRight) * MM_TO_PX;
  const marginHeight = (paper.height - settings.marginTop - settings.marginBottom) * MM_TO_PX;
  return {
    bodyWidth: Math.max(240, marginWidth),
    bodyHeight: Math.max(140, marginHeight - header - footer - 12),
    safetyGap: 18
  };
}

function createMeasureHost(width, scale) {
  const host = document.createElement("div");
  host.className = "measure-host sheet-body";
  host.style.position = "absolute";
  host.style.left = "-10000px";
  host.style.top = "0";
  host.style.visibility = "hidden";
  host.style.pointerEvents = "none";
  host.style.width = `${width}px`;
  host.style.zoom = Math.max(0.3, scale / 100);
  document.body.appendChild(host);
  return host;
}

function measurePreviewItem(host, item) {
  if (item.kind === "block" && isMermaidBlock(item.block)) {
    const lines = blockPlainText(item.block).split(/\r?\n/).filter(Boolean).length;
    return { height: Math.min(520, Math.max(220, 130 + lines * 24)) };
  }
  host.innerHTML = renderPreviewItem(item);
  const child = host.firstElementChild;
  const height = child ? child.getBoundingClientRect().height : 0;
  host.innerHTML = "";
  return { height: Math.ceil(height + 8) };
}

function isSectionStart(item) {
  return item.kind === "title" || ["heading_1", "heading_2", "heading_3", "heading_4"].includes(item.block?.type);
}

function newSheet(title) {
  return {
    title,
    items: [],
    used: 0
  };
}

function flattenBlocks(blocks, depth = 0) {
  const out = [];
  let numberedIndex = 0;
  for (const block of blocks) {
    if (block.type === "numbered_list_item") numberedIndex += 1;
    else numberedIndex = 0;
    const splitBlocks = block.type === "table" ? splitTableBlock(block) : splitLongTextBlock(block);
    splitBlocks.forEach((part, partIndex) => {
      out.push({
        ...part,
        depth,
        listIndex: block.type === "numbered_list_item" ? numberedIndex : undefined,
        continued: partIndex > 0
      });
    });
    if (block.children?.length && block.type !== "table") {
      out.push(...flattenBlocks(block.children, depth + 1));
    }
  }
  return out;
}

function splitTableBlock(block) {
  const rows = block.children || [];
  if (rows.length <= 6) return [block];
  const hasHeader = Boolean(block.table?.has_column_header);
  const header = hasHeader ? rows[0] : null;
  const bodyRows = hasHeader ? rows.slice(1) : rows;
  const chunks = [];
  for (let i = 0; i < bodyRows.length; i += 5) {
    const chunkRows = bodyRows.slice(i, i + 5);
    chunks.push({
      ...block,
      id: `${block.id || "table"}-part-${chunks.length}`,
      children: header ? [header, ...chunkRows] : chunkRows
    });
  }
  return chunks;
}

function splitLongTextBlock(block) {
  const textTypes = new Set(["paragraph", "bulleted_list_item", "numbered_list_item", "quote", "callout"]);
  if (!textTypes.has(block.type)) return [block];
  const payload = block[block.type];
  const text = blockPlainText(block);
  if (!payload?.rich_text || text.length < 680) return [block];
  const chunks = chunkText(text, 520);
  return chunks.map((chunk, index) => ({
    ...block,
    id: `${block.id || block.type}-part-${index}`,
    [block.type]: {
      ...payload,
      rich_text: [{
        type: "text",
        text: { content: chunk, link: null },
        annotations: { bold: false, italic: false, strikethrough: false, underline: false, code: false, color: "default" },
        plain_text: chunk,
        href: null
      }]
    }
  }));
}

function chunkText(text, targetLength) {
  const sentences = String(text).split(/(?<=[.!?])\s+/);
  const chunks = [];
  let current = "";
  for (const sentence of sentences) {
    if ((current + " " + sentence).trim().length > targetLength && current) {
      chunks.push(current);
      current = sentence;
    } else {
      current = (current ? `${current} ` : "") + sentence;
    }
  }
  if (current) chunks.push(current);
  return chunks.flatMap((chunk) => {
    if (chunk.length <= targetLength * 1.35) return [chunk];
    const pieces = [];
    for (let i = 0; i < chunk.length; i += targetLength) pieces.push(chunk.slice(i, i + targetLength));
    return pieces;
  });
}

function estimateBlockHeight(block) {
  const text = blockPlainText(block);
  const lineCount = Math.max(1, Math.ceil(text.length / 58));
  const base = {
    heading_1: 54,
    heading_2: 44,
    heading_3: 36,
    heading_4: 32,
    paragraph: 28,
    bulleted_list_item: 30,
    numbered_list_item: 30,
    to_do: 30,
    quote: 50,
    callout: 70,
    divider: 24,
    code: 30 + Math.max(1, text.split(/\r?\n/).length) * 20,
    table: 44 + (block.children?.length || 1) * 38,
    table_row: 32,
    image: 220,
    video: 92,
    audio: 64,
    pdf: 92,
    bookmark: 72,
    embed: 72,
    equation: 42,
    toggle: 34,
    column_list: 20,
    column: 20,
    synced_block: 24,
    table_of_contents: 40,
    breadcrumb: 28,
    file: 56,
    unsupported: 56
  }[block.type] || 28;
  return base + Math.max(0, lineCount - 1) * 22;
}

function estimateTitleHeight(title) {
  return 68 + Math.max(0, Math.ceil(String(title || "").length / 32) - 1) * 36;
}

function estimatePropertiesHeight(properties) {
  const entries = Object.entries(properties).filter(([name, value]) => name !== "Name" && !isPropertyHidden(name) && value !== "" && value != null);
  return entries.length ? 34 + Math.min(entries.length, 10) * 23 : 0;
}

function renderSheet(sheet, index, total, print = false) {
  const settings = state.settings;
  const paper = paperSize(settings);
  const scale = print ? 1 : previewScaleForPaper(paper);
  const unit = print ? "mm" : "px";
  const unitScale = print ? 1 : MM_TO_PX;
  const style = [
    `--sheet-w:${paper.width * unitScale}${unit}`,
    `--sheet-h:${paper.height * unitScale}${unit}`,
    `--margin-top:${settings.marginTop * unitScale}${unit}`,
    `--margin-right:${settings.marginRight * unitScale}${unit}`,
    `--margin-bottom:${settings.marginBottom * unitScale}${unit}`,
    `--margin-left:${settings.marginLeft * unitScale}${unit}`,
    `--doc-scale:${settings.scale / 100}`,
    `--preview-scale:${scale}`
  ].join(";");

  return `
    <article class="paper-sheet" style="${style}">
      <div class="paper-margin">
        ${settings.showHeader ? `<header class="sheet-header"><span>${escapeHtml(sheet.title)}</span><span>Notion PDF Studio</span></header>` : ""}
        <div class="sheet-body">
          ${sheet.items.map(renderPreviewItem).join("")}
        </div>
        ${settings.showFooter ? `<footer class="sheet-footer"><span>${formatDate(new Date().toISOString())}</span><span>${index + 1} / ${total}</span></footer>` : ""}
      </div>
    </article>
  `;
}

function previewScaleForPaper(paper) {
  const available = Math.max(260, (els.previewCanvas?.clientWidth || 520) - 36);
  const naturalWidth = paper.width * MM_TO_PX;
  return Math.min(MAX_PREVIEW_SCALE, Math.max(0.32, available / naturalWidth));
}

function renderPreviewItem(item) {
  if (item.kind === "title") {
    return `<h1 class="notion-page-title">${escapeHtml(item.bundle.title)}</h1>`;
  }
  if (item.kind === "properties") {
    return renderProperties(item.bundle.properties || {});
  }
  return renderBlock(item.block);
}

function renderProperties(properties) {
  const entries = Object.entries(properties).filter(([name, value]) => name !== "Name" && !isPropertyHidden(name) && value !== "" && value != null);
  if (!entries.length) return "";
  return `<dl class="notion-props">${entries
    .slice(0, 10)
    .map(([name, value]) => `<dt>${escapeHtml(name)}</dt><dd>${propertyCell(value)}</dd>`)
    .join("")}</dl>`;
}

function renderBlock(block) {
  const depthStyle = block.depth ? `style="margin-left:${Math.min(block.depth * 22, 88)}px"` : "";
  const type = block.type || "unsupported";
  if (type === "divider") return `<div class="block divider" ${depthStyle}></div>`;
  if (type === "table") return renderTable(block);
  if (["image", "video", "audio", "pdf", "file"].includes(type)) return renderMedia(block, depthStyle);
  if (type === "equation") return `<div class="block equation" ${depthStyle}>${renderEquation(block.equation?.expression || "", true)}</div>`;
  if (type === "code") {
    if (isMermaidBlock(block)) {
      const source = blockPlainText(block);
      return `<div class="block mermaid-diagram" ${depthStyle} data-mermaid-source="${escapeAttr(source)}"><pre>${escapeHtml(source)}</pre></div>`;
    }
    const text = richTextToHtml(block.code?.rich_text || []);
    return `<pre class="block code" ${depthStyle}>${text}</pre>`;
  }
  if (type === "quote") {
    return `<blockquote class="block quote" ${depthStyle}>${richTextToHtml(block.quote?.rich_text || [])}</blockquote>`;
  }
  if (type === "callout") {
    const icon = block.callout?.icon?.emoji || "!";
    return `<div class="block callout" ${depthStyle}><span>${escapeHtml(icon)}</span><div>${richTextToHtml(block.callout?.rich_text || [])}</div></div>`;
  }
  if (type === "to_do") {
    const checked = block.to_do?.checked;
    return `<div class="block to_do" ${depthStyle}><span class="todo-box ${checked ? "checked" : ""}">${checked ? "✓" : ""}</span><div>${richTextToHtml(block.to_do?.rich_text || [])}</div></div>`;
  }
  if (type === "bulleted_list_item") {
    return `<div class="block bulleted_list_item" ${depthStyle}><span class="list-marker">•</span><div>${richTextToHtml(block.bulleted_list_item?.rich_text || [])}</div></div>`;
  }
  if (type === "numbered_list_item") {
    return `<div class="block numbered_list_item" ${depthStyle}><span class="list-marker">${block.listIndex || 1}.</span><div>${richTextToHtml(block.numbered_list_item?.rich_text || [])}</div></div>`;
  }
  if (type === "toggle") {
    return `<div class="block toggle" ${depthStyle}><span class="toggle-caret">▾</span><div>${richTextToHtml(block.toggle?.rich_text || [])}</div></div>`;
  }
  if (type === "button") {
    const label = richTextToHtml(block.button?.rich_text || []) || escapeHtml(block.button?.name || "Botón");
    return `<div class="block notion-button" ${depthStyle}>${label}</div>`;
  }
  if (type === "column_list") return `<div class="block column-list" ${depthStyle}></div>`;
  if (type === "column") return `<div class="block column" ${depthStyle}></div>`;
  if (type === "synced_block") return `<div class="block synced-block" ${depthStyle}>${block.synced_block?.synced_from ? "Sincronizado desde otro bloque" : "Bloque sincronizado"}</div>`;
  if (type === "table_of_contents") return `<div class="block table-of-contents" ${depthStyle}>Índice</div>`;
  if (type === "breadcrumb") return `<div class="block breadcrumb" ${depthStyle}>Ruta de navegación</div>`;
  if (["heading_1", "heading_2", "heading_3", "heading_4"].includes(type)) {
    return `<div class="block ${type}" ${depthStyle}>${richTextToHtml(block[type]?.rich_text || [])}</div>`;
  }
  if (type === "paragraph") {
    const html = richTextToHtml(block.paragraph?.rich_text || []);
    return `<p class="block paragraph" ${depthStyle}>${html || "&nbsp;"}</p>`;
  }
  if (type === "bookmark" || type === "embed" || type === "link_preview") {
    const url = block[type]?.url || "";
    return `<div class="block unsupported" ${depthStyle}>${escapeHtml(type.replaceAll("_", " "))}: ${linkHtml(url)}</div>`;
  }
  if (type === "child_page" || type === "child_database") {
    return `<div class="block unsupported" ${depthStyle}>↳ ${escapeHtml(block[type]?.title || type.replaceAll("_", " "))}</div>`;
  }
  const unsupportedType = block.unsupported?.block_type || type;
  return `<div class="block unsupported" ${depthStyle}>Bloque de Notion no soportado: ${escapeHtml(unsupportedType)}</div>`;
}

function renderTable(block) {
  const rows = block.children || [];
  if (!rows.length) return `<div class="block unsupported">Tabla vacía</div>`;
  return `<div class="notion-table-block"><table><tbody>${rows
    .map((row, rowIndex) => `<tr>${(row.table_row?.cells || [])
      .map((cell) => rowIndex === 0 && block.table?.has_column_header
        ? `<th>${richTextToHtml(cell)}</th>`
        : `<td>${richTextToHtml(cell)}</td>`)
      .join("")}</tr>`)
    .join("")}</tbody></table></div>`;
}

function renderMedia(block, depthStyle) {
  const payload = block[block.type] || {};
  const file = payload.type === "external" ? payload.external : payload.file;
  const url = file?.url || "";
  if (block.type === "image" && url) {
    return `<img class="notion-image" ${depthStyle} src="${escapeAttr(url)}" alt="" loading="lazy">`;
  }
  return `<div class="media-placeholder" ${depthStyle}>${escapeHtml(block.type.toUpperCase())}: ${linkHtml(url || "Archivo de Notion")}</div>`;
}

function updatePrintStyle() {
  const settings = state.settings;
  const paper = settings.paper === "Letter" ? "Letter" : "A4";
  els.printStyle.textContent = `
    @page {
      size: ${paper} ${settings.orientation || "portrait"};
      margin: 0;
    }
  `;
}

function updateConnection() {
  const mode = state.source?.mode || "local";
  els.connectionLabel.textContent = state.token ? "Token cargado" : "Sin token";
  els.sourceMode.textContent = state.source ? `${state.source.name} · ${modeLabel(mode)}` : "Vista local";
  renderSavedDatabases();
}

function renderPreviewSettingsState() {
  els.previewSettingsPanel.classList.toggle("hidden", state.previewSettingsCollapsed);
  els.togglePreviewSettingsButton.classList.toggle("is-collapsed", state.previewSettingsCollapsed);
  els.togglePreviewSettingsButton.title = state.previewSettingsCollapsed ? "Mostrar ajustes de vista previa" : "Ocultar ajustes de vista previa";
  els.togglePreviewSettingsButton.setAttribute("aria-label", els.togglePreviewSettingsButton.title);
  els.togglePreviewSettingsButton.setAttribute("aria-expanded", String(!state.previewSettingsCollapsed));
}

function renderSidebarToggleState() {
  const collapsed = document.body.classList.contains("sidebar-collapsed");
  els.collapseSidebarButton.classList.toggle("is-collapsed", collapsed);
  els.collapseSidebarButton.title = collapsed ? "Mostrar panel lateral" : "Ocultar panel lateral";
  els.collapseSidebarButton.setAttribute("aria-label", els.collapseSidebarButton.title);
  els.collapseSidebarButton.setAttribute("aria-expanded", String(!collapsed));
}

function renderFiltersCollapsedState() {
  els.filterBody.classList.toggle("hidden", state.filtersCollapsed);
  els.toggleFiltersButton.classList.toggle("is-collapsed", state.filtersCollapsed);
  els.toggleFiltersButton.title = state.filtersCollapsed ? "Mostrar filtros" : "Ocultar filtros";
  els.toggleFiltersButton.setAttribute("aria-label", els.toggleFiltersButton.title);
  els.toggleFiltersButton.setAttribute("aria-expanded", String(!state.filtersCollapsed));
}

function applySettingsToControls() {
  els.paperInput.value = state.settings.paper;
  els.orientationInput.value = state.settings.orientation;
  els.exportModeInput.value = state.settings.exportMode || "single";
  els.scaleInput.value = state.settings.scale;
  els.marginTopInput.value = state.settings.marginTop;
  els.marginRightInput.value = state.settings.marginRight;
  els.marginBottomInput.value = state.settings.marginBottom;
  els.marginLeftInput.value = state.settings.marginLeft;
  els.headerToggle.checked = state.settings.showHeader;
  els.footerToggle.checked = state.settings.showFooter;
  els.propertiesToggle.checked = state.settings.showProperties;
  els.newSheetToggle.checked = state.settings.startEachPageOnNewSheet;
  els.pageSizeInput.value = String(state.tablePageSize || 10);
}

function modeLabel(mode) {
  return {
    local: "local",
    notion: "Notion",
    data_source: "origen de datos",
    database: "base de datos"
  }[mode] || mode;
}

function propertyType(name) {
  return state.schema?.[name]?.type || "rich_text";
}

function defaultFilter() {
  const first = Object.keys(state.schema || {}).find((name) => supportedFilterTypes().has(state.schema[name].type));
  const type = first ? propertyType(first) : "rich_text";
  return { property: first || "", operator: defaultOperator(type), value: "" };
}

function supportedFilterTypes() {
  return new Set([
    "title",
    "rich_text",
    "select",
    "status",
    "multi_select",
    "checkbox",
    "date",
    "number",
    "url",
    "email",
    "phone_number",
    "files",
    "people",
    "relation",
    "created_time",
    "last_edited_time"
  ]);
}

function operatorsForType(type) {
  if (type === "checkbox") return [{ value: "equals", label: "es" }];
  if (type === "files") {
    return [
      { value: "is_empty", label: "vacío" },
      { value: "is_not_empty", label: "no vacío" }
    ];
  }
  if (type === "relation" || type === "people") {
    return [
      { value: "contains", label: "contiene" },
      { value: "does_not_contain", label: "no contiene" },
      { value: "is_empty", label: "vacío" },
      { value: "is_not_empty", label: "no vacío" }
    ];
  }
  if (type === "number") {
    return [
      { value: "equals", label: "=" },
      { value: "greater_than", label: ">" },
      { value: "less_than", label: "<" },
      { value: "greater_than_or_equal_to", label: ">=" },
      { value: "less_than_or_equal_to", label: "<=" },
      { value: "is_empty", label: "vacío" },
      { value: "is_not_empty", label: "no vacío" }
    ];
  }
  if (type === "date" || type === "created_time" || type === "last_edited_time") {
    return [
      { value: "equals", label: "en" },
      { value: "before", label: "antes" },
      { value: "after", label: "después" },
      { value: "on_or_before", label: "en/antes" },
      { value: "on_or_after", label: "en/después" },
      { value: "is_empty", label: "vacío" },
      { value: "is_not_empty", label: "no vacío" }
    ];
  }
  if (type === "select" || type === "status") {
    return [
      { value: "equals", label: "es" },
      { value: "does_not_equal", label: "no es" },
      { value: "is_empty", label: "vacío" },
      { value: "is_not_empty", label: "no vacío" }
    ];
  }
  if (type === "multi_select") {
    return [
      { value: "contains", label: "contiene" },
      { value: "does_not_contain", label: "no contiene" },
      { value: "is_empty", label: "vacío" },
      { value: "is_not_empty", label: "no vacío" }
    ];
  }
  return [
    { value: "contains", label: "contiene" },
    { value: "does_not_contain", label: "no contiene" },
    { value: "equals", label: "es" },
    { value: "does_not_equal", label: "no es" },
    { value: "is_empty", label: "vacío" },
    { value: "is_not_empty", label: "no vacío" }
  ];
}

function defaultOperator(type) {
  return operatorsForType(type)[0]?.value || "contains";
}

function paperSize(settings) {
  const paper = PAPER_SIZES[settings.paper] || PAPER_SIZES.A4;
  if (settings.orientation === "landscape") {
    return { width: paper.height, height: paper.width };
  }
  return paper;
}

function pageIcon(page) {
  if (page.icon?.type === "emoji") return escapeHtml(page.icon.emoji);
  return "□";
}

function propertyCell(value) {
  if (value?.kind === "relation") {
    const items = value.items || [];
    if (!items.length) return `<span class="muted">Vacío</span>`;
    return `<span class="tag-list relation-tags">${items
      .map((item) => `<span class="tag relation-tag" title="${escapeAttr(item.id)}">${escapeHtml(item.title || item.id)}</span>`)
      .join("")}${value.hasMore ? `<span class="tag">+ más</span>` : ""}</span>`;
  }
  if (Array.isArray(value)) {
    if (!value.length) return `<span class="muted">Vacío</span>`;
    return `<span class="tag-list">${value.map((item) => {
      if (item && typeof item === "object") {
        const label = item.name || item.title || item.id || JSON.stringify(item);
        return item.url
          ? `<a class="tag" href="${escapeAttr(item.url)}">${escapeHtml(label)}</a>`
          : `<span class="tag">${escapeHtml(label)}</span>`;
      }
      return `<span class="tag">${escapeHtml(item)}</span>`;
    }).join("")}</span>`;
  }
  if (typeof value === "boolean") return value ? "Sí" : "No";
  if (value == null || value === "") return `<span class="muted">Vacío</span>`;
  return escapeHtml(String(value));
}

function propertyText(value) {
  if (value?.kind === "relation") return (value.items || []).map((item) => item.title || item.id).join(" ");
  if (Array.isArray(value)) return value.map(propertyText).join(" ");
  if (value && typeof value === "object") return value.name || value.title || value.id || Object.values(value).map(propertyText).join(" ");
  return String(value ?? "");
}

async function renderMermaidDiagrams(root, scope) {
  if (!root) return;
  const diagrams = Array.from(root.querySelectorAll(".mermaid-diagram:not([data-rendered='true'])"));
  if (!diagrams.length) return;
  try {
    await ensureMermaid();
  } catch {
    diagrams.forEach((node) => {
      const source = node.getAttribute("data-mermaid-source") || "";
      node.classList.add("mermaid-error");
      node.innerHTML = `<strong>No se pudo cargar Mermaid.</strong><pre>${escapeHtml(source)}</pre>`;
      node.dataset.rendered = "true";
    });
    return;
  }
  if (!window.mermaid) return;
  if (!mermaidReady) {
    window.mermaid.initialize({
      startOnLoad: false,
      securityLevel: "strict",
      theme: "base",
      themeVariables: {
        primaryColor: "#f6f5f1",
        primaryTextColor: "#25231f",
        primaryBorderColor: "#d8d1c6",
        lineColor: "#7b756c",
        fontFamily: "ui-sans-serif, -apple-system, BlinkMacSystemFont, Segoe UI, Helvetica, Arial, sans-serif",
        tertiaryColor: "#fbfaf8"
      }
    });
    mermaidReady = true;
  }

  for (let index = 0; index < diagrams.length; index += 1) {
    const node = diagrams[index];
    const source = node.getAttribute("data-mermaid-source") || "";
    if (!source.trim()) continue;
    try {
      const id = `mermaid-${scope}-${Date.now()}-${index}-${Math.random().toString(36).slice(2)}`;
      const result = await window.mermaid.render(id, source);
      node.innerHTML = result.svg;
      node.dataset.rendered = "true";
    } catch (error) {
      node.classList.add("mermaid-error");
      node.innerHTML = `<strong>No se pudo renderizar el diagrama Mermaid.</strong><pre>${escapeHtml(source)}</pre>`;
    }
  }
}

async function collectMermaidAssets() {
  const assets = {};
  const nodes = Array.from(els.previewCanvas.querySelectorAll(".mermaid-diagram[data-rendered='true']"));
  for (const node of nodes) {
    const source = node.getAttribute("data-mermaid-source") || "";
    const svg = node.querySelector("svg");
    if (!source || !svg || hasMermaidAsset(assets, source)) continue;
    try {
      rememberMermaidAsset(assets, source, await svgToPngAsset(svg));
    } catch {
      // The server will fall back to the Mermaid source code.
    }
  }
  for (const source of mermaidSourcesFromBundles()) {
    if (hasMermaidAsset(assets, source)) continue;
    try {
      rememberMermaidAsset(assets, source, await renderMermaidSourceToPng(source));
    } catch {
      // The server will fall back to the Mermaid source code.
    }
  }
  return assets;
}

function rememberMermaidAsset(assets, source, asset) {
  assets[source] = asset;
  assets[normalizeMermaidSource(source)] = asset;
}

function hasMermaidAsset(assets, source) {
  return Boolean(assets[source] || assets[normalizeMermaidSource(source)]);
}

function normalizeMermaidSource(source) {
  return String(source || "").replace(/\r\n/g, "\n").replace(/\s+$/g, "").trim();
}

function mermaidSourcesFromBundles() {
  return Array.from(new Set(state.bundles
    .flatMap((bundle) => flattenBlocks(bundle.blocks || []))
    .filter(isMermaidBlock)
    .map(blockPlainText)
    .filter((source) => source.trim())));
}

async function renderMermaidSourceToPng(source) {
  await ensureMermaid();
  if (!mermaidReady) {
    window.mermaid.initialize({
      startOnLoad: false,
      securityLevel: "strict",
      theme: "base",
      themeVariables: {
        primaryColor: "#f6f5f1",
        primaryTextColor: "#25231f",
        primaryBorderColor: "#d8d1c6",
        lineColor: "#7b756c",
        fontFamily: "ui-sans-serif, -apple-system, BlinkMacSystemFont, Segoe UI, Helvetica, Arial, sans-serif",
        tertiaryColor: "#fbfaf8"
      }
    });
    mermaidReady = true;
  }
  const id = `mermaid-word-${Date.now()}-${Math.random().toString(36).slice(2)}`;
  const result = await window.mermaid.render(id, source);
  const host = document.createElement("div");
  host.className = "mermaid-export-host";
  host.style.position = "fixed";
  host.style.left = "-12000px";
  host.style.top = "0";
  host.style.width = "900px";
  host.style.background = "#fff";
  host.innerHTML = result.svg;
  document.body.appendChild(host);
  try {
    const svg = host.querySelector("svg");
    if (!svg) throw new Error("Mermaid no devolvió SVG.");
    return await svgToPngAsset(svg);
  } finally {
    host.remove();
  }
}

async function svgToPngAsset(svg) {
  const clone = svg.cloneNode(true);
  clone.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  const box = svg.viewBox?.baseVal;
  const width = Math.ceil(Number(svg.getAttribute("width")) || box?.width || svg.getBoundingClientRect().width || 900);
  const height = Math.ceil(Number(svg.getAttribute("height")) || box?.height || svg.getBoundingClientRect().height || 520);
  clone.setAttribute("width", width);
  clone.setAttribute("height", height);

  const xml = new XMLSerializer().serializeToString(clone);
  const blob = new Blob([xml], { type: "image/svg+xml;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  try {
    const image = await loadImage(url);
    const canvas = document.createElement("canvas");
    const scale = Math.min(2, window.devicePixelRatio || 1.5);
    canvas.width = Math.max(1, Math.round(width * scale));
    canvas.height = Math.max(1, Math.round(height * scale));
    const context = canvas.getContext("2d");
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(image, 0, 0, canvas.width, canvas.height);
    return {
      data: canvas.toDataURL("image/png"),
      width,
      height
    };
  } finally {
    URL.revokeObjectURL(url);
  }
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = reject;
    image.src = src;
  });
}

async function ensurePreviewLibraries() {
  const tasks = [];
  if (hasMathContent()) tasks.push(ensureKatex());
  if (hasMermaidContent()) tasks.push(ensureMermaid());
  await Promise.allSettled(tasks);
}

function requestKatexRefresh() {
  if (katexRefreshQueued) return;
  katexRefreshQueued = true;
  ensureKatex()
    .then(() => renderPreview(true))
    .catch(() => {})
    .finally(() => {
      katexRefreshQueued = false;
    });
}

function ensureKatex() {
  if (window.katex) return Promise.resolve(window.katex);
  if (!katexPromise) {
    katexPromise = Promise.all([
      loadExternalStyle(KATEX_CSS_URL),
      loadExternalScript(KATEX_JS_URL)
    ]).then(() => window.katex);
  }
  return katexPromise;
}

function ensureMermaid() {
  if (window.mermaid) return Promise.resolve(window.mermaid);
  if (!mermaidPromise) {
    mermaidPromise = loadExternalScript(MERMAID_JS_URL).then(() => window.mermaid);
  }
  return mermaidPromise;
}

function loadExternalScript(src) {
  const existing = document.querySelector(`script[src="${escapeCssAttr(src)}"]`);
  if (existing) {
    return new Promise((resolve, reject) => {
      if (existing.dataset.loaded === "true") resolve();
      else {
        existing.addEventListener("load", resolve, { once: true });
        existing.addEventListener("error", reject, { once: true });
      }
    });
  }
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = src;
    script.async = true;
    script.addEventListener("load", () => {
      script.dataset.loaded = "true";
      resolve();
    }, { once: true });
    script.addEventListener("error", reject, { once: true });
    document.head.appendChild(script);
  });
}

function loadExternalStyle(href) {
  if (document.querySelector(`link[href="${escapeCssAttr(href)}"]`)) return Promise.resolve();
  if (!katexCssPromise) {
    katexCssPromise = new Promise((resolve, reject) => {
      const link = document.createElement("link");
      link.rel = "stylesheet";
      link.href = href;
      link.addEventListener("load", resolve, { once: true });
      link.addEventListener("error", reject, { once: true });
      document.head.appendChild(link);
    });
  }
  return katexCssPromise;
}

function hasMathContent() {
  return state.bundles.some((bundle) => flattenBlocks(bundle.blocks || []).some((block) => block.type === "equation" || blockHasInlineEquation(block)));
}

function hasMermaidContent() {
  return state.bundles.some((bundle) => flattenBlocks(bundle.blocks || []).some(isMermaidBlock));
}

function blockHasInlineEquation(block) {
  const payload = block?.[block.type] || {};
  return Array.isArray(payload.rich_text) && payload.rich_text.some((part) => part.type === "equation");
}

function renderEquation(expression, displayMode = false) {
  const source = String(expression || "");
  if (!source.trim()) return "";
  if (window.katex?.renderToString) {
    try {
      return window.katex.renderToString(source, {
        displayMode,
        throwOnError: false,
        strict: "ignore",
        output: "html"
      });
    } catch {
      return `<span class="math-fallback">${escapeHtml(source)}</span>`;
    }
  }
  return `<span class="math-fallback">${escapeHtml(source)}</span>`;
}

function isMermaidBlock(block) {
  return block?.type === "code" && String(block.code?.language || "").toLowerCase() === "mermaid";
}

function richTextToHtml(richText) {
  return (richText || []).map((part) => {
    let text = part.type === "equation"
      ? `<span class="rich-equation">${renderEquation(part.equation?.expression || part.plain_text || "", false)}</span>`
      : escapeHtml(part.plain_text || part.text?.content || part.mention?.plain_text || "");
    const annotations = part.annotations || {};
    if (annotations.code) text = `<code class="rich-code">${text}</code>`;
    if (annotations.bold) text = `<strong>${text}</strong>`;
    if (annotations.italic) text = `<em>${text}</em>`;
    if (annotations.underline) text = `<u>${text}</u>`;
    if (annotations.strikethrough) text = `<s>${text}</s>`;
    if (part.href) text = `<a href="${escapeAttr(part.href)}">${text}</a>`;
    return text;
  }).join("");
}

function blockPlainText(block) {
  const payload = block[block.type] || {};
  if (block.type === "table") {
    return (block.children || [])
      .flatMap((row) => row.table_row?.cells || [])
      .flat()
      .map((part) => part.plain_text || "")
      .join(" ");
  }
  return (payload.rich_text || [])
    .map((part) => part.plain_text || part.equation?.expression || "")
    .join("");
}

function linkHtml(url) {
  if (!url) return "";
  const safe = escapeHtml(url);
  return `<a href="${escapeAttr(url)}">${safe}</a>`;
}

function formatDate(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);
  return new Intl.DateTimeFormat(undefined, { month: "short", day: "numeric", year: "numeric" }).format(date);
}

function setLoading(loading, message) {
  state.loading = loading;
  for (const button of [els.syncButton, els.loadPreviewButton, els.exportButton, els.exportWordButton, els.exportGoogleDocsButton, els.saveDbButton, els.readyExportButton]) {
    button.disabled = loading;
  }
  updateReadyExportCard();
  if (message) setStatus(message);
}

function setStatus(message, error = false) {
  els.status.textContent = message;
  els.status.classList.toggle("error", error);
}

async function api(path, body) {
  const response = await fetch(path, {
    method: body ? "POST" : "GET",
    headers: body ? { "Content-Type": "application/json" } : {},
    body: body ? JSON.stringify(body) : undefined
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload.error || `La petición ha fallado con HTTP ${response.status}`);
  }
  return payload;
}

async function downloadFromApi(path, body) {
  const response = await fetch(path, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    throw new Error(payload.error || `La descarga ha fallado con HTTP ${response.status}`);
  }
  const blob = await response.blob();
  return {
    blob,
    filename: responseFileName(response.headers.get("Content-Disposition")) || "notion-export.docx"
  };
}

function responseFileName(disposition) {
  const value = String(disposition || "");
  const utf = value.match(/filename\*=UTF-8''([^;]+)/i);
  if (utf) return decodeURIComponent(utf[1]);
  const basic = value.match(/filename="?([^";]+)"?/i);
  return basic?.[1] || "";
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename || "notion-export.docx";
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeAttr(value) {
  return escapeHtml(value);
}

function escapeCssAttr(value) {
  return String(value ?? "").replaceAll("\\", "\\\\").replaceAll('"', '\\"');
}
