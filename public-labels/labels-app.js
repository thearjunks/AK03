function localDateValue(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

const historyToday = new Date();
const historyWeekStart = new Date(historyToday);
historyWeekStart.setDate(historyWeekStart.getDate() - 6);

const state = {
  rows: [],
  snapshots: [],
  page: 1,
  pageSize: 25,
  total: 0,
  historyRows: [],
  historyTotal: 0,
  historyFilters: {
    search: "",
    featureKey: "all",
    changeType: "all",
    source: "Strapi Content History",
    dateFrom: localDateValue(historyWeekStart),
    dateTo: localDateValue(historyToday),
  },
  adminSessionActive: false,
  language: {
    rows: [],
    page: 1,
    pageSize: 50,
    total: 0,
    filters: {
      search: "",
      issueType: "all",
      featureKey: "all",
      status: "all",
    },
  },
  dictionary: {
    jobId: "",
    rows: [],
    summary: {},
    filter: "all",
    page: 1,
    pageSize: 50,
  },
  filters: {
    search: "",
    featureKey: "all",
    status: "all",
    changeType: "all",
    languageIssue: "all",
  },
  advanced: {},
  currentTarget: "labels",
  ready: false,
};

const $ = (id) => document.getElementById(id);
const els = {
  lastFetch: $("lastFetch"),
  fetchButton: $("fetchButton"),
  refreshButton: $("refreshButton"),
  downloadExcel: $("downloadExcel"),
  downloadExcel2: $("downloadExcel2"),
  stats: {
    total: $("statTotal"),
    existing: $("statExisting"),
    new: $("statNew"),
    modified: $("statModified"),
    removed: $("statRemoved"),
    issues: $("statIssues"),
    featureKeys: $("statFeatureKeys"),
    existingPct: $("statExistingPct"),
    newPct: $("statNewPct"),
    modifiedPct: $("statModifiedPct"),
    removedPct: $("statRemovedPct"),
    issuesPct: $("statIssuesPct"),
  },
  snapshotList: $("snapshotList"),
  globalSearch: $("globalSearch"),
  featureKeySearch: $("featureKeySearch"),
  englishTextSearch: $("englishTextSearch"),
  arabicTextSearch: $("arabicTextSearch"),
  featureFilter: $("featureFilter"),
  statusFilter: $("statusFilter"),
  changeFilter: $("changeFilter"),
  languageIssueFilter: $("languageIssueFilter"),
  advancedToggle: $("advancedToggle"),
  advancedPanel: $("advancedPanel"),
  clearAdvanced: $("clearAdvanced"),
  activeFilterPill: $("activeFilterPill"),
  labelsBody: $("labelsBody"),
  labelsPageTitle: $("labelsPageTitle"),
  labelsPageDescription: $("labelsPageDescription"),
  labelsPageExport: $("labelsPageExport"),
  tableInfo: $("tableInfo"),
  pageSize: $("pageSize"),
  prevPage: $("prevPage"),
  nextPage: $("nextPage"),
  pageNumber: $("pageNumber"),
  drawerBackdrop: $("drawerBackdrop"),
  drawerTitle: $("drawerTitle"),
  drawerSubtitle: $("drawerSubtitle"),
  drawerContent: $("drawerContent"),
  closeDrawer: $("closeDrawer"),
  toast: $("toast"),
  dictionarySource: $("dictionarySource"),
  dictionaryFileInput: $("dictionaryFileInput"),
  dictionaryBrowseButton: $("dictionaryBrowseButton"),
  dictionaryDropZone: $("dictionaryDropZone"),
  dictionaryFileMeta: $("dictionaryFileMeta"),
  dictionaryProgress: $("dictionaryProgress"),
  dictionaryProgressLabel: $("dictionaryProgressLabel"),
  dictionaryProgressValue: $("dictionaryProgressValue"),
  dictionaryProgressBar: $("dictionaryProgressBar"),
  dictionaryResults: $("dictionaryResults"),
  dictionarySummary: $("dictionarySummary"),
  dictionaryFilters: $("dictionaryFilters"),
  dictionaryTableInfo: $("dictionaryTableInfo"),
  dictionaryPageInfo: $("dictionaryPageInfo"),
  dictionaryPageSize: $("dictionaryPageSize"),
  dictionaryPrevPage: $("dictionaryPrevPage"),
  dictionaryNextPage: $("dictionaryNextPage"),
  dictionaryPageNumber: $("dictionaryPageNumber"),
  dictionaryBody: $("dictionaryBody"),
  dictionaryDownloadUpdated: $("dictionaryDownloadUpdated"),
  dictionaryDownloadExceptions: $("dictionaryDownloadExceptions"),
  dictionaryDownloadSummary: $("dictionaryDownloadSummary"),
  snapshotContent: $("snapshotContent"),
  languageFreshness: $("languageFreshness"),
  languageSearch: $("languageSearch"),
  languageIssueTypeFilter: $("languageIssueTypeFilter"),
  languageFeatureFilter: $("languageFeatureFilter"),
  languageStatusFilter: $("languageStatusFilter"),
  languageSummary: $("languageSummary"),
  languageBody: $("languageBody"),
  languageTableInfo: $("languageTableInfo"),
  languagePageSize: $("languagePageSize"),
  languagePrevPage: $("languagePrevPage"),
  languageNextPage: $("languageNextPage"),
  languagePageNumber: $("languagePageNumber"),
  auditContent: $("auditContent"),
  historySearch: $("historySearch"),
  historyFeatureFilter: $("historyFeatureFilter"),
  historyChangeFilter: $("historyChangeFilter"),
  historySourceFilter: $("historySourceFilter"),
  historyDateFrom: $("historyDateFrom"),
  historyDateTo: $("historyDateTo"),
  historyConnectionStatus: $("historyConnectionStatus"),
  historyAdminEmail: $("historyAdminEmail"),
  historyAdminPassword: $("historyAdminPassword"),
  historyConnectButton: $("historyConnectButton"),
  historySyncButton: $("historySyncButton"),
  historyDisconnectButton: $("historyDisconnectButton"),
  apiContent: $("apiContent"),
};

els.historyDateFrom.value = state.historyFilters.dateFrom;
els.historyDateTo.value = state.historyFilters.dateTo;

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString();
}

function formatDate(value) {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleString(undefined, { year: "numeric", month: "short", day: "2-digit", hour: "2-digit", minute: "2-digit" });
}

function pct(value, total) {
  if (!total) return "0% of total";
  return `${((Number(value || 0) / total) * 100).toFixed(1)}% of total`;
}

function toast(message) {
  els.toast.textContent = message;
  els.toast.hidden = false;
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => { els.toast.hidden = true; }, 3200);
}

async function api(path, options) {
  const response = await fetch(path, options);
  const json = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(json.error || `Request failed: ${response.status}`);
  return json;
}

function queryParams(includePaging = true) {
  const params = new URLSearchParams();
  for (const [key, value] of Object.entries(state.filters)) {
    if (value && value !== "all") params.set(key, value);
  }
  for (const [key, value] of Object.entries(state.advanced)) {
    if (value) params.set(key, value);
  }
  if (includePaging) {
    params.set("page", state.page);
    params.set("pageSize", state.pageSize);
  }
  return params;
}

function historyQueryParams() {
  const params = new URLSearchParams();
  for (const [key, value] of Object.entries(state.historyFilters)) {
    if (value && value !== "all") params.set(key, value);
  }
  params.set("page", 1);
  params.set("pageSize", 100);
  return params;
}

function renderStats(totals = {}) {
  const total = totals.total || 0;
  els.stats.total.textContent = formatNumber(total);
  els.stats.existing.textContent = formatNumber(totals.existing);
  els.stats.new.textContent = formatNumber(totals.new);
  els.stats.modified.textContent = formatNumber(totals.modified);
  els.stats.removed.textContent = formatNumber(totals.removed);
  els.stats.issues.textContent = formatNumber(totals.languageIssues);
  els.stats.featureKeys.textContent = `Across ${formatNumber(totals.featureEntries ?? totals.featureKeys)} Feature Entries`;
  els.stats.existingPct.textContent = pct(totals.existing, total);
  els.stats.newPct.textContent = pct(totals.new, total);
  els.stats.modifiedPct.textContent = pct(totals.modified, total);
  els.stats.removedPct.textContent = pct(totals.removed, total);
  els.stats.issuesPct.textContent = pct(totals.languageIssues, total);
}

function fillSelect(select, values, selected) {
  const first = select.querySelector("option")?.outerHTML || "<option value=\"all\">All</option>";
  select.innerHTML = first + values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join("");
  select.value = selected || "all";
}

function renderOptions(options = {}) {
  fillSelect(els.featureFilter, options.featureKeys || [], state.filters.featureKey);
  fillSelect(els.statusFilter, options.statuses || [], state.filters.status);
  fillSelect(els.changeFilter, options.changes || [], state.filters.changeType);
  fillSelect(els.historyFeatureFilter, options.featureKeys || [], state.historyFilters.featureKey);
}

function renderSnapshots(snapshots = []) {
  state.snapshots = snapshots;
  if (!snapshots.length) {
    if (els.snapshotList) els.snapshotList.innerHTML = `<div class="empty-box">No snapshots yet</div>`;
    els.snapshotContent.innerHTML = `<div class="empty-box">No snapshots yet</div>`;
    return;
  }
  const html = snapshots.slice(0, 5).map((snapshot, index) => {
    const counts = snapshot.counts || {};
    const changeCounts = snapshot.changeCounts || counts;
    return `
      <article class="snapshot-item ${index === 0 ? "active" : ""}">
        <div><span class="radio-dot"></span><strong>${formatDate(snapshot.fetchedAt)}</strong>${index === 0 ? "<em>Latest</em>" : ""}</div>
        <p>${formatNumber(counts.existing || counts.total || 0)} Active Labels · ${formatNumber(counts.featureKeys || 65)} Feature Entries</p>
        <footer>
          <button data-snapshot="${escapeHtml(snapshot.id)}" data-change="new" type="button">New: ${formatNumber(changeCounts.new)}</button>
          <button data-snapshot="${escapeHtml(snapshot.id)}" data-change="modified" type="button">Modified: ${formatNumber(changeCounts.modified)}</button>
          <button data-snapshot="${escapeHtml(snapshot.id)}" data-change="removed" type="button">Removed: ${formatNumber(changeCounts.removed)}</button>
        </footer>
      </article>`;
  }).join("");
  if (els.snapshotList) els.snapshotList.innerHTML = html;
  els.snapshotContent.innerHTML = `<div class="snapshot-full">${snapshots.map((snapshot) => `
    <button class="snapshot-row" data-snapshot="${escapeHtml(snapshot.id)}" data-change="new" type="button">
      ${formatDate(snapshot.fetchedAt)} · New ${formatNumber((snapshot.changeCounts || snapshot.counts)?.new)} · Modified ${formatNumber((snapshot.changeCounts || snapshot.counts)?.modified)} · Removed ${formatNumber((snapshot.changeCounts || snapshot.counts)?.removed)}
    </button>`).join("")}</div>`;
}

function issueClass(issue) {
  if (issue === "None") return "none";
  if (issue.includes("Arabic has English")) return "danger";
  if (issue.includes("English has Arabic")) return "danger";
  if (issue.includes("Missing")) return "warning";
  return "warning";
}

function badgeClass(change) {
  return String(change || "").toLowerCase().replaceAll(" ", "-") || "existing";
}

function actorCell(row) {
  const name = row.changedBy || "-";
  const details = [row.changedByUsername, row.changedByEmail].filter(Boolean).join(" · ");
  return `<span class="actor-cell"><strong>${escapeHtml(name)}</strong>${details ? `<small>${escapeHtml(details)}</small>` : ""}</span>`;
}

function renderTable() {
  if (!state.rows.length) {
    els.labelsBody.innerHTML = `<tr><td colspan="12" class="empty-row">No labels match the current filters.</td></tr>`;
  } else {
    els.labelsBody.innerHTML = state.rows.map((row, index) => `
      <tr>
        <td><span class="id-text">${escapeHtml(row.labelId)}</span></td>
        <td>${escapeHtml(row.featureKey)}</td>
        <td>${escapeHtml(row.labelKey)}</td>
        <td class="text-cell">${escapeHtml(row.englishText)}</td>
        <td class="text-cell rtl">${escapeHtml(row.arabicText)}</td>
        <td><span class="issue ${issueClass(row.languageIssue)}">${escapeHtml(row.languageIssue)}</span></td>
        <td><span class="status">${escapeHtml(row.status || "-")}</span></td>
        <td>${formatDate(row.firstSeenAt)}</td>
        <td><span class="change ${badgeClass(row.changeType)}">${escapeHtml(row.changeType || "-")}</span></td>
        <td>${escapeHtml(row.actorRole || "No activity in last 7 days")}</td>
        <td>${actorCell(row)}</td>
        <td><button class="row-action" data-index="${index}" type="button">View</button></td>
      </tr>`).join("");
  }
  const start = state.total ? ((state.page - 1) * state.pageSize) + 1 : 0;
  const end = Math.min(state.page * state.pageSize, state.total);
  els.tableInfo.textContent = `Showing ${formatNumber(start)} to ${formatNumber(end)} of ${formatNumber(state.total)} labels`;
  els.pageNumber.textContent = state.page;
  els.prevPage.disabled = state.page <= 1;
  els.nextPage.disabled = state.page * state.pageSize >= state.total;
}

function renderSidePanels(totals = {}) {
  els.apiContent.innerHTML = `<code>/api/status</code><code>/api/labels</code><code>/api/language-consistency</code><code>/api/dictionary/process</code><code>/api/download/xlsx</code>`;
}

function languageIssueClass(issueType) {
  if (issueType.includes("missing")) return "warning";
  if (issueType.includes("contains") || issueType.includes("full English")) return "danger";
  if (issueType.includes("case")) return "case";
  return "mapping";
}

function languageQueryParams() {
  const params = new URLSearchParams();
  for (const [key, value] of Object.entries(state.language.filters)) {
    if (value && value !== "all") params.set(key, value);
  }
  params.set("page", state.language.page);
  params.set("pageSize", state.language.pageSize);
  return params;
}

function renderLanguageSummary(totals = {}) {
  const metrics = [
    ["Affected Labels", totals.affectedLabels],
    ["Issue Groups", totals.issueGroups],
    ["Issue Occurrences", totals.issueOccurrences],
    ["English Case", totals.englishCaseMismatchGroups],
    ["Translation Mismatch", totals.translationMismatchGroups],
    ["Duplicate Translation", totals.duplicateTranslationGroups],
    ["Label Key Mismatch", totals.labelKeyMismatchGroups],
    ["Field Language Issues", totals.directLanguageIssueGroups],
  ];
  els.languageSummary.innerHTML = metrics.map(([label, value]) => `
    <article><span>${escapeHtml(label)}</span><strong>${formatNumber(value)}</strong></article>`).join("");
}

function renderLanguageOptions(options = {}) {
  fillSelect(els.languageIssueTypeFilter, options.issueTypes || [], state.language.filters.issueType);
  fillSelect(els.languageFeatureFilter, options.featureKeys || [], state.language.filters.featureKey);
  fillSelect(els.languageStatusFilter, options.statuses || [], state.language.filters.status);
}

function renderLanguageTable(data = {}) {
  const rows = data.rows || [];
  state.language.rows = rows;
  state.language.total = data.total || 0;
  if (!rows.length) {
    els.languageBody.innerHTML = `<tr><td colspan="9" class="empty-row">No language inconsistencies match the current filters.</td></tr>`;
  } else {
    els.languageBody.innerHTML = rows.map((row, index) => {
      const moduleScreen = [row.module, row.screenName].filter(Boolean).join(" / ") || "-";
      return `
        <tr>
          <td><span class="consistency-issue ${languageIssueClass(row.issueType)}">${escapeHtml(row.issueType)}</span></td>
          <td class="text-cell">${escapeHtml(row.englishText || "-")}</td>
          <td class="text-cell rtl">${escapeHtml(row.arabicText || "-")}</td>
          <td><strong class="occurrence-count">${formatNumber(row.occurrenceCount)}</strong></td>
          <td>${escapeHtml(row.featureKey || "-")}</td>
          <td>${escapeHtml(row.labelKey || "-")}</td>
          <td>${escapeHtml(moduleScreen)}</td>
          <td><span class="status">${escapeHtml(row.status || "-")}</span></td>
          <td><button class="row-action language-row-action" data-index="${index}" type="button">View</button></td>
        </tr>`;
    }).join("");
  }
  const start = state.language.total ? ((state.language.page - 1) * state.language.pageSize) + 1 : 0;
  const end = Math.min(state.language.page * state.language.pageSize, state.language.total);
  els.languageTableInfo.textContent = `Showing ${formatNumber(start)} to ${formatNumber(end)} of ${formatNumber(state.language.total)} issue occurrences`;
  els.languagePageNumber.textContent = state.language.page;
  els.languagePrevPage.disabled = state.language.page <= 1;
  els.languageNextPage.disabled = state.language.page * state.language.pageSize >= state.language.total;
  els.languageFreshness.textContent = `Source: ${formatDate(data.lastSnapshot?.fetchedAt)}`;
}

async function loadLanguageConsistency() {
  const data = await api(`/api/language-consistency?${languageQueryParams()}`);
  renderLanguageSummary(data.totals || {});
  renderLanguageOptions(data.options || {});
  renderLanguageTable(data);
}

function formatFileSize(bytes) {
  const size = Number(bytes || 0);
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

function dictionaryFilteredRows() {
  const rows = state.dictionary.rows;
  const filter = state.dictionary.filter;
  if (filter === "matched") return rows.filter((row) => row.matchStatus === "Matched");
  if (filter === "not-found") return rows.filter((row) => row.matchStatus === "Not Found in Dictionary");
  if (filter === "translation-required") return rows.filter((row) => row.reviewStatus === "Translation Required");
  if (filter === "needs-review") return rows.filter((row) => row.reviewStatus === "Needs Review");
  if (filter === "updated") return rows.filter((row) => row.matchStatus === "Updated");
  return rows;
}

function dictionaryBadgeClass(value) {
  const normalized = String(value || "").toLowerCase();
  if (normalized.includes("not found") || normalized.includes("translation required")) return "warning";
  if (normalized.includes("review")) return "review";
  if (normalized.includes("updated")) return "updated";
  if (normalized === "yes" || normalized.includes("matched") || normalized.includes("not required")) return "matched";
  return "neutral";
}

function renderDictionarySummary() {
  const summary = state.dictionary.summary;
  const metrics = [
    ["Total Labels", summary.totalLabels],
    ["Matched Labels", summary.matchedLabels],
    ["Updated Labels", summary.updatedLabels],
    ["Labels Not Found", summary.notFound],
    ["Translation Required", summary.translationRequired],
    ["Needs Review", summary.needsReview],
    ["Processing Progress", `${summary.processingProgress || 0}%`],
  ];
  els.dictionarySummary.innerHTML = metrics.map(([label, value]) => `
    <article><span>${escapeHtml(label)}</span><strong>${typeof value === "string" ? escapeHtml(value) : formatNumber(value)}</strong></article>`).join("");
}

function renderDictionaryTable() {
  const filtered = dictionaryFilteredRows();
  const pageCount = Math.max(1, Math.ceil(filtered.length / state.dictionary.pageSize));
  state.dictionary.page = Math.min(state.dictionary.page, pageCount);
  const startIndex = (state.dictionary.page - 1) * state.dictionary.pageSize;
  const pageRows = filtered.slice(startIndex, startIndex + state.dictionary.pageSize);
  els.dictionaryBody.innerHTML = pageRows.length ? pageRows.map((row) => `
    <tr class="${row.dictionaryMatch === "No" || row.reviewStatus !== "Not Required" ? "dictionary-exception-row" : ""}">
      <td><strong>${formatNumber(row.rowNumber)}</strong><small>${escapeHtml(row.sheetName)}</small></td>
      <td class="text-cell">${escapeHtml(row.originalEnglish)}</td>
      <td class="text-cell">${escapeHtml(row.suggestedEnglish || row.standardizedEnglish || "-")}</td>
      <td class="text-cell rtl">${escapeHtml(row.arabicTranslation || "-")}</td>
      <td><span class="dictionary-badge ${dictionaryBadgeClass(row.matchStatus)}">${escapeHtml(row.matchStatus)}</span></td>
      <td><span class="dictionary-method">${escapeHtml(row.matchMethod || "-")} · ${formatNumber(row.matchConfidence)}%</span></td>
      <td><span class="dictionary-badge ${dictionaryBadgeClass(row.dictionaryMatch)}">${escapeHtml(row.dictionaryMatch)}</span></td>
      <td><span class="dictionary-badge ${dictionaryBadgeClass(row.reviewStatus)}">${escapeHtml(row.reviewStatus)}</span></td>
    </tr>`).join("") : `<tr><td colspan="8" class="empty-row">No labels match this result filter.</td></tr>`;

  const start = filtered.length ? startIndex + 1 : 0;
  const end = Math.min(startIndex + state.dictionary.pageSize, filtered.length);
  els.dictionaryTableInfo.textContent = `${formatNumber(filtered.length)} labels in this view`;
  els.dictionaryPageInfo.textContent = `Showing ${formatNumber(start)} to ${formatNumber(end)} of ${formatNumber(filtered.length)} labels`;
  els.dictionaryPageNumber.textContent = state.dictionary.page;
  els.dictionaryPrevPage.disabled = state.dictionary.page <= 1;
  els.dictionaryNextPage.disabled = state.dictionary.page >= pageCount;
}

function setDictionaryProgress(value, label) {
  els.dictionaryProgress.hidden = false;
  els.dictionaryProgressBar.value = value;
  els.dictionaryProgressValue.textContent = `${value}%`;
  els.dictionaryProgressLabel.textContent = label;
}

async function processDictionaryFile(file) {
  if (!file) return;
  if (!/\.(xlsx|xlsm)$/i.test(file.name)) {
    toast("Upload an .xlsx or .xlsm Excel workbook.");
    return;
  }
  if (file.size > 50 * 1024 * 1024) {
    toast("The workbook exceeds the 50 MB upload limit.");
    return;
  }

  els.dictionaryBrowseButton.disabled = true;
  els.dictionaryResults.hidden = true;
  els.dictionaryFileMeta.hidden = false;
  els.dictionaryFileMeta.innerHTML = `<strong>${escapeHtml(file.name)}</strong><span>${formatFileSize(file.size)} · selected ${formatDate(new Date())}</span>`;
  setDictionaryProgress(15, "Reading workbook...");
  try {
    const body = await file.arrayBuffer();
    setDictionaryProgress(45, "Matching labels with the All Labels Dictionary...");
    const data = await api("/api/dictionary/process", {
      method: "POST",
      headers: {
        "Content-Type": "application/octet-stream",
        "X-File-Name": encodeURIComponent(file.name),
      },
      body,
    });
    setDictionaryProgress(100, "Processing complete");
    state.dictionary.jobId = data.jobId;
    state.dictionary.rows = data.results || [];
    state.dictionary.summary = data.summary || {};
    state.dictionary.filter = "all";
    state.dictionary.page = 1;
    els.dictionaryFileMeta.innerHTML = `<strong>${escapeHtml(data.uploadedFile?.name || file.name)}</strong><span>${formatFileSize(data.uploadedFile?.size || file.size)} · uploaded ${formatDate(data.uploadedFile?.uploadedAt)}</span>`;
    els.dictionarySource.textContent = `Dictionary source: ${formatDate(data.dictionarySource?.fetchedAt)} · ${formatNumber(data.dictionarySource?.sourceTotals?.labels)} labels`;
    els.dictionaryFilters.querySelectorAll("button").forEach((button) => button.classList.toggle("active", button.dataset.dictionaryFilter === "all"));
    renderDictionarySummary();
    renderDictionaryTable();
    els.dictionaryResults.hidden = false;
    toast(`Processed ${formatNumber(data.summary?.totalLabels)} labels across ${formatNumber(data.summary?.sheetsProcessed)} sheets.`);
  } catch (error) {
    setDictionaryProgress(0, "Processing failed");
    toast(error.message);
  } finally {
    els.dictionaryBrowseButton.disabled = false;
    els.dictionaryFileInput.value = "";
  }
}

function downloadDictionaryFile(fileType) {
  if (!state.dictionary.jobId) {
    toast("Upload and process a workbook first.");
    return;
  }
  window.location.href = `/api/dictionary/jobs/${encodeURIComponent(state.dictionary.jobId)}/${fileType}`;
}

function renderHistory(data = {}) {
  const rows = data.rows || [];
  const totals = data.totals || {};
  const summary = `
    <div class="history-summary">
      <span>Total <strong>${formatNumber(totals.total)}</strong></span>
      <span>Created <strong>${formatNumber(totals.created)}</strong></span>
      <span>Modified <strong>${formatNumber(totals.modified)}</strong></span>
      <span>Removed <strong>${formatNumber(totals.removed)}</strong></span>
      <span>Feature Keys <strong>${formatNumber(totals.featureKeys)}</strong></span>
    </div>`;
  const table = rows.length ? `
    <div class="table-wrap">
      <table class="history-table">
        <thead>
          <tr>
            <th>Date / Time</th>
            <th>Change</th>
            <th>Status</th>
            <th>Actor Role</th>
            <th>Changed By</th>
            <th>Source</th>
            <th>Feature Key</th>
            <th>Label Key</th>
            <th>English Text</th>
            <th>Arabic Text</th>
            <th>Summary</th>
          </tr>
        </thead>
        <tbody>${rows.map((row) => `
          <tr>
            <td>${formatDate(row.changedAt)}</td>
            <td><span class="change ${badgeClass(row.changeType)}">${escapeHtml(row.changeType)}</span></td>
            <td><span class="status">${escapeHtml(row.status || "-")}</span></td>
            <td>${escapeHtml(row.actorRole || `${row.changeType} By`)}</td>
            <td>${actorCell(row)}</td>
            <td>${escapeHtml(row.source || "-")}</td>
            <td>${escapeHtml(row.featureKey || "-")}</td>
            <td>${escapeHtml(row.labelKey || "-")}</td>
            <td class="text-cell">${escapeHtml(row.englishText || "-")}</td>
            <td class="text-cell rtl">${escapeHtml(row.arabicText || "-")}</td>
            <td>${escapeHtml(row.changeSummary || "-")}</td>
          </tr>`).join("")}</tbody>
      </table>
    </div>` : `<div class="empty-box">No content history matches the current filters.</div>`;
  els.auditContent.innerHTML = `${summary}${table}`;
}

function renderHistoryConnection(status = {}) {
  state.adminSessionActive = Boolean(status.adminSessionActive);
  els.historyAdminEmail.hidden = state.adminSessionActive;
  els.historyAdminPassword.hidden = state.adminSessionActive;
  els.historyAdminEmail.closest("label").hidden = state.adminSessionActive;
  els.historyAdminPassword.closest("label").hidden = state.adminSessionActive;
  els.historyConnectButton.hidden = state.adminSessionActive;
  els.historyDisconnectButton.hidden = !state.adminSessionActive;
  els.historySyncButton.disabled = !state.adminSessionActive;
  if (state.adminSessionActive) {
    const sync = status.historySync;
    els.historyConnectionStatus.textContent = sync?.syncedAt
      ? `Connected. Last Strapi history sync: ${formatDate(sync.syncedAt)} · ${formatNumber(sync.eventsFound)} events from ${formatNumber(sync.documentsChecked)} entries.`
      : "Connected. Sync Strapi History to collect revision users and dates.";
  } else {
    const sync = status.historySync;
    els.historyConnectionStatus.textContent = sync?.syncedAt
      ? `Last Strapi history sync: ${formatDate(sync.syncedAt)} · ${formatNumber(sync.eventsFound)} actor-aware events. Connect again to refresh it.`
      : "Connect your Strapi admin account to fetch revision dates and user names. Credentials stay in server memory only.";
  }
}

async function connectHistory() {
  const email = els.historyAdminEmail.value.trim();
  const password = els.historyAdminPassword.value;
  els.historyConnectButton.disabled = true;
  els.historyConnectButton.textContent = "Connecting...";
  try {
    const result = await api("/api/admin-login", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email, password }),
    });
    els.historyAdminPassword.value = "";
    toast(result.message);
    await loadStatus();
  } catch (error) {
    toast(error.message);
  } finally {
    els.historyConnectButton.disabled = false;
    els.historyConnectButton.textContent = "Connect";
  }
}

async function syncHistory() {
  els.historySyncButton.disabled = true;
  els.historySyncButton.textContent = "Syncing all entries...";
  try {
    const result = await api("/api/history/sync", { method: "POST" });
    toast(result.message);
    await loadStatus();
    await loadHistory();
  } catch (error) {
    toast(error.message);
  } finally {
    els.historySyncButton.textContent = "Sync Strapi History";
    els.historySyncButton.disabled = !state.adminSessionActive;
  }
}

async function disconnectHistory() {
  const result = await api("/api/admin-logout", { method: "POST" });
  els.historyAdminPassword.value = "";
  toast(result.message);
  await loadStatus();
}

async function loadHistory() {
  const data = await api(`/api/history?${historyQueryParams()}`);
  state.historyRows = data.rows || [];
  state.historyTotal = data.total || 0;
  renderHistory(data);
}

async function loadStatus() {
  const status = await api("/api/status");
  renderHistoryConnection(status);
  renderStats(status.totals || {});
  renderOptions(status.options || {});
  renderSnapshots(status.snapshots || []);
  renderSidePanels(status.totals || {});
  await loadHistory();
  if (status.lastSnapshot) {
    els.lastFetch.textContent = `Last fetch: ${formatDate(status.lastSnapshot.fetchedAt)}`;
    els.dictionarySource.textContent = `Dictionary source: ${formatDate(status.lastSnapshot.fetchedAt)} · ${formatNumber(status.lastSnapshot.sourceTotals?.labels)} labels`;
  }
}

async function loadLabels() {
  const data = await api(`/api/labels?${queryParams()}`);
  state.rows = data.rows || [];
  state.total = data.total || 0;
  renderOptions(data.options || {});
  renderSnapshots(data.snapshots || []);
  renderSidePanels(data.totals || {});
  renderTable();
  updateActivePill();
}

async function fetchLabels() {
  els.fetchButton.disabled = true;
  els.fetchButton.textContent = "Fetching...";
  try {
    const data = await api("/api/fetch", { method: "POST" });
    toast(data.message || "Labels loaded");
    await loadStatus();
    await loadLabels();
    await loadHistory();
    await loadLanguageConsistency();
  } catch (error) {
    toast(error.message);
  } finally {
    els.fetchButton.disabled = false;
    els.fetchButton.textContent = "Fetch All Labels";
  }
}

function updateActivePill() {
  const active = [];
  if (state.filters.search) active.push(`Search: ${state.filters.search}`);
  if (state.filters.featureKey !== "all") active.push(`Feature: ${state.filters.featureKey}`);
  if (state.filters.status !== "all") active.push(`Status: ${state.filters.status}`);
  if (state.filters.changeType !== "all") active.push(`Change: ${state.filters.changeType}`);
  if (state.filters.languageIssue !== "all") active.push(`Language: ${state.filters.languageIssue}`);
  if (Object.values(state.advanced).some(Boolean)) active.push("Advanced search");
  els.activeFilterPill.hidden = !active.length;
  els.activeFilterPill.textContent = active.join(" · ");
}

function setFilter(key, value) {
  state.filters[key] = value;
  state.page = 1;
  loadLabels().catch((error) => toast(error.message));
}

function openDrawer(row) {
  els.drawerTitle.textContent = row.labelKey || row.labelId;
  const displayedIssue = row.issueType || row.languageIssue || "None";
  els.drawerSubtitle.textContent = `${row.featureKey || "-"} · ${row.changeType || "-"} · ${displayedIssue}`;
  const consistencyFields = row.issueType ? [
    ["Consistency Issue", row.issueType],
    ["Occurrence Count", row.occurrenceCount],
    ["Issue Details", row.details],
    ["Observed Variants", (row.variants || []).join(" | ") || "-"],
  ] : [];
  const fields = [
    ...consistencyFields,
    ["Label ID", row.labelId],
    ["Legacy Label ID", row.legacyLabelId],
    ["Feature Key", row.featureKey],
    ["Label Key", row.labelKey],
    ["English Text", row.englishText],
    ["Arabic Text", row.arabicText],
    ["Description", row.description],
    ["Module", row.module],
    ["Screen Name", row.screenName],
    ["Component", `${row.component || "-"} (${row.componentId || "-"})`],
    ["Status", row.status],
    ["Change Type", row.changeType],
    ["Actor Role", row.actorRole || "No activity in last 7 days"],
    ["Changed By", row.changedBy || "-"],
    ["Changed By Username", row.changedByUsername || "-"],
    ["Changed By Email", row.changedByEmail || "-"],
    ["History Changed Date", formatDate(row.historyChangedAt)],
    ["Language Issue", row.languageIssue],
    ["First Seen Date", formatDate(row.firstSeenAt)],
    ["Last Modified Date", formatDate(row.lastModifiedAt)],
    ["Entry ID", row.entryId],
    ["Document ID", row.documentId],
    ["Locale", row.locale],
    ["Comments", row.comments || "-"],
  ];
  els.drawerContent.innerHTML = fields.map(([label, value]) => `
    <div class="detail-field"><span>${escapeHtml(label)}</span><p>${escapeHtml(value || "-")}</p></div>`).join("");
  els.drawerBackdrop.hidden = false;
}

function closeDrawer() {
  els.drawerBackdrop.hidden = true;
}

async function loadSnapshotChange(snapshotId, change) {
  const data = await api(`/api/snapshots/${encodeURIComponent(snapshotId)}/${encodeURIComponent(change)}`);
  state.rows = data.rows || [];
  state.total = data.total || 0;
  state.page = 1;
  renderTable();
  renderStats(data.totals || {});
  toast(`Loaded ${formatNumber(data.total)} ${change} labels from snapshot`);
  navigate("labels", { loadLabelsPage: false });
}

const ROUTE_BY_TARGET = {
  dashboard: "/dashboard",
  labels: "/labels",
  "labels-new": "/labels/new",
  "labels-modified": "/labels/modified",
  "labels-removed": "/labels/removed",
  compare: "/dictionary-generator",
  snapshots: "/snapshots",
  language: "/language-issues",
  reports: "/reports",
  settings: "/settings",
  audit: "/content-history",
  api: "/api-status",
};

const LABEL_PAGE_CONFIG = {
  labels: {
    title: "All Labels",
    description: "Search all labels by Label ID, Feature Key, Label Key, English Text, or Arabic Text.",
    changeType: "all",
    exportLabel: "Export All Labels",
  },
  "labels-new": {
    title: "New Labels",
    description: "Labels newly introduced in the latest Strapi comparison.",
    changeType: "New",
    exportLabel: "Export New Labels",
  },
  "labels-modified": {
    title: "Modified Labels",
    description: "Existing labels whose English, Arabic, key, or related content changed.",
    changeType: "Modified",
    exportLabel: "Export Modified Labels",
  },
  "labels-removed": {
    title: "Removed Labels",
    description: "Labels present previously but missing from the latest Strapi comparison.",
    changeType: "Removed",
    exportLabel: "Export Removed Labels",
  },
};

const TARGET_BY_ROUTE = {
  ...Object.fromEntries(Object.entries(ROUTE_BY_TARGET).map(([target, route]) => [route, target])),
  "/": "labels",
  "/compare": "compare",
  "/language": "language",
  "/audit": "audit",
  "/api": "api",
  "/dahbaord": "dashboard",
};

function navigate(target, { historyMode = "push", loadLabelsPage = true } = {}) {
  const labelPage = LABEL_PAGE_CONFIG[target];
  const pageTarget = labelPage ? "labelsPage" : target;
  const route = ROUTE_BY_TARGET[target] || ROUTE_BY_TARGET.labels;
  state.currentTarget = target;
  if (labelPage) {
    state.filters.changeType = labelPage.changeType;
    state.page = 1;
    els.labelsPageTitle.textContent = labelPage.title;
    els.labelsPageDescription.textContent = labelPage.description;
    els.labelsPageExport.textContent = labelPage.exportLabel;
    els.changeFilter.value = labelPage.changeType;
  }
  if (historyMode === "push" && window.location.pathname !== route) window.history.pushState({ target }, "", route);
  if (historyMode === "replace" && window.location.pathname !== route) window.history.replaceState({ target }, "", route);
  document.body.classList.toggle("labels-only-mode", Boolean(labelPage));
  document.querySelectorAll(".page").forEach((page) => page.classList.toggle("active-page", page.id === pageTarget));
  document.querySelectorAll(".side-nav [data-nav-target]").forEach((link) => link.classList.toggle("active", link.dataset.navTarget === target));
  if (labelPage) {
    window.scrollTo({ top: 0, behavior: "smooth" });
    if (state.ready && loadLabelsPage) loadLabels().catch((error) => toast(error.message));
  }
  if (target === "dashboard" && state.ready) loadStatus().catch((error) => toast(error.message));
  if (target === "audit") loadHistory().catch((error) => toast(error.message));
  if (target === "language") loadLanguageConsistency().catch((error) => toast(error.message));
}

function downloadExcel() {
  window.location.href = `/api/download/xlsx?${queryParams(false)}`;
}

let searchTimer;
els.globalSearch.addEventListener("input", (event) => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => setFilter("search", event.target.value), 250);
});
function bindTopSearch(input, field) {
  input.addEventListener("input", () => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      state.advanced[field] = input.value.trim();
      state.page = 1;
      loadLabels().catch((error) => toast(error.message));
    }, 250);
  });
}
bindTopSearch(els.featureKeySearch, "featureKey");
bindTopSearch(els.englishTextSearch, "englishText");
bindTopSearch(els.arabicTextSearch, "arabicText");
els.featureFilter.addEventListener("change", (event) => setFilter("featureKey", event.target.value));
els.statusFilter.addEventListener("change", (event) => setFilter("status", event.target.value));
els.changeFilter.addEventListener("change", (event) => setFilter("changeType", event.target.value));
els.languageIssueFilter.addEventListener("change", (event) => setFilter("languageIssue", event.target.value));
els.languageSearch.addEventListener("input", (event) => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => {
    state.language.filters.search = event.target.value.trim();
    state.language.page = 1;
    loadLanguageConsistency().catch((error) => toast(error.message));
  }, 250);
});
els.languageIssueTypeFilter.addEventListener("change", (event) => {
  state.language.filters.issueType = event.target.value;
  state.language.page = 1;
  loadLanguageConsistency().catch((error) => toast(error.message));
});
els.languageFeatureFilter.addEventListener("change", (event) => {
  state.language.filters.featureKey = event.target.value;
  state.language.page = 1;
  loadLanguageConsistency().catch((error) => toast(error.message));
});
els.languageStatusFilter.addEventListener("change", (event) => {
  state.language.filters.status = event.target.value;
  state.language.page = 1;
  loadLanguageConsistency().catch((error) => toast(error.message));
});
els.languagePageSize.addEventListener("change", (event) => {
  state.language.pageSize = Number(event.target.value);
  state.language.page = 1;
  loadLanguageConsistency().catch((error) => toast(error.message));
});
els.languagePrevPage.addEventListener("click", () => {
  if (state.language.page <= 1) return;
  state.language.page -= 1;
  loadLanguageConsistency().catch((error) => toast(error.message));
});
els.languageNextPage.addEventListener("click", () => {
  if (state.language.page * state.language.pageSize >= state.language.total) return;
  state.language.page += 1;
  loadLanguageConsistency().catch((error) => toast(error.message));
});
els.historySearch.addEventListener("input", (event) => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => {
    state.historyFilters.search = event.target.value.trim();
    loadHistory().catch((error) => toast(error.message));
  }, 250);
});
els.historyFeatureFilter.addEventListener("change", (event) => {
  state.historyFilters.featureKey = event.target.value;
  loadHistory().catch((error) => toast(error.message));
});
els.historyChangeFilter.addEventListener("change", (event) => {
  state.historyFilters.changeType = event.target.value;
  loadHistory().catch((error) => toast(error.message));
});
els.historySourceFilter.addEventListener("change", (event) => {
  state.historyFilters.source = event.target.value;
  loadHistory().catch((error) => toast(error.message));
});
els.historyDateFrom.addEventListener("change", (event) => {
  state.historyFilters.dateFrom = event.target.value;
  loadHistory().catch((error) => toast(error.message));
});
els.historyDateTo.addEventListener("change", (event) => {
  state.historyFilters.dateTo = event.target.value;
  loadHistory().catch((error) => toast(error.message));
});
els.historyConnectButton.addEventListener("click", connectHistory);
els.historySyncButton.addEventListener("click", syncHistory);
els.historyDisconnectButton.addEventListener("click", () => disconnectHistory().catch((error) => toast(error.message)));
els.historyAdminPassword.addEventListener("keydown", (event) => {
  if (event.key === "Enter") connectHistory();
});
els.pageSize.addEventListener("change", (event) => { state.pageSize = Number(event.target.value); state.page = 1; loadLabels().catch((error) => toast(error.message)); });
els.prevPage.addEventListener("click", () => { if (state.page > 1) { state.page -= 1; loadLabels().catch((error) => toast(error.message)); } });
els.nextPage.addEventListener("click", () => { if (state.page * state.pageSize < state.total) { state.page += 1; loadLabels().catch((error) => toast(error.message)); } });
els.fetchButton.addEventListener("click", fetchLabels);
els.refreshButton.addEventListener("click", () => loadStatus().then(loadLabels).then(loadLanguageConsistency).catch((error) => toast(error.message)));
els.downloadExcel.addEventListener("click", downloadExcel);
els.downloadExcel2.addEventListener("click", downloadExcel);
els.labelsPageExport.addEventListener("click", downloadExcel);
els.dictionaryBrowseButton.addEventListener("click", () => els.dictionaryFileInput.click());
els.dictionaryDropZone.addEventListener("click", () => els.dictionaryFileInput.click());
els.dictionaryFileInput.addEventListener("change", () => processDictionaryFile(els.dictionaryFileInput.files?.[0]));
els.dictionaryDropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  els.dictionaryDropZone.classList.add("dragging");
});
els.dictionaryDropZone.addEventListener("dragleave", () => els.dictionaryDropZone.classList.remove("dragging"));
els.dictionaryDropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  els.dictionaryDropZone.classList.remove("dragging");
  processDictionaryFile(event.dataTransfer?.files?.[0]);
});
els.dictionaryFilters.addEventListener("click", (event) => {
  const button = event.target.closest("[data-dictionary-filter]");
  if (!button) return;
  state.dictionary.filter = button.dataset.dictionaryFilter;
  state.dictionary.page = 1;
  els.dictionaryFilters.querySelectorAll("button").forEach((item) => item.classList.toggle("active", item === button));
  renderDictionaryTable();
});
els.dictionaryPageSize.addEventListener("change", (event) => {
  state.dictionary.pageSize = Number(event.target.value);
  state.dictionary.page = 1;
  renderDictionaryTable();
});
els.dictionaryPrevPage.addEventListener("click", () => {
  if (state.dictionary.page <= 1) return;
  state.dictionary.page -= 1;
  renderDictionaryTable();
});
els.dictionaryNextPage.addEventListener("click", () => {
  const totalPages = Math.ceil(dictionaryFilteredRows().length / state.dictionary.pageSize);
  if (state.dictionary.page >= totalPages) return;
  state.dictionary.page += 1;
  renderDictionaryTable();
});
els.dictionaryDownloadUpdated.addEventListener("click", () => downloadDictionaryFile("updated"));
els.dictionaryDownloadExceptions.addEventListener("click", () => downloadDictionaryFile("exceptions"));
els.dictionaryDownloadSummary.addEventListener("click", () => downloadDictionaryFile("summary"));
els.advancedToggle.addEventListener("click", () => { els.advancedPanel.hidden = !els.advancedPanel.hidden; });
els.clearAdvanced.addEventListener("click", () => {
  state.advanced = {};
  document.querySelectorAll("[data-advanced]").forEach((input) => { input.value = ""; });
  state.page = 1;
  loadLabels().catch((error) => toast(error.message));
});
document.querySelectorAll("[data-advanced]").forEach((input) => {
  input.addEventListener("input", () => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      state.advanced[input.dataset.advanced] = input.value.trim();
      state.page = 1;
      loadLabels().catch((error) => toast(error.message));
    }, 250);
  });
});
els.labelsBody.addEventListener("click", (event) => {
  const button = event.target.closest(".row-action");
  if (!button) return;
  openDrawer(state.rows[Number(button.dataset.index)]);
});
els.languageBody.addEventListener("click", (event) => {
  const button = event.target.closest(".language-row-action");
  if (!button) return;
  openDrawer(state.language.rows[Number(button.dataset.index)]);
});
els.closeDrawer.addEventListener("click", closeDrawer);
els.drawerBackdrop.addEventListener("click", (event) => { if (event.target === els.drawerBackdrop) closeDrawer(); });
document.addEventListener("click", (event) => {
  const snapshotButton = event.target.closest("[data-snapshot][data-change]");
  if (snapshotButton) {
    loadSnapshotChange(snapshotButton.dataset.snapshot, snapshotButton.dataset.change).catch((error) => toast(error.message));
    return;
  }
  const quick = event.target.closest("[data-quick-change]");
  if (quick) {
    els.changeFilter.value = quick.dataset.quickChange;
    setFilter("changeType", quick.dataset.quickChange);
  }
});
document.querySelectorAll(".side-nav [data-nav-target]").forEach((link) => {
  link.addEventListener("click", (event) => {
    event.preventDefault();
    navigate(link.dataset.navTarget);
  });
});
document.querySelectorAll("[data-nav-card]").forEach((card) => {
  card.addEventListener("click", () => navigate(card.dataset.navCard));
});
window.addEventListener("popstate", () => navigate(TARGET_BY_ROUTE[window.location.pathname] || "labels", { historyMode: "none" }));

const initialTarget = TARGET_BY_ROUTE[window.location.pathname] || "labels";
navigate(initialTarget, { historyMode: window.location.pathname === "/" ? "replace" : "none" });
loadStatus().then(() => {
  state.ready = true;
  return LABEL_PAGE_CONFIG[state.currentTarget] ? loadLabels() : null;
}).catch((error) => toast(error.message));
