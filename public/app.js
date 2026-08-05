const state = {
  rows: [],
  allRows: [],
  selected: null,
  page: 1,
  pageSize: 25,
  filters: {
    search: "",
    brand: "all",
    category: "all",
    group: "all",
    publishStatus: "all",
    changeType: "all",
    sort: "updatedAt:desc",
    createdFrom: "",
    createdTo: "",
    modifiedFrom: "",
    modifiedTo: "",
  },
};

const labsState = {
  rows: [],
  allRows: [],
  page: 1,
  pageSize: 25,
  filters: {
    search: "",
    publishStatus: "all",
    changeType: "all",
    sort: "updatedAt:desc",
  },
};

const $ = (id) => document.getElementById(id);
const PUBLISH_STATUS_OPTIONS = ["Draft", "Published", "Modified After Publish", "Pending Review", "Modified"];

const els = {
  fetchButton: $("fetchButton"),
  refreshStatus: $("refreshStatus"),
  lastFetch: $("lastFetch"),
  sourceNote: $("sourceNote"),
  adminModeStatus: $("adminModeStatus"),
  adminLoginForm: $("adminLoginForm"),
  adminEmail: $("adminEmail"),
  adminPassword: $("adminPassword"),
  adminLoginButton: $("adminLoginButton"),
  adminLogoutButton: $("adminLogoutButton"),
  snapshotList: $("snapshotList"),
  devicesBody: $("devicesBody"),
  searchInput: $("searchInput"),
  brandFilter: $("brandFilter"),
  categoryFilter: $("categoryFilter"),
  groupFilter: $("groupFilter"),
  publishFilter: $("publishFilter"),
  changeFilter: $("changeFilter"),
  sortSelect: $("sortSelect"),
  createdFrom: $("createdFrom"),
  createdTo: $("createdTo"),
  modifiedFrom: $("modifiedFrom"),
  modifiedTo: $("modifiedTo"),
  clearFilters: $("clearFilters"),
  pageInfo: $("pageInfo"),
  prevPage: $("prevPage"),
  nextPage: $("nextPage"),
  drawer: $("deviceDrawer"),
  drawerTitle: $("drawerTitle"),
  drawerSubtitle: $("drawerSubtitle"),
  drawerContent: $("drawerContent"),
  closeDrawer: $("closeDrawer"),
  toast: $("toast"),
  statusChart: $("statusChart"),
  recentChanges: $("recentChanges"),
  latestCreated: $("latestCreated"),
  removedLog: $("removedLog"),
  compareList: $("compareList"),
  compareDetails: $("compareDetails"),
  labsCollectionUid: $("labsCollectionUid"),
  labsSearchInput: $("labsSearchInput"),
  labsPublishFilter: $("labsPublishFilter"),
  labsChangeFilter: $("labsChangeFilter"),
  fetchLabsButton: $("fetchLabsButton"),
  labsBody: $("labsBody"),
  labsPageInfo: $("labsPageInfo"),
  labsPrevPage: $("labsPrevPage"),
  labsNextPage: $("labsNextPage"),
  openLabsEn: $("openLabsEn"),
  openLabsAr: $("openLabsAr"),
};

const statIds = {
  total: "statTotal",
  new: "statNew",
  modified: "statModified",
  removed: "statRemoved",
  published: "statPublished",
  draft: "statDraft",
  missingArabic: "statMissingArabic",
  missingImages: "statMissingImages",
  missingMandatory: "statMissingMandatory",
};

const labsStatIds = {
  total: "labsStatTotal",
  new: "labsStatNew",
  modified: "labsStatModified",
  removed: "labsStatRemoved",
  published: "labsStatPublished",
  draft: "labsStatDraft",
  missingArabic: "labsStatMissingArabic",
  missingImages: "labsStatMissingImages",
};

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function formatDate(value) {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleString(undefined, { year: "numeric", month: "short", day: "2-digit", hour: "2-digit", minute: "2-digit" });
}

function listText(value) {
  if (Array.isArray(value)) return value.filter(Boolean).join(", ");
  return String(value ?? "");
}

function showToast(message) {
  els.toast.textContent = message;
  els.toast.hidden = false;
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => { els.toast.hidden = true; }, 3600);
}

function setBusy(isBusy) {
  els.fetchButton.disabled = isBusy;
  els.fetchButton.textContent = isBusy ? "Refreshing..." : "Refresh CMS Data";
}

function queryParams() {
  const params = new URLSearchParams();
  for (const [key, value] of Object.entries(state.filters)) {
    if (value && value !== "all") params.set(key, value);
  }
  params.set("page", state.page);
  params.set("pageSize", state.pageSize);
  return params;
}

function labsQueryParams() {
  const params = new URLSearchParams();
  for (const [key, value] of Object.entries(labsState.filters)) {
    if (value && value !== "all") params.set(key, value);
  }
  params.set("page", labsState.page);
  params.set("pageSize", labsState.pageSize);
  return params;
}

async function api(path, options) {
  const response = await fetch(path, options);
  const json = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(json.error || `Request failed: ${response.status}`);
  return json;
}

async function loadStatus() {
  const status = await api("/api/status");
  renderSnapshots(status.snapshots || []);
  renderStats(status.totals || {});
  renderOptions(status.options || {});
  els.sourceNote.textContent = status.sourceNote || "Public Strapi device API";
  renderAdminStatus(status);
  if (status.lastSnapshot) els.lastFetch.textContent = `Last refresh ${formatDate(status.lastSnapshot.fetchedAt)} · ${status.lastSnapshot.sourceTotals?.current ?? 0} current devices`;
  return status;
}

function renderAdminStatus(status = {}) {
  if (!els.adminModeStatus) return;
  const isAdmin = status.sourceMode === "admin";
  els.adminModeStatus.textContent = isAdmin
    ? (status.adminSessionActive ? "Admin session connected" : "Admin token active")
    : "Public API mode";
  els.adminModeStatus.classList.toggle("connected", isAdmin);
  els.adminLogoutButton.disabled = !isAdmin || status.envAdminTokenConfigured;
}

async function loadDevices() {
  const data = await api(`/api/devices?${queryParams()}`);
  state.rows = data.rows || [];
  state.allRows = data.allRows || [];
  renderStats(data.totals || {});
  renderOptions(data.options || {});
  renderTable(data.total || 0);
  renderOverview(data.allRows || []);
  renderCompare(data.allRows || []);
  renderSnapshots(data.snapshots || []);
}

async function loadLabsStatus() {
  const status = await api("/api/labs/status");
  updateLabsLinks(status);
  if (els.labsCollectionUid && !els.labsCollectionUid.value) els.labsCollectionUid.value = status.collectionUid || "";
  renderLabsStats(status.totals || {});
  fillSelect(els.labsPublishFilter, "All publishing states", PUBLISH_STATUS_OPTIONS, labsState.filters.publishStatus);
  return status;
}

async function loadLabs() {
  const data = await api(`/api/labs?${labsQueryParams()}`);
  labsState.rows = data.rows || [];
  labsState.allRows = data.allRows || [];
  renderLabsStats(data.totals || {});
  updateLabsLinks(data);
  if (els.labsCollectionUid && !els.labsCollectionUid.value) els.labsCollectionUid.value = data.collectionUid || "";
  renderLabsTable(data.total || 0);
}

async function refreshCms() {
  setBusy(true);
  try {
    const data = await api("/api/fetch", { method: "POST" });
    state.page = 1;
    state.rows = data.rows || [];
    state.allRows = data.rows || [];
    renderStats(data.totals || {});
    renderOptions(data.options || {});
    renderSnapshots(data.snapshots || []);
    renderOverview(data.rows || []);
    renderCompare(data.rows || []);
    els.lastFetch.textContent = `Last refresh ${formatDate(data.snapshot?.fetchedAt)} · EN ${data.sourceTotals?.en ?? 0}, AR ${data.sourceTotals?.ar ?? 0}, merged ${data.sourceTotals?.merged ?? 0}`;
    showToast(`Fetched ${data.sourceTotals?.merged ?? 0} device assets`);
    await loadDevices();
  } catch (error) {
    showToast(error.message);
  } finally {
    setBusy(false);
  }
}

function setLabsBusy(isBusy) {
  els.fetchLabsButton.disabled = isBusy;
  els.fetchLabsButton.textContent = isBusy ? "Refreshing..." : "Refresh Labs";
}

async function refreshLabs() {
  setLabsBusy(true);
  try {
    const data = await api("/api/labs/fetch", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ collectionUid: els.labsCollectionUid.value.trim() }),
    });
    labsState.page = 1;
    labsState.rows = data.rows || [];
    labsState.allRows = data.rows || [];
    updateLabsLinks(data);
    renderLabsStats(data.totals || {});
    renderLabsTable(data.rows?.length || 0);
    showToast(`Fetched ${data.sourceTotals?.merged ?? 0} labs records`);
    await loadLabs();
  } catch (error) {
    showToast(error.message);
  } finally {
    setLabsBusy(false);
  }
}

async function connectAdmin(event) {
  event.preventDefault();
  const email = els.adminEmail.value.trim();
  const password = els.adminPassword.value;
  if (!email || !password) {
    showToast("Enter both Strapi email and password.");
    return;
  }
  els.adminLoginButton.disabled = true;
  els.adminLoginButton.textContent = "Connecting...";
  try {
    const result = await api("/api/admin-login", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email, password }),
    });
    els.adminPassword.value = "";
    showToast(result.adminTotal ? `Admin connected. Strapi reports ${result.adminTotal} English entries.` : "Admin connected. Refreshing CMS data.");
    await refreshCms();
    await loadStatus();
  } catch (error) {
    els.adminPassword.value = "";
    showToast(error.message);
  } finally {
    els.adminLoginButton.disabled = false;
    els.adminLoginButton.textContent = "Connect Admin";
  }
}

async function disconnectAdmin() {
  els.adminLogoutButton.disabled = true;
  try {
    const result = await api("/api/admin-logout", { method: "POST" });
    showToast(result.message || "Admin session cleared.");
    await loadStatus();
  } catch (error) {
    showToast(error.message);
  } finally {
    els.adminLogoutButton.disabled = false;
  }
}

function renderStats(totals = {}) {
  const values = {
    total: totals.total ?? 0,
    new: totals.new ?? 0,
    modified: totals.modified ?? 0,
    removed: totals.removed ?? 0,
    published: totals.published ?? 0,
    draft: totals.draft ?? 0,
    missingArabic: totals.missingArabic ?? 0,
    missingImages: totals.missingImages ?? 0,
    missingMandatory: totals.missingMandatory ?? 0,
  };
  for (const [key, id] of Object.entries(statIds)) $(id).textContent = values[key];
}

function renderLabsStats(totals = {}) {
  const values = {
    total: totals.total ?? 0,
    new: totals.new ?? 0,
    modified: totals.modified ?? 0,
    removed: totals.removed ?? 0,
    published: totals.published ?? 0,
    draft: totals.draft ?? 0,
    missingArabic: totals.missingArabic ?? 0,
    missingImages: totals.missingImages ?? 0,
  };
  for (const [key, id] of Object.entries(labsStatIds)) $(id).textContent = values[key];
}

function updateLabsLinks(data = {}) {
  const collectionUid = data.collectionUid || els.labsCollectionUid?.value || "api::stc-label.stc-label";
  const enUrl = data.adminUrlEn || `https://content.stc.com.kw/admin/content-manager/collection-types/${collectionUid}?plugins%5Bi18n%5D%5Blocale%5D=en`;
  const arUrl = data.adminUrlAr || `https://content.stc.com.kw/admin/content-manager/collection-types/${collectionUid}?plugins%5Bi18n%5D%5Blocale%5D=ar`;
  els.openLabsEn.href = enUrl;
  els.openLabsAr.href = arUrl;
}

function renderOptions(options) {
  fillSelect(els.brandFilter, "All brands", options.brands, state.filters.brand);
  fillSelect(els.categoryFilter, "All categories", options.categories, state.filters.category);
  fillSelect(els.groupFilter, "All device groups", options.groups, state.filters.group);
  fillSelect(els.publishFilter, "All publishing states", mergeOptions(PUBLISH_STATUS_OPTIONS, options.publishStatuses), state.filters.publishStatus);
}

function mergeOptions(defaults = [], values = []) {
  return [...new Set([...defaults, ...(values || [])].filter(Boolean))];
}

function fillSelect(select, first, values = [], selected = "all") {
  const current = selected || select.value || "all";
  select.innerHTML = `<option value="all">${first}</option>${values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join("")}`;
  select.value = [...select.options].some((option) => option.value === current) ? current : "all";
}

function renderSnapshots(snapshots = []) {
  if (!snapshots.length) {
    els.snapshotList.className = "snapshot-list empty";
    els.snapshotList.textContent = "No snapshots yet";
    return;
  }
  els.snapshotList.className = "snapshot-list";
  els.snapshotList.innerHTML = snapshots.map((snapshot) => `
    <button class="snapshot-item" type="button" data-snapshot="${escapeHtml(snapshot.id)}">
      <strong>${formatDate(snapshot.fetchedAt)}</strong>
      <span>${snapshot.counts?.total ?? 0} devices · New ${snapshot.counts?.new ?? 0} · Modified ${snapshot.counts?.modified ?? 0} · Removed ${snapshot.counts?.removed ?? 0}</span>
    </button>
  `).join("");
}

function renderTable(total) {
  if (!state.rows.length) {
    els.devicesBody.innerHTML = `<tr><td colspan="11" class="empty-row">No matching devices.</td></tr>`;
  } else {
    els.devicesBody.innerHTML = state.rows.map((row) => `
      <tr>
        <td><strong>${escapeHtml(row.assetName || "-")}</strong><small>${escapeHtml(row.documentId)}</small></td>
        <td>
          <strong>${escapeHtml(row.deviceNameEn || "-")}</strong>
          <small class="rtl">${escapeHtml(row.deviceNameAr || "Arabic missing")}</small>
        </td>
        <td>${escapeHtml(row.brandName || "-")}</td>
        <td>${escapeHtml(row.deviceType || "-")}</td>
        <td>${escapeHtml(row.deviceGroupId || "-")}</td>
        <td><span class="badge ${badgeClass(row.publishStatus)}">${escapeHtml(row.publishStatus || "-")}</span></td>
        <td><span class="badge ${badgeClass(row.changeType)}">${escapeHtml(row.changeType || "-")}</span></td>
        <td>${row.missingArabic ? '<span class="badge warning">Missing</span>' : '<span class="badge published">Available</span>'}</td>
        <td>${row.missingImages ? '<span class="badge warning">Missing</span>' : `<span class="badge published">${(row.mediaUrls || []).length}</span>`}</td>
        <td>${formatDate(row.updatedAt)}</td>
        <td><button class="row-action" type="button" data-open="${escapeHtml(row.documentId)}">View</button></td>
      </tr>
    `).join("");
  }
  const start = total ? (state.page - 1) * state.pageSize + 1 : 0;
  const end = Math.min(total, state.page * state.pageSize);
  els.pageInfo.textContent = `${start}-${end} of ${total} devices`;
  els.prevPage.disabled = state.page <= 1;
  els.nextPage.disabled = end >= total;
}

function renderLabsTable(total) {
  if (!labsState.rows.length) {
    els.labsBody.innerHTML = `<tr><td colspan="9" class="empty-row">No matching labs records.</td></tr>`;
  } else {
    els.labsBody.innerHTML = labsState.rows.map((row) => `
      <tr>
        <td>
          <strong>${escapeHtml(row.titleEn || "-")}</strong>
          <small class="rtl">${escapeHtml(row.titleAr || "Arabic missing")}</small>
          <small>${escapeHtml(row.documentId)}</small>
        </td>
        <td>${escapeHtml(row.collectionUid || "-")}</td>
        <td><span class="badge ${badgeClass(row.publishStatus)}">${escapeHtml(row.publishStatus || "-")}</span></td>
        <td><span class="badge ${badgeClass(row.changeType)}">${escapeHtml(row.changeType || "-")}</span></td>
        <td>${row.missingArabic ? '<span class="badge warning">Missing</span>' : '<span class="badge published">Available</span>'}</td>
        <td>${(row.mediaUrls || []).length ? `<span class="badge published">${(row.mediaUrls || []).length}</span>` : '<span class="badge warning">Missing</span>'}</td>
        <td>EN ${escapeHtml(row.fieldCountEn ?? 0)} · AR ${escapeHtml(row.fieldCountAr ?? 0)}</td>
        <td>${formatDate(row.updatedAt)}</td>
        <td><button class="row-action" type="button" data-open-lab="${escapeHtml(row.documentId)}">View</button></td>
      </tr>
    `).join("");
  }
  const start = total ? (labsState.page - 1) * labsState.pageSize + 1 : 0;
  const end = Math.min(total, labsState.page * labsState.pageSize);
  els.labsPageInfo.textContent = `${start}-${end} of ${total} labs`;
  els.labsPrevPage.disabled = labsState.page <= 1;
  els.labsNextPage.disabled = end >= total;
}

function badgeClass(value = "") {
  return String(value).toLowerCase().replaceAll(" ", "-");
}

function renderOverview(rows) {
  const totals = {
    Published: rows.filter((row) => row.publishStatus === "Published").length,
    Draft: rows.filter((row) => row.publishStatus === "Draft").length,
    Modified: rows.filter((row) => row.changeType === "Modified").length,
    "Missing Arabic": rows.filter((row) => row.missingArabic).length,
    "Missing Images": rows.filter((row) => row.missingImages).length,
  };
  const max = Math.max(1, ...Object.values(totals));
  els.statusChart.innerHTML = Object.entries(totals).map(([label, value]) => `
    <div class="bar-row">
      <span>${escapeHtml(label)}</span>
      <div><i style="width:${Math.max(3, (value / max) * 100)}%"></i></div>
      <strong>${value}</strong>
    </div>
  `).join("");

  renderMiniList(els.recentChanges, [...rows].sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt))).slice(0, 8));
  renderMiniList(els.latestCreated, [...rows].sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt))).slice(0, 8), "createdAt");
  renderMiniList(els.removedLog, rows.filter((row) => row.changeType === "Removed").slice(0, 8));
}

function renderMiniList(container, rows, dateField = "updatedAt") {
  if (!rows.length) {
    container.innerHTML = `<div class="empty">No records.</div>`;
    return;
  }
  container.innerHTML = rows.map((row) => `
    <button type="button" data-open="${escapeHtml(row.documentId)}" class="mini-item">
      <strong>${escapeHtml(row.assetName || row.deviceNameEn || "-")}</strong>
      <span>${escapeHtml(row.brandName || "-")} · ${escapeHtml(row.publishStatus || "-")} · ${formatDate(row[dateField])}</span>
    </button>
  `).join("");
}

function renderCompare(rows) {
  const changed = rows.filter((row) => ["Modified", "Removed", "New"].includes(row.changeType)).slice(0, 80);
  if (!changed.length) {
    els.compareList.innerHTML = `<div class="empty">No changes in the latest snapshot.</div>`;
    return;
  }
  els.compareList.innerHTML = changed.map((row) => `
    <button type="button" class="mini-item" data-compare="${escapeHtml(row.documentId)}">
      <strong>${escapeHtml(row.assetName || "-")}</strong>
      <span>${escapeHtml(row.changeType)} · ${escapeHtml(row.publishStatus || "-")}</span>
    </button>
  `).join("");
}

function renderCompareDetails(row) {
  if (!row?.changeDetails?.length) {
    els.compareDetails.className = "diff-list empty";
    els.compareDetails.textContent = "No field-level differences captured for this device.";
    return;
  }
  els.compareDetails.className = "diff-list";
  els.compareDetails.innerHTML = row.changeDetails.map((change) => `
    <div class="diff-item">
      <h4>${escapeHtml(change.field)}</h4>
      <div class="diff-cols">
        <pre class="before">${escapeHtml(listText(change.before))}</pre>
        <pre class="after">${escapeHtml(listText(change.after))}</pre>
      </div>
    </div>
  `).join("");
}

function openDrawer(row) {
  state.selected = row;
  els.drawerTitle.textContent = row.assetName || "Device Details";
  els.drawerSubtitle.textContent = `${row.brandName || "-"} · ${row.deviceType || "-"} · ${row.publishStatus || "-"}`;
  els.drawerContent.innerHTML = `
    <section class="detail-grid">
      ${detail("Device Name EN", row.deviceNameEn)}
      ${detail("Device Name AR", row.deviceNameAr, "rtl")}
      ${detail("Device Group", row.deviceGroupId)}
      ${detail("Brand", row.brandName)}
      ${detail("Category", row.deviceType)}
      ${detail("Purchase Types", listText(row.purchaseTypes))}
      ${detail("Trouble Tickets", listText(row.troubleTickets))}
      ${detail("Web URL", row.webDetailUrl)}
      ${detail("Created", formatDate(row.createdAt))}
      ${detail("Updated", formatDate(row.updatedAt))}
      ${detail("Published", formatDate(row.publishedAt))}
      ${detail("Mandatory Fields", row.missingMandatory?.length ? row.missingMandatory.join(", ") : "Complete")}
    </section>
    <section class="detail-section">
      <h3>Media / Images</h3>
      ${(row.mediaUrls || []).length ? `<ul>${row.mediaUrls.map((url) => `<li>${url.startsWith("http") ? `<a href="${escapeHtml(url)}" target="_blank" rel="noreferrer">${escapeHtml(url)}</a>` : escapeHtml(url)}</li>`).join("")}</ul>` : "<p>No image references returned by the source API.</p>"}
    </section>
    <section class="detail-section">
      <h3>English and Arabic CMS Fields</h3>
      <div class="locale-raw-grid">
        <article>
          <h4>English</h4>
          <pre>${escapeHtml(row.rawJsonEn || "{}")}</pre>
        </article>
        <article>
          <h4>Arabic</h4>
          <pre class="rtl">${escapeHtml(row.rawJsonAr || "{}")}</pre>
        </article>
      </div>
    </section>
  `;
  els.drawer.hidden = false;
}

function openLabDrawer(row) {
  els.drawerTitle.textContent = row.titleEn || row.titleAr || "Labs Details";
  els.drawerSubtitle.textContent = `${row.collectionUid || "-"} · ${row.publishStatus || "-"} · ${row.changeType || "-"}`;
  els.drawerContent.innerHTML = `
    <section class="detail-grid">
      ${detail("Title EN", row.titleEn)}
      ${detail("Title AR", row.titleAr, "rtl")}
      ${detail("Collection UID", row.collectionUid)}
      ${detail("Publishing Status", row.publishStatus)}
      ${detail("Change Status", row.changeType)}
      ${detail("Created", formatDate(row.createdAt))}
      ${detail("Updated", formatDate(row.updatedAt))}
      ${detail("Published", formatDate(row.publishedAt))}
      ${detail("Field Count EN", row.fieldCountEn)}
      ${detail("Field Count AR", row.fieldCountAr)}
    </section>
    <section class="detail-section">
      <h3>Media / Images</h3>
      ${(row.mediaUrls || []).length ? `<ul>${row.mediaUrls.map((url) => `<li>${url.startsWith("http") ? `<a href="${escapeHtml(url)}" target="_blank" rel="noreferrer">${escapeHtml(url)}</a>` : escapeHtml(url)}</li>`).join("")}</ul>` : "<p>No image references returned by the source API.</p>"}
    </section>
    <section class="detail-section">
      <h3>English and Arabic CMS Fields</h3>
      <div class="locale-raw-grid">
        <article>
          <h4>English</h4>
          <pre>${escapeHtml(row.rawJsonEn || "{}")}</pre>
        </article>
        <article>
          <h4>Arabic</h4>
          <pre class="rtl">${escapeHtml(row.rawJsonAr || "{}")}</pre>
        </article>
      </div>
    </section>
  `;
  els.drawer.hidden = false;
}

function detail(label, value, className = "") {
  return `<div><span>${escapeHtml(label)}</span><strong class="${className}">${escapeHtml(value || "-")}</strong></div>`;
}

function setFilter(key, value) {
  state.filters[key] = value;
  state.page = 1;
  loadDevices();
}

function setLabsFilter(key, value) {
  labsState.filters[key] = value;
  labsState.page = 1;
  loadLabs();
}

function download(format) {
  const params = queryParams();
  params.delete("page");
  params.delete("pageSize");
  window.location.href = `/api/download/${format}?${params}`;
}

function activateSection(section) {
  const target = section === "devices" ? "assets" : section;
  if (!document.getElementById(target)) return;
  document.querySelectorAll("[data-section]").forEach((button) => button.classList.toggle("active", button.dataset.section === target));
  document.querySelectorAll(".section").forEach((item) => item.classList.toggle("active-section", item.id === target));
  if (location.hash.replace("#", "") !== section) history.replaceState(null, "", `#${section}`);
}

document.addEventListener("click", (event) => {
  const openId = event.target.closest("[data-open]")?.dataset.open;
  if (openId) {
    const row = state.allRows.find((item) => item.documentId === openId) || state.rows.find((item) => item.documentId === openId);
    if (row) openDrawer(row);
  }
  const openLabId = event.target.closest("[data-open-lab]")?.dataset.openLab;
  if (openLabId) {
    const row = labsState.allRows.find((item) => item.documentId === openLabId) || labsState.rows.find((item) => item.documentId === openLabId);
    if (row) openLabDrawer(row);
  }
  const compareId = event.target.closest("[data-compare]")?.dataset.compare;
  if (compareId) renderCompareDetails(state.allRows.find((item) => item.documentId === compareId));
  const section = event.target.closest("[data-section]")?.dataset.section;
  if (section) {
    activateSection(section === "assets" ? "devices" : section);
  }
});

els.fetchButton.addEventListener("click", refreshCms);
els.fetchLabsButton.addEventListener("click", refreshLabs);
els.adminLoginForm.addEventListener("submit", connectAdmin);
els.adminLogoutButton.addEventListener("click", disconnectAdmin);
els.refreshStatus.addEventListener("click", () => loadStatus().then(loadDevices).catch((error) => showToast(error.message)));
els.closeDrawer.addEventListener("click", () => { els.drawer.hidden = true; });
els.drawer.addEventListener("click", (event) => { if (event.target === els.drawer) els.drawer.hidden = true; });
els.searchInput.addEventListener("input", (event) => setFilter("search", event.target.value));
els.brandFilter.addEventListener("change", (event) => setFilter("brand", event.target.value));
els.categoryFilter.addEventListener("change", (event) => setFilter("category", event.target.value));
els.groupFilter.addEventListener("change", (event) => setFilter("group", event.target.value));
els.publishFilter.addEventListener("change", (event) => setFilter("publishStatus", event.target.value));
els.changeFilter.addEventListener("change", (event) => setFilter("changeType", event.target.value));
els.labsSearchInput.addEventListener("input", (event) => setLabsFilter("search", event.target.value));
els.labsPublishFilter.addEventListener("change", (event) => setLabsFilter("publishStatus", event.target.value));
els.labsChangeFilter.addEventListener("change", (event) => setLabsFilter("changeType", event.target.value));
els.sortSelect.addEventListener("change", (event) => setFilter("sort", event.target.value));
els.createdFrom.addEventListener("change", (event) => setFilter("createdFrom", event.target.value));
els.createdTo.addEventListener("change", (event) => setFilter("createdTo", event.target.value));
els.modifiedFrom.addEventListener("change", (event) => setFilter("modifiedFrom", event.target.value));
els.modifiedTo.addEventListener("change", (event) => setFilter("modifiedTo", event.target.value));
els.clearFilters.addEventListener("click", () => {
  Object.assign(state.filters, { search: "", brand: "all", category: "all", group: "all", publishStatus: "all", changeType: "all", sort: "updatedAt:desc", createdFrom: "", createdTo: "", modifiedFrom: "", modifiedTo: "" });
  els.searchInput.value = "";
  els.createdFrom.value = "";
  els.createdTo.value = "";
  els.modifiedFrom.value = "";
  els.modifiedTo.value = "";
  loadDevices();
});
els.prevPage.addEventListener("click", () => { if (state.page > 1) { state.page -= 1; loadDevices(); } });
els.nextPage.addEventListener("click", () => { state.page += 1; loadDevices(); });
els.labsPrevPage.addEventListener("click", () => { if (labsState.page > 1) { labsState.page -= 1; loadLabs(); } });
els.labsNextPage.addEventListener("click", () => { labsState.page += 1; loadLabs(); });
for (const id of ["exportXlsx", "exportXlsx2"]) $(id).addEventListener("click", () => download("xlsx"));
for (const id of ["exportCsv", "exportCsv2"]) $(id).addEventListener("click", () => download("csv"));
for (const id of ["exportPdf", "exportPdf2"]) $(id).addEventListener("click", () => download("pdf"));
window.addEventListener("hashchange", () => activateSection(location.hash.replace("#", "") || "overview"));
activateSection(location.hash.replace("#", "") || "overview");

loadStatus()
  .then(loadDevices)
  .then(loadLabsStatus)
  .then(loadLabs)
  .then(() => activateSection(location.hash.replace("#", "") || "overview"))
  .catch((error) => showToast(error.message));

setInterval(() => {
  refreshCms();
}, 120000);
