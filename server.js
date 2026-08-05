import crypto from "node:crypto";
import fs from "node:fs/promises";
import https from "node:https";
import path from "node:path";
import { fileURLToPath } from "node:url";
import express from "express";
import ExcelJS from "exceljs";
import PDFDocument from "pdfkit";

if (process.env.DASHBOARD_APP !== "devices") {
  await import("./labels-server.js");
} else {
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DATA_DIR = path.join(__dirname, "data");
const EXPORT_DIR = path.join(__dirname, "outputs", "dashboard-downloads");
const STATE_PATH = path.join(DATA_DIR, "device-state.json");
const LABS_STATE_PATH = path.join(DATA_DIR, "labs-state.json");
const DEVICE_API = "https://content.stc.com.kw/api/stc-device-groups";
const ADMIN_DEVICE_UID = "api::stc-device-group.stc-device-group";
const DEFAULT_LABS_ADMIN_UID = process.env.STRAPI_LABS_UID || "api::stc-label.stc-label";
const STRAPI_ADMIN_BASE = "https://content.stc.com.kw/admin/content-manager/collection-types";
const ENV_ADMIN_TOKEN = process.env.STRAPI_ADMIN_TOKEN || "";
let runtimeAdminToken = "";
const PAGE_SIZE = 100;
const MANDATORY_FIELDS = ["assetName", "deviceGroupId", "deviceName", "brandName", "deviceType", "purchaseType", "gridImage"];

const app = express();
app.use(express.json({ limit: "10mb" }));
app.use(express.static(path.join(__dirname, "public")));

function adminToken() {
  return runtimeAdminToken || ENV_ADMIN_TOKEN;
}

function isAdminMode() {
  return Boolean(adminToken());
}

function nowIso() {
  return new Date().toISOString();
}

function todayKuwaitDate() {
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Kuwait",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(new Date());
}

function asText(value) {
  if (value == null) return "";
  if (typeof value === "string") return value.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  return JSON.stringify(value);
}

function hash(value, length = 24) {
  return crypto.createHash("sha256").update(asText(value)).digest("hex").slice(0, length);
}

function stableJson(value) {
  if (Array.isArray(value)) return `[${value.map(stableJson).join(",")}]`;
  if (value && typeof value === "object") {
    return `{${Object.keys(value).sort().map((key) => `${JSON.stringify(key)}:${stableJson(value[key])}`).join(",")}}`;
  }
  return JSON.stringify(value);
}

function publicUrl(url) {
  if (!url) return "";
  if (/^https?:\/\//i.test(url)) return url;
  return `https://www.stc.com.kw${url.startsWith("/") ? "" : "/"}${url}`;
}

function collectMedia(value, output = []) {
  if (!value) return output;
  if (Array.isArray(value)) {
    for (const item of value) collectMedia(item, output);
    return output;
  }
  if (typeof value === "object") {
    if (value.url || value.mime || value.ext || value.provider_metadata) {
      output.push({
        id: value.id ?? "",
        name: value.name ?? value.alternativeText ?? value.caption ?? "",
        url: publicUrl(value.url),
        mime: value.mime ?? "",
      });
    } else if (value.id && Object.keys(value).length <= 2) {
      output.push({ id: value.id, name: "", url: "", mime: "" });
    }
    for (const nested of Object.values(value)) collectMedia(nested, output);
  }
  return output;
}

function textValues(value, result = []) {
  if (value == null) return result;
  if (typeof value === "string") {
    if (value.trim()) result.push(value);
    return result;
  }
  if (Array.isArray(value)) {
    for (const item of value) textValues(item, result);
    return result;
  }
  if (typeof value === "object") {
    for (const nested of Object.values(value)) textValues(nested, result);
  }
  return result;
}

function flatten(value, prefix = "", output = {}) {
  if (value == null) {
    output[prefix] = "";
    return output;
  }
  if (Array.isArray(value)) {
    output[prefix] = value.map((item) => (typeof item === "object" ? stableJson(item) : asText(item))).join("\n");
    return output;
  }
  if (typeof value === "object") {
    for (const [key, nested] of Object.entries(value)) {
      const next = prefix ? `${prefix}.${key}` : key;
      if (nested && typeof nested === "object" && !Array.isArray(nested)) flatten(nested, next, output);
      else output[next] = Array.isArray(nested) ? nested.map((item) => (typeof item === "object" ? stableJson(item) : asText(item))).join("\n") : asText(nested);
    }
    return output;
  }
  output[prefix] = asText(value);
  return output;
}

function publishingStatus(entry) {
  if (!entry) return "Missing";
  if (!entry.publishedAt) return "Draft";
  if (entry.updatedAt && new Date(entry.updatedAt) > new Date(entry.publishedAt)) return "Modified After Publish";
  return "Published";
}

function mandatoryMissing(entry, media) {
  const missing = [];
  for (const field of MANDATORY_FIELDS) {
    if (field === "purchaseType" && !(entry?.purchaseType ?? []).length) missing.push(field);
    else if (field === "gridImage" && !media.length) missing.push(field);
    else if (field !== "purchaseType" && field !== "gridImage" && !asText(entry?.[field]).trim()) missing.push(field);
  }
  return missing;
}

function mergeByDocumentId(enEntries, arEntries) {
  const byId = new Map();
  for (const entry of enEntries) byId.set(entry.documentId, { documentId: entry.documentId, en: entry, ar: null });
  for (const entry of arEntries) {
    const existing = byId.get(entry.documentId) ?? { documentId: entry.documentId, en: null, ar: null };
    existing.ar = entry;
    byId.set(entry.documentId, existing);
  }
  return [...byId.values()].map(({ documentId, en, ar }) => {
    const base = en ?? ar ?? {};
    const media = [...collectMedia(en), ...collectMedia(ar)].filter((item, index, arr) => (
      index === arr.findIndex((candidate) => `${candidate.id}|${candidate.url}|${candidate.name}` === `${item.id}|${item.url}|${item.name}`)
    ));
    const missingMandatory = mandatoryMissing(en ?? ar, media);
    const enText = textValues(en).join(" ");
    const arText = textValues(ar).join(" ");
    const signature = hash(stableJson({ en, ar }));
    return {
      id: documentId,
      documentId,
      entryIdEn: en?.id ?? "",
      entryIdAr: ar?.id ?? "",
      assetName: base.assetName ?? "",
      deviceGroupId: base.deviceGroupId ?? "",
      deviceNameEn: en?.deviceName ?? base.deviceName ?? "",
      deviceNameAr: ar?.deviceName ?? "",
      brandName: base.brandName ?? "",
      deviceType: base.deviceType ?? "",
      category: base.deviceType ?? "",
      webDetailUrl: base.webDetailUrl ?? "",
      preOrder: Boolean(base.preOrder),
      purchaseTypes: (base.purchaseType ?? []).map((item) => item.journeyType).filter(Boolean),
      troubleTickets: (base.purchaseType ?? []).map((item) => item.troubleTicket).filter(Boolean),
      colors: (base.color ?? base.variant?.color ?? []).map((item) => `${item.color ?? item.hex ?? ""}${item.colorName ? ` ${item.colorName}` : ""}`).filter(Boolean),
      capacities: (base.capacity ?? base.variant?.capacity ?? []).map((item) => asText(item.capacity ?? item.value ?? item.name)).filter(Boolean),
      createdAt: base.createdAt ?? "",
      updatedAt: base.updatedAt ?? "",
      publishedAt: base.publishedAt ?? "",
      publishStatus: publishingStatus(en ?? ar),
      hasArabic: Boolean(ar),
      translationStatus: ar ? "Available" : "Missing Arabic",
      missingArabic: !ar,
      missingImages: media.length === 0,
      missingMandatory,
      mandatoryStatus: missingMandatory.length ? "Missing fields" : "Complete",
      media,
      mediaUrls: media.map((item) => item.url || `media-id:${item.id}`).filter(Boolean),
      en,
      ar,
      flatEn: flatten(en ?? {}),
      flatAr: flatten(ar ?? {}),
      rawJsonEn: JSON.stringify(en ?? {}, null, 2),
      rawJsonAr: JSON.stringify(ar ?? {}, null, 2),
      searchable: [documentId, base.assetName, en?.deviceName, ar?.deviceName, base.brandName, base.deviceType, base.webDetailUrl, enText, arText].join(" ").toLowerCase(),
      signature,
    };
  }).sort((a, b) => a.assetName.localeCompare(b.assetName));
}

async function ensureDirs() {
  await fs.mkdir(DATA_DIR, { recursive: true });
  await fs.mkdir(EXPORT_DIR, { recursive: true });
}

async function loadState() {
  await ensureDirs();
  try {
    return JSON.parse(await fs.readFile(STATE_PATH, "utf8"));
  } catch {
    return { initializedAt: null, devices: {}, snapshots: [] };
  }
}

async function saveState(state) {
  await ensureDirs();
  await fs.writeFile(STATE_PATH, JSON.stringify(state, null, 2), "utf8");
}

async function loadLabsState() {
  await ensureDirs();
  try {
    return JSON.parse(await fs.readFile(LABS_STATE_PATH, "utf8"));
  } catch {
    return { initializedAt: null, labs: {}, snapshots: [] };
  }
}

async function saveLabsState(state) {
  await ensureDirs();
  await fs.writeFile(LABS_STATE_PATH, JSON.stringify(state, null, 2), "utf8");
}

function cleanCollectionUid(uid) {
  return asText(uid || DEFAULT_LABS_ADMIN_UID).trim() || DEFAULT_LABS_ADMIN_UID;
}

function adminCollectionUrl(uid, locale = "en") {
  return `${STRAPI_ADMIN_BASE}/${uid}?plugins%5Bi18n%5D%5Blocale%5D=${encodeURIComponent(locale)}`;
}

async function fetchPublicLocale(locale) {
  const all = [];
  let page = 1;
  let pageCount = 1;
  do {
    const url = new URL(DEVICE_API);
    url.searchParams.set("pagination[page]", String(page));
    url.searchParams.set("pagination[pageSize]", String(PAGE_SIZE));
    url.searchParams.set("populate", "*");
    url.searchParams.set("locale", locale);
    url.searchParams.set("publicationState", "preview");
    const json = await getJson(url);
    all.push(...(json.data ?? []));
    pageCount = json.meta?.pagination?.pageCount ?? 1;
    page += 1;
  } while (page <= pageCount);
  return all;
}

async function fetchAdminLocale(locale) {
  const token = adminToken();
  if (!token) throw new Error("Admin credentials are not connected.");
  const all = [];
  let page = 1;
  let pageCount = 1;
  do {
    const url = new URL(`https://content.stc.com.kw/content-manager/collection-types/${ADMIN_DEVICE_UID}`);
    url.searchParams.set("page", String(page));
    url.searchParams.set("pageSize", String(PAGE_SIZE));
    url.searchParams.set("sort", "assetName:ASC");
    url.searchParams.set("plugins[i18n][locale]", locale);
    const json = await getJson(url, { Authorization: `Bearer ${token}` });
    const entries = json.results ?? json.data?.results ?? json.data ?? [];
    const pagination = json.pagination ?? json.meta?.pagination ?? {};
    all.push(...entries);
    pageCount = pagination.pageCount ?? Math.ceil((pagination.total ?? all.length) / PAGE_SIZE) ?? 1;
    page += 1;
  } while (page <= pageCount);

  const detailed = [];
  for (const entry of all) {
    const documentId = entry.documentId ?? entry.id;
    if (!documentId) {
      detailed.push(entry);
      continue;
    }
    const detailUrl = new URL(`https://content.stc.com.kw/content-manager/collection-types/${ADMIN_DEVICE_UID}/${documentId}`);
    detailUrl.searchParams.set("plugins[i18n][locale]", locale);
    const detail = await getJson(detailUrl, { Authorization: `Bearer ${token}` });
    detailed.push(detail.data ?? detail);
  }
  return detailed;
}

async function fetchAdminCollectionLocale(uid, locale) {
  const token = adminToken();
  if (!token) throw new Error("Connect Strapi admin credentials before fetching this collection.");
  const all = [];
  let page = 1;
  let pageCount = 1;
  do {
    const url = new URL(`https://content.stc.com.kw/content-manager/collection-types/${uid}`);
    url.searchParams.set("page", String(page));
    url.searchParams.set("pageSize", String(PAGE_SIZE));
    url.searchParams.set("sort", "updatedAt:DESC");
    url.searchParams.set("plugins[i18n][locale]", locale);
    const json = await getJson(url, { Authorization: `Bearer ${token}` });
    const entries = json.results ?? json.data?.results ?? json.data ?? [];
    const pagination = json.pagination ?? json.meta?.pagination ?? {};
    all.push(...entries);
    pageCount = pagination.pageCount ?? Math.ceil((pagination.total ?? all.length) / PAGE_SIZE) ?? 1;
    page += 1;
  } while (page <= pageCount);

  const detailed = [];
  for (const entry of all) {
    const documentId = entry.documentId ?? entry.id;
    if (!documentId) {
      detailed.push(entry);
      continue;
    }
    const detailUrl = new URL(`https://content.stc.com.kw/content-manager/collection-types/${uid}/${documentId}`);
    detailUrl.searchParams.set("plugins[i18n][locale]", locale);
    const detail = await getJson(detailUrl, { Authorization: `Bearer ${token}` });
    detailed.push(detail.data ?? detail);
  }
  return detailed;
}

async function fetchDeviceSource() {
  if (isAdminMode()) {
    const [en, ar] = await Promise.all([fetchAdminLocale("en"), fetchAdminLocale("ar")]);
    return { en, ar, source: "authenticated Strapi admin Content Manager API" };
  }
  const [en, ar] = await Promise.all([fetchPublicLocale("en"), fetchPublicLocale("ar")]);
  return { en, ar, source: "public Strapi content API" };
}

function displayName(entry) {
  return entry?.assetName
    || entry?.labelName
    || entry?.labName
    || entry?.name
    || entry?.title
    || entry?.key
    || entry?.slug
    || entry?.documentId
    || entry?.id
    || "";
}

function mergeGenericByDocumentId(enEntries, arEntries, collectionUid) {
  const byId = new Map();
  for (const entry of enEntries) byId.set(entry.documentId ?? entry.id, { documentId: entry.documentId ?? String(entry.id), en: entry, ar: null });
  for (const entry of arEntries) {
    const id = entry.documentId ?? String(entry.id);
    const existing = byId.get(id) ?? { documentId: id, en: null, ar: null };
    existing.ar = entry;
    byId.set(id, existing);
  }
  return [...byId.values()].map(({ documentId, en, ar }) => {
    const base = en ?? ar ?? {};
    const media = collectMedia(base);
    const rawEn = en ? JSON.stringify(en, null, 2) : "";
    const rawAr = ar ? JSON.stringify(ar, null, 2) : "";
    return {
      documentId,
      collectionUid,
      titleEn: displayName(en) || displayName(base),
      titleAr: displayName(ar),
      createdAt: base.createdAt ?? en?.createdAt ?? ar?.createdAt ?? "",
      updatedAt: base.updatedAt ?? en?.updatedAt ?? ar?.updatedAt ?? "",
      publishedAt: base.publishedAt ?? en?.publishedAt ?? ar?.publishedAt ?? "",
      publishStatus: publishingStatus(base),
      hasArabic: Boolean(ar),
      translationStatus: ar ? "Available" : "Missing Arabic",
      missingArabic: !ar,
      mediaUrls: [...new Set(media.map((item) => item.url).filter(Boolean))],
      fieldCountEn: en ? Object.keys(flatten(en)).length : 0,
      fieldCountAr: ar ? Object.keys(flatten(ar)).length : 0,
      rawJsonEn: rawEn,
      rawJsonAr: rawAr,
      signature: hash(stableJson({ en, ar })),
    };
  }).sort((a, b) => asText(a.titleEn || a.titleAr).localeCompare(asText(b.titleEn || b.titleAr)));
}

function compactLab(row, includeRaw = false) {
  const copy = {
    documentId: row.documentId,
    collectionUid: row.collectionUid,
    titleEn: row.titleEn,
    titleAr: row.titleAr,
    createdAt: row.createdAt,
    updatedAt: row.updatedAt,
    publishedAt: row.publishedAt,
    publishStatus: row.publishStatus,
    changeType: row.changeType,
    translationStatus: row.translationStatus,
    missingArabic: row.missingArabic,
    mediaUrls: row.mediaUrls,
    fieldCountEn: row.fieldCountEn,
    fieldCountAr: row.fieldCountAr,
    signature: row.signature,
  };
  if (includeRaw) {
    copy.rawJsonEn = row.rawJsonEn;
    copy.rawJsonAr = row.rawJsonAr;
  }
  return copy;
}

function comparableSnapshot(state, source) {
  return (state.snapshots ?? []).find((snapshot) => snapshot.source === source) ?? null;
}

function wasCreatedAfterSnapshot(createdAt, snapshot) {
  if (!createdAt || !snapshot?.fetchedAt) return false;
  const created = new Date(createdAt);
  const fetched = new Date(snapshot.fetchedAt);
  return !Number.isNaN(created.getTime()) && !Number.isNaN(fetched.getTime()) && created > fetched;
}

function compareLabs(state, currentLabs, collectionUid) {
  const seenAt = nowIso();
  const source = "authenticated Strapi admin Content Manager API";
  const baseline = comparableSnapshot(state, source);
  const previous = state.labs ?? {};
  const currentIds = new Set(currentLabs.map((item) => item.documentId));
  const rows = [];
  const changes = { new: [], modified: [], removed: [] };
  for (const lab of currentLabs) {
    const old = previous[lab.documentId];
    let changeType = "Unchanged";
    if (!old) changeType = state.initializedAt && (!baseline || wasCreatedAfterSnapshot(lab.createdAt, baseline)) ? "New" : "Unchanged";
    else if (old.snapshot?.changeType === "New" && old.signature === lab.signature) changeType = "New";
    else if (old.signature !== lab.signature) changeType = "Modified";
    const row = { ...lab, changeType, firstSeenAt: old?.firstSeenAt ?? seenAt, lastSeenAt: seenAt };
    rows.push(row);
    if (changeType === "New") changes.new.push(compactLab(row));
    if (changeType === "Modified") changes.modified.push(compactLab(row));
  }
  for (const [id, old] of Object.entries(previous)) {
    if (currentIds.has(id)) continue;
    const removed = { ...old.snapshot, documentId: id, changeType: "Removed", lastSeenAt: seenAt };
    rows.push(removed);
    changes.removed.push(compactLab(removed));
  }
  state.initializedAt ||= seenAt;
  state.collectionUid = collectionUid;
  state.labs = {};
  for (const row of rows.filter((row) => row.changeType !== "Removed")) {
    state.labs[row.documentId] = {
      firstSeenAt: row.firstSeenAt,
      lastSeenAt: row.lastSeenAt,
      signature: row.signature,
      snapshot: compactLab(row, true),
    };
  }
  const snapshot = {
    id: `LAB-${hash(seenAt, 10)}`,
    fetchedAt: seenAt,
    source,
    collectionUid,
    sourceTotals: { current: currentLabs.length },
    counts: summarizeLabs(rows),
    changes,
  };
  state.lastSnapshot = snapshot;
  state.snapshots = [snapshot, ...(state.snapshots ?? [])].slice(0, 20);
  return { rows, snapshot };
}

function getJson(url, headers = {}) {
  return new Promise((resolve, reject) => {
    const request = https.get(url, { rejectUnauthorized: false, headers: { Accept: "application/json", ...headers } }, (response) => {
      let body = "";
      response.setEncoding("utf8");
      response.on("data", (chunk) => { body += chunk; });
      response.on("end", () => {
        if ((response.statusCode ?? 500) >= 400) {
          reject(new Error(`STC device API returned ${response.statusCode}: ${body.slice(0, 300)}`));
          return;
        }
        try {
          resolve(JSON.parse(body.replace(/^\uFEFF/, "")));
        } catch (error) {
          reject(error);
        }
      });
    });
    request.on("error", reject);
    request.setTimeout(60000, () => request.destroy(new Error("STC device API request timed out")));
  });
}

function postJson(url, payload, headers = {}) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify(payload);
    const request = https.request(url, {
      method: "POST",
      rejectUnauthorized: false,
      headers: {
        Accept: "application/json",
        "Content-Type": "application/json",
        "Content-Length": Buffer.byteLength(body),
        ...headers,
      },
    }, (response) => {
      let responseBody = "";
      response.setEncoding("utf8");
      response.on("data", (chunk) => { responseBody += chunk; });
      response.on("end", () => {
        let json = {};
        try {
          json = responseBody ? JSON.parse(responseBody.replace(/^\uFEFF/, "")) : {};
        } catch {
          json = { raw: responseBody };
        }
        if ((response.statusCode ?? 500) >= 400) {
          const message = json.error?.message || json.message || responseBody.slice(0, 300) || "Login failed";
          reject(new Error(`Strapi admin login returned ${response.statusCode}: ${message}`));
          return;
        }
        resolve(json);
      });
    });
    request.on("error", reject);
    request.setTimeout(60000, () => request.destroy(new Error("Strapi admin login request timed out")));
    request.write(body);
    request.end();
  });
}

function extractAdminJwt(json) {
  return json?.data?.token
    || json?.data?.jwt
    || json?.data?.accessToken
    || json?.token
    || json?.jwt
    || json?.accessToken
    || "";
}

async function loginAdmin(email, password) {
  const loginJson = await postJson("https://content.stc.com.kw/admin/login", { email, password });
  const token = extractAdminJwt(loginJson);
  if (!token) throw new Error("Login succeeded, but Strapi did not return an admin token.");

  const testUrl = new URL(`https://content.stc.com.kw/content-manager/collection-types/${ADMIN_DEVICE_UID}`);
  testUrl.searchParams.set("page", "1");
  testUrl.searchParams.set("pageSize", "1");
  testUrl.searchParams.set("sort", "assetName:ASC");
  testUrl.searchParams.set("plugins[i18n][locale]", "en");
  const testJson = await getJson(testUrl, { Authorization: `Bearer ${token}` });
  runtimeAdminToken = token;
  const pagination = testJson.pagination ?? testJson.meta?.pagination ?? {};
  return { total: pagination.total ?? null };
}

function compareAndUpdate(state, currentDevices, source) {
  const seenAt = nowIso();
  const baseline = comparableSnapshot(state, source);
  const previous = state.devices ?? {};
  const currentIds = new Set(currentDevices.map((device) => device.documentId));
  const rows = [];
  const changes = { new: [], modified: [], removed: [], unchanged: [] };

  for (const device of currentDevices) {
    const old = previous[device.documentId];
    let changeType = "Unchanged";
    let changeDetails = [];
    if (!old) changeType = state.initializedAt && (!baseline || wasCreatedAfterSnapshot(device.createdAt, baseline)) ? "New" : "Unchanged";
    else if (old.snapshot?.changeType === "New" && old.signature === device.signature) changeType = "New";
    else if (old.signature !== device.signature) {
      changeType = "Modified";
      changeDetails = diffDevice(old.snapshot, device);
    }
    const row = {
      ...device,
      changeType,
      firstSeenAt: old?.firstSeenAt ?? seenAt,
      lastSeenAt: seenAt,
      previousValues: old?.snapshot ?? null,
      changeDetails,
    };
    rows.push(row);
    changes[changeType === "New" ? "new" : changeType === "Modified" ? "modified" : "unchanged"].push(compactRow(row));
  }

  for (const [documentId, old] of Object.entries(previous)) {
    if (currentIds.has(documentId)) continue;
    const removed = {
      ...old.snapshot,
      documentId,
      changeType: "Removed",
      firstSeenAt: old.firstSeenAt,
      lastSeenAt: seenAt,
      publishStatus: old.snapshot?.publishStatus ?? "",
      signature: old.signature,
      previousValues: old.snapshot,
      changeDetails: [{ field: "Availability", before: "Present", after: "Removed from source" }],
    };
    rows.push(removed);
    changes.removed.push(compactRow(removed));
  }

  state.initializedAt ||= seenAt;
  state.devices = {};
  for (const row of rows.filter((row) => row.changeType !== "Removed")) {
    state.devices[row.documentId] = {
      firstSeenAt: row.firstSeenAt,
      lastSeenAt: row.lastSeenAt,
      signature: row.signature,
      snapshot: compactRow(row, true),
    };
  }

  const snapshot = {
    id: `SNAP-${hash(seenAt, 10)}`,
    fetchedAt: seenAt,
    source,
    sourceTotals: { current: currentDevices.length },
    counts: summarize(rows),
    changes,
  };
  state.lastSnapshot = snapshot;
  state.snapshots = [snapshot, ...(state.snapshots ?? [])].slice(0, 20);
  return { rows, snapshot };
}

function compactRow(row, includeRaw = false) {
  const copy = {
    documentId: row.documentId,
    assetName: row.assetName,
    deviceGroupId: row.deviceGroupId,
    deviceNameEn: row.deviceNameEn,
    deviceNameAr: row.deviceNameAr,
    brandName: row.brandName,
    deviceType: row.deviceType,
    webDetailUrl: row.webDetailUrl,
    preOrder: row.preOrder,
    purchaseTypes: row.purchaseTypes,
    troubleTickets: row.troubleTickets,
    colors: row.colors,
    capacities: row.capacities,
    createdAt: row.createdAt,
    updatedAt: row.updatedAt,
    publishedAt: row.publishedAt,
    publishStatus: row.publishStatus,
    changeType: row.changeType,
    translationStatus: row.translationStatus,
    missingArabic: row.missingArabic,
    missingImages: row.missingImages,
    missingMandatory: row.missingMandatory,
    mediaUrls: row.mediaUrls,
    firstSeenAt: row.firstSeenAt,
    lastSeenAt: row.lastSeenAt,
    changeDetails: row.changeDetails ?? [],
    signature: row.signature,
  };
  if (includeRaw) {
    copy.rawJsonEn = row.rawJsonEn;
    copy.rawJsonAr = row.rawJsonAr;
    copy.flatEn = row.flatEn;
    copy.flatAr = row.flatAr;
    copy.en = row.en;
    copy.ar = row.ar;
  }
  return copy;
}

function diffDevice(previous, current) {
  const fields = ["assetName", "deviceNameEn", "deviceNameAr", "brandName", "deviceType", "webDetailUrl", "purchaseTypes", "troubleTickets", "mediaUrls", "publishStatus", "translationStatus"];
  return fields
    .filter((field) => stableJson(previous?.[field]) !== stableJson(current?.[field]))
    .map((field) => ({ field, before: previous?.[field] ?? "", after: current?.[field] ?? "" }));
}

function activeRows(state) {
  return Object.values(state.devices ?? {}).map((item) => ({ ...item.snapshot, firstSeenAt: item.firstSeenAt, lastSeenAt: item.lastSeenAt }));
}

function summarize(rows) {
  const counts = rows.reduce((acc, row) => {
    const key = row.changeType || "Unchanged";
    acc[key] = (acc[key] ?? 0) + 1;
    return acc;
  }, {});
  return {
    total: rows.length,
    new: counts.New ?? 0,
    modified: counts.Modified ?? 0,
    removed: counts.Removed ?? 0,
    unchanged: counts.Unchanged ?? 0,
    draft: rows.filter((row) => row.publishStatus === "Draft").length,
    published: rows.filter((row) => row.publishStatus === "Published").length,
    modifiedAfterPublish: rows.filter((row) => row.publishStatus === "Modified After Publish").length,
    missingArabic: rows.filter((row) => row.missingArabic).length,
    missingImages: rows.filter((row) => row.missingImages).length,
    missingMandatory: rows.filter((row) => row.missingMandatory?.length).length,
    brands: new Set(rows.map((row) => row.brandName).filter(Boolean)).size,
    deviceGroups: new Set(rows.map((row) => row.deviceGroupId).filter(Boolean)).size,
  };
}

function activeLabs(state) {
  return Object.values(state.labs ?? {}).map((item) => ({ ...item.snapshot, firstSeenAt: item.firstSeenAt, lastSeenAt: item.lastSeenAt }));
}

function summarizeLabs(rows) {
  const counts = rows.reduce((acc, row) => {
    const key = row.changeType || "Unchanged";
    acc[key] = (acc[key] ?? 0) + 1;
    return acc;
  }, {});
  return {
    total: rows.length,
    new: counts.New ?? 0,
    modified: counts.Modified ?? 0,
    removed: counts.Removed ?? 0,
    unchanged: counts.Unchanged ?? 0,
    draft: rows.filter((row) => row.publishStatus === "Draft").length,
    published: rows.filter((row) => row.publishStatus === "Published").length,
    modifiedAfterPublish: rows.filter((row) => row.publishStatus === "Modified After Publish").length,
    missingArabic: rows.filter((row) => row.missingArabic).length,
    missingImages: rows.filter((row) => !(row.mediaUrls ?? []).length).length,
  };
}

function filterLabs(rows, query = {}) {
  const search = asText(query.search).toLowerCase().trim();
  return rows.filter((row) => {
    if (query.publishStatus && query.publishStatus !== "all") {
      const requestedStatus = asText(query.publishStatus);
      const isModifiedStatus = requestedStatus === "Modified"
        && (row.changeType === "Modified" || row.publishStatus === "Modified" || row.publishStatus === "Modified After Publish");
      if (!isModifiedStatus && row.publishStatus !== requestedStatus) return false;
    }
    if (query.changeType && query.changeType !== "all" && row.changeType !== query.changeType) return false;
    if (!search) return true;
    return [row.documentId, row.titleEn, row.titleAr, row.collectionUid, row.rawJsonEn, row.rawJsonAr]
      .some((value) => asText(value).toLowerCase().includes(search));
  });
}

function filterRows(rows, query = {}) {
  const search = asText(query.search).toLowerCase().trim();
  return rows.filter((row) => {
    if (query.brand && query.brand !== "all" && row.brandName !== query.brand) return false;
    if (query.category && query.category !== "all" && row.deviceType !== query.category) return false;
    if (query.group && query.group !== "all" && row.deviceGroupId !== query.group) return false;
    if (query.publishStatus && query.publishStatus !== "all") {
      const requestedStatus = asText(query.publishStatus);
      const isModifiedStatus = requestedStatus === "Modified"
        && (row.changeType === "Modified" || row.publishStatus === "Modified" || row.publishStatus === "Modified After Publish");
      if (!isModifiedStatus && row.publishStatus !== requestedStatus) return false;
    }
    if (query.changeType && query.changeType !== "all" && row.changeType !== query.changeType) return false;
    if (query.language && query.language !== "all" && (query.language === "ar" ? row.missingArabic : false)) return false;
    if (query.createdFrom && row.createdAt && row.createdAt.slice(0, 10) < query.createdFrom) return false;
    if (query.createdTo && row.createdAt && row.createdAt.slice(0, 10) > query.createdTo) return false;
    if (query.modifiedFrom && row.updatedAt && row.updatedAt.slice(0, 10) < query.modifiedFrom) return false;
    if (query.modifiedTo && row.updatedAt && row.updatedAt.slice(0, 10) > query.modifiedTo) return false;
    if (!search) return true;
    return [row.assetName, row.deviceGroupId, row.deviceNameEn, row.deviceNameAr, row.brandName, row.deviceType, row.webDetailUrl, row.documentId, row.rawJsonEn, row.rawJsonAr]
      .some((value) => asText(value).toLowerCase().includes(search));
  });
}

function sortRows(rows, sort = "updatedAt:desc") {
  const [field, direction] = sort.split(":");
  const factor = direction === "asc" ? 1 : -1;
  return [...rows].sort((a, b) => asText(a[field]).localeCompare(asText(b[field]), undefined, { numeric: true }) * factor);
}

function optionsFrom(rows) {
  const unique = (field) => [...new Set(rows.map((row) => row[field]).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  return {
    brands: unique("brandName"),
    categories: unique("deviceType"),
    groups: unique("deviceGroupId"),
    publishStatuses: unique("publishStatus"),
  };
}

function snapshotSummary(snapshot) {
  return {
    id: snapshot.id,
    fetchedAt: snapshot.fetchedAt,
    source: snapshot.source,
    sourceTotals: snapshot.sourceTotals,
    counts: snapshot.counts,
  };
}

const exportHeaders = [
  ["documentId", "Document ID"],
  ["assetName", "Asset Name"],
  ["deviceGroupId", "Device Group ID"],
  ["deviceNameEn", "Device Name EN"],
  ["deviceNameAr", "Device Name AR"],
  ["brandName", "Brand"],
  ["deviceType", "Category"],
  ["webDetailUrl", "Web Detail URL"],
  ["preOrder", "Pre Order"],
  ["purchaseTypes", "Purchase Types"],
  ["troubleTickets", "Trouble Tickets"],
  ["colors", "Colors"],
  ["capacities", "Capacities"],
  ["publishStatus", "Publishing Status"],
  ["changeType", "Change Status"],
  ["translationStatus", "Translation Status"],
  ["missingMandatory", "Missing Mandatory Fields"],
  ["mediaUrls", "Media/Image URLs"],
  ["createdAt", "Created At"],
  ["updatedAt", "Last Modified At"],
  ["publishedAt", "Published At"],
  ["rawJsonEn", "Full Raw JSON EN"],
  ["rawJsonAr", "Full Raw JSON AR"],
];

function exportValue(row, key) {
  const value = row[key];
  if (Array.isArray(value)) return value.join("\n");
  return asText(value);
}

async function buildExcel(rows) {
  await ensureDirs();
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "STC Device Asset Dashboard";
  workbook.created = new Date();
  workbook.modified = new Date();

  const summary = workbook.addWorksheet("Summary");
  const totals = summarize(rows);
  summary.addRows([
    ["Metric", "Value"],
    ["Total Devices", totals.total],
    ["New Devices", totals.new],
    ["Modified Devices", totals.modified],
    ["Removed Devices", totals.removed],
    ["Unchanged Devices", totals.unchanged],
    ["Draft Devices", totals.draft],
    ["Published Devices", totals.published],
    ["Modified After Publish", totals.modifiedAfterPublish],
    ["Missing Arabic Translation", totals.missingArabic],
    ["Missing Images", totals.missingImages],
    ["Missing Mandatory Fields", totals.missingMandatory],
  ]);
  styleSheet(summary);

  const devices = workbook.addWorksheet("Device Assets");
  devices.addRow(exportHeaders.map(([, label]) => label));
  for (const row of rows) devices.addRow(exportHeaders.map(([key]) => exportValue(row, key)));
  styleSheet(devices);
  devices.views = [{ state: "frozen", ySplit: 1 }];
  devices.autoFilter = { from: "A1", to: `${devices.getColumn(exportHeaders.length).letter}1` };

  const changes = workbook.addWorksheet("Change Details");
  changes.addRow(["Document ID", "Asset Name", "Change Type", "Field", "Before", "After"]);
  for (const row of rows) {
    for (const change of row.changeDetails ?? []) changes.addRow([row.documentId, row.assetName, row.changeType, change.field, asText(change.before), asText(change.after)]);
  }
  styleSheet(changes);

  const fileName = `stc-device-assets-${rows.length}-rows-${Date.now()}.xlsx`;
  const filePath = path.join(EXPORT_DIR, fileName);
  await workbook.xlsx.writeFile(filePath);
  return { fileName, filePath };
}

function styleSheet(sheet) {
  sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  sheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4F008C" } };
  sheet.getRow(1).alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  sheet.columns.forEach((column) => {
    column.width = Math.min(54, Math.max(14, ...column.values.map((value) => asText(value).length).slice(1, 80)) + 2);
    column.alignment = { vertical: "top", wrapText: true };
  });
}

function csvEscape(value) {
  const text = asText(value);
  return /[",\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

async function buildCsv(rows) {
  const header = exportHeaders.map(([, label]) => csvEscape(label)).join(",");
  const body = rows.map((row) => exportHeaders.map(([key]) => csvEscape(exportValue(row, key))).join(",")).join("\n");
  const fileName = `stc-device-assets-${rows.length}-rows-${Date.now()}.csv`;
  const filePath = path.join(EXPORT_DIR, fileName);
  await fs.writeFile(filePath, `${header}\n${body}`, "utf8");
  return { fileName, filePath };
}

async function buildPdf(rows) {
  const fileName = `stc-device-assets-${rows.length}-rows-${Date.now()}.pdf`;
  const filePath = path.join(EXPORT_DIR, fileName);
  const doc = new PDFDocument({ margin: 36, size: "A4", layout: "landscape" });
  const stream = await fs.open(filePath, "w");
  doc.pipe(stream.createWriteStream());
  const totals = summarize(rows);
  doc.fontSize(18).text("STC Device Asset Dashboard Export", { continued: false });
  doc.moveDown(0.5).fontSize(10).text(`Generated: ${new Date().toLocaleString()} | Devices: ${rows.length}`);
  doc.moveDown(0.5).text(`Published: ${totals.published} | Draft: ${totals.draft} | Modified: ${totals.modified} | Missing Arabic: ${totals.missingArabic} | Missing Images: ${totals.missingImages}`);
  doc.moveDown();
  const headers = ["Asset", "Brand", "Type", "Publish", "Change", "Arabic", "Updated"];
  doc.fontSize(8).font("Helvetica-Bold").text(headers.join("   "));
  doc.font("Helvetica");
  for (const row of rows.slice(0, 500)) {
    doc.text([row.assetName, row.brandName, row.deviceType, row.publishStatus, row.changeType, row.translationStatus, row.updatedAt].map((v) => asText(v).slice(0, 28)).join("   "));
    if (doc.y > 535) doc.addPage().fontSize(8);
  }
  doc.end();
  await new Promise((resolve) => doc.on("end", resolve));
  await stream.close();
  return { fileName, filePath };
}

app.get("/api/status", async (_req, res) => {
  const state = await loadState();
  const rows = activeRows(state);
  const adminMode = isAdminMode();
  res.json({
    initializedAt: state.initializedAt,
    lastSnapshot: state.lastSnapshot ? snapshotSummary(state.lastSnapshot) : null,
    snapshots: (state.snapshots ?? []).map(snapshotSummary),
    totals: summarize(rows),
    options: optionsFrom(rows),
    sourceMode: adminMode ? "admin" : "public",
    adminSessionActive: Boolean(runtimeAdminToken),
    envAdminTokenConfigured: Boolean(ENV_ADMIN_TOKEN),
    sourceNote: adminMode
      ? "Authenticated admin mode is enabled; refresh uses Strapi Content Manager data."
      : "Public API mode is enabled; connect Strapi admin credentials to include drafts and admin-only entries.",
  });
});

app.post("/api/admin-login", async (req, res) => {
  try {
    const email = String(req.body?.email ?? "").trim();
    const password = String(req.body?.password ?? "");
    if (!email || !password) return res.status(400).json({ error: "Enter both Strapi email and password." });
    const result = await loginAdmin(email, password);
    res.json({
      ok: true,
      sourceMode: "admin",
      adminSessionActive: true,
      adminTotal: result.total,
      message: "Strapi admin session connected.",
    });
  } catch (error) {
    runtimeAdminToken = "";
    res.status(401).json({ error: error.message });
  }
});

app.post("/api/admin-logout", (_req, res) => {
  runtimeAdminToken = "";
  res.json({
    ok: true,
    sourceMode: ENV_ADMIN_TOKEN ? "admin" : "public",
    adminSessionActive: false,
    message: ENV_ADMIN_TOKEN ? "Runtime credentials cleared. Environment admin token is still active." : "Admin session cleared. Public API mode is active.",
  });
});

app.get("/api/devices", async (req, res) => {
  const state = await loadState();
  let rows = activeRows(state);
  rows = sortRows(filterRows(rows, req.query), req.query.sort);
  const page = Math.max(1, Number(req.query.page || 1));
  const pageSize = Math.min(100, Math.max(10, Number(req.query.pageSize || 25)));
  res.json({
    rows: rows.slice((page - 1) * pageSize, page * pageSize),
    allRows: rows,
    total: rows.length,
    page,
    pageSize,
    totals: summarize(rows),
    options: optionsFrom(activeRows(state)),
    snapshots: (state.snapshots ?? []).map(snapshotSummary),
  });
});

app.get("/api/devices/:documentId", async (req, res) => {
  const state = await loadState();
  const row = activeRows(state).find((item) => item.documentId === req.params.documentId);
  if (!row) return res.status(404).json({ error: "Device not found" });
  res.json({ row });
});

app.post("/api/fetch", async (_req, res) => {
  try {
    const state = await loadState();
    const { en, ar, source } = await fetchDeviceSource();
    const current = mergeByDocumentId(en, ar);
    const { rows, snapshot } = compareAndUpdate(state, current, source);
    snapshot.source = source;
    await saveState(state);
    res.json({
      rows,
      totals: summarize(rows),
      snapshot: snapshotSummary(snapshot),
      snapshots: (state.snapshots ?? []).map(snapshotSummary),
      options: optionsFrom(rows),
      sourceTotals: { en: en.length, ar: ar.length, merged: current.length },
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/labs/status", async (_req, res) => {
  const state = await loadLabsState();
  const rows = activeLabs(state);
  res.json({
    collectionUid: state.collectionUid || DEFAULT_LABS_ADMIN_UID,
    adminUrlEn: adminCollectionUrl(state.collectionUid || DEFAULT_LABS_ADMIN_UID, "en"),
    adminUrlAr: adminCollectionUrl(state.collectionUid || DEFAULT_LABS_ADMIN_UID, "ar"),
    lastSnapshot: state.lastSnapshot ? snapshotSummary(state.lastSnapshot) : null,
    snapshots: (state.snapshots ?? []).map(snapshotSummary),
    totals: summarizeLabs(rows),
    sourceMode: isAdminMode() ? "admin" : "public",
  });
});

app.get("/api/labs", async (req, res) => {
  const state = await loadLabsState();
  let rows = activeLabs(state);
  rows = sortRows(filterLabs(rows, req.query), req.query.sort || "updatedAt:desc");
  const page = Math.max(1, Number(req.query.page || 1));
  const pageSize = Math.min(100, Math.max(10, Number(req.query.pageSize || 25)));
  res.json({
    rows: rows.slice((page - 1) * pageSize, page * pageSize),
    allRows: rows,
    total: rows.length,
    page,
    pageSize,
    totals: summarizeLabs(rows),
    collectionUid: state.collectionUid || DEFAULT_LABS_ADMIN_UID,
    adminUrlEn: adminCollectionUrl(state.collectionUid || DEFAULT_LABS_ADMIN_UID, "en"),
    adminUrlAr: adminCollectionUrl(state.collectionUid || DEFAULT_LABS_ADMIN_UID, "ar"),
  });
});

app.get("/api/labs/:documentId", async (req, res) => {
  const state = await loadLabsState();
  const row = activeLabs(state).find((item) => item.documentId === req.params.documentId);
  if (!row) return res.status(404).json({ error: "Lab item not found" });
  res.json({ row });
});

app.post("/api/labs/fetch", async (req, res) => {
  try {
    const collectionUid = cleanCollectionUid(req.body?.collectionUid);
    const [en, ar] = await Promise.all([
      fetchAdminCollectionLocale(collectionUid, "en"),
      fetchAdminCollectionLocale(collectionUid, "ar"),
    ]);
    const current = mergeGenericByDocumentId(en, ar, collectionUid);
    const state = await loadLabsState();
    const { rows, snapshot } = compareLabs(state, current, collectionUid);
    await saveLabsState(state);
    res.json({
      rows,
      totals: summarizeLabs(rows),
      snapshot: snapshotSummary(snapshot),
      snapshots: (state.snapshots ?? []).map(snapshotSummary),
      collectionUid,
      adminUrlEn: adminCollectionUrl(collectionUid, "en"),
      adminUrlAr: adminCollectionUrl(collectionUid, "ar"),
      sourceTotals: { en: en.length, ar: ar.length, merged: current.length },
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get("/api/snapshots/:id/:changeType", async (req, res) => {
  const state = await loadState();
  const snapshot = (state.snapshots ?? []).find((item) => item.id === req.params.id);
  if (!snapshot) return res.status(404).json({ error: "Snapshot not found" });
  const key = String(req.params.changeType || "").toLowerCase();
  const rows = snapshot.changes?.[key] ?? [];
  res.json({ snapshot: snapshotSummary(snapshot), changeType: key, rows, totals: summarize(rows) });
});

app.get("/api/download/:format", async (req, res) => {
  try {
    const state = await loadState();
    let rows = sortRows(filterRows(activeRows(state), req.query), req.query.sort);
    if (!rows.length) return res.status(400).json({ error: "No devices available. Click Refresh CMS Data first." });
    const format = String(req.params.format || "xlsx").toLowerCase();
    const builder = format === "csv" ? buildCsv : format === "pdf" ? buildPdf : buildExcel;
    const { fileName, filePath } = await builder(rows);
    res.download(filePath, fileName);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

const port = Number(process.env.PORT || 4321);
app.listen(port, () => {
  console.log(`STC Device Asset Dashboard running at http://localhost:${port}`);
});
}
