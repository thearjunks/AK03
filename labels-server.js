import fs from "node:fs/promises";
import https from "node:https";
import crypto from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import express from "express";
import ExcelJS from "exceljs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DATA_PATH = path.join(__dirname, "data", "label-state.json");
const ADMIN_SUPPLEMENT_PATH = path.join(__dirname, "data", "admin-label-supplement.json");
const EXPORT_DIR = path.join(__dirname, "outputs", "dashboard-downloads");
const DICTIONARY_EXPORT_DIR = path.join(__dirname, "outputs", "dictionary-generator");
const STRAPI_LABELS_API = "https://content.stc.com.kw/api/stc-labels";
const PAGE_SIZE = 25;
const FETCH_PAGE_SIZE = 100;

const app = express();
app.use(express.json({ limit: "10mb" }));
app.use(express.static(path.join(__dirname, "public-labels")));

function text(value) {
  return String(value ?? "");
}

function nowIso() {
  return new Date().toISOString();
}

function hash(value, length = 12) {
  return crypto.createHash("sha256").update(text(value)).digest("hex").slice(0, length).toUpperCase();
}

async function writeState(state) {
  await fs.mkdir(path.dirname(DATA_PATH), { recursive: true });
  await fs.writeFile(DATA_PATH, `${JSON.stringify(state, null, 2)}\n`, "utf8");
}

function getJson(url) {
  return new Promise((resolve, reject) => {
    const request = https.get(url, {
      headers: {
        Accept: "application/json",
        "User-Agent": "AK-stc-labels-dashboard/1.0",
      },
    }, (response) => {
      let body = "";
      response.setEncoding("utf8");
      response.on("data", (chunk) => { body += chunk; });
      response.on("end", () => {
        if (response.statusCode < 200 || response.statusCode >= 300) {
          reject(new Error(`Strapi returned ${response.statusCode}: ${body.slice(0, 250)}`));
          return;
        }
        try {
          resolve(JSON.parse(body));
        } catch (error) {
          reject(error);
        }
      });
    });
    request.on("error", reject);
    request.setTimeout(60000, () => request.destroy(new Error("Strapi request timed out")));
  });
}

function hasArabic(value) {
  return /[\u0600-\u06FF]/.test(text(value));
}

function hasEnglish(value) {
  return /[A-Za-z]/.test(text(value));
}

function isFullEnglish(value) {
  const clean = text(value).trim();
  return Boolean(clean) && hasEnglish(clean) && !hasArabic(clean);
}

async function readState() {
  const raw = await fs.readFile(DATA_PATH, "utf8");
  const state = JSON.parse(raw);
  const changeHistory = state.changeHistory?.length ? state.changeHistory : historyFromSnapshots(state.snapshots ?? []);
  return {
    initializedAt: state.initializedAt ?? null,
    labels: state.labels ?? {},
    snapshots: state.snapshots ?? [],
    lastSnapshot: state.lastSnapshot ?? null,
    changeHistory,
  };
}

async function readAdminSupplement() {
  try {
    const raw = await fs.readFile(ADMIN_SUPPLEMENT_PATH, "utf8");
    const supplement = JSON.parse(raw);
    return {
      capturedAt: supplement.capturedAt ?? null,
      inventoryCount: Number(supplement.inventoryCount) || 0,
      entries: Array.isArray(supplement.entries) ? supplement.entries : [],
    };
  } catch (error) {
    if (error.code === "ENOENT") return { capturedAt: null, inventoryCount: 0, entries: [] };
    throw error;
  }
}

function rowsFromState(state, includeRemoved = true) {
  const rows = Object.values(state.labels ?? {});
  return includeRemoved ? rows : rows.filter((row) => row.changeType !== "Removed");
}

function stableLabelKey(row) {
  return [
    text(row.featureKey).trim(),
    text(row.labelKey).trim(),
    text(row.component).trim(),
    text(row.componentId).trim(),
  ].join("\u001f");
}

function fallbackLabelKey(row) {
  return [
    text(row.featureKey).trim(),
    text(row.labelKey).trim(),
    text(row.component).trim(),
  ].join("\u001f");
}

function labelSignature(row) {
  return hash(JSON.stringify({
    featureName: row.featureName || "",
    featureKey: row.featureKey || "",
    labelKey: row.labelKey || "",
    englishText: row.englishText || "",
    arabicText: row.arabicText || "",
    component: row.component || "",
    componentId: row.componentId || "",
    entryId: row.entryId || "",
    documentId: row.documentId || "",
  }), 24);
}

function historyId(type, row, changedAt, details = []) {
  return `HIS-${hash(JSON.stringify({
    type,
    labelId: row.labelId,
    featureKey: row.featureKey,
    labelKey: row.labelKey,
    component: row.component,
    componentId: row.componentId,
    changedAt,
    details,
  }), 16)}`;
}

function rowChanged(previous, current) {
  if (!previous) return false;
  return [
    "featureKey",
    "labelKey",
    "englishText",
    "arabicText",
  ].some((field) => text(previous[field]) !== text(current[field]));
}

function changeDetails(previous, current) {
  if (!previous) return [];
  return [
    ["featureKey", "Feature Key"],
    ["labelKey", "Label Key"],
    ["englishText", "English Text"],
    ["arabicText", "Arabic Text"],
  ]
    .filter(([field]) => text(previous[field]) !== text(current[field]))
    .map(([field, label]) => ({ field, label, before: previous[field] ?? "", after: current[field] ?? "" }));
}

function changedBy(row) {
  return row.updatedBy?.firstname || row.updatedBy?.username || row.updatedBy?.email || row.createdBy?.firstname || row.createdBy?.username || row.createdBy?.email || "";
}

function historyRecord(type, row, changedAt, details = [], snapshotId = "") {
  return {
    historyId: historyId(type, row, changedAt, details),
    labelId: row.labelId,
    featureName: row.featureName,
    featureKey: row.featureKey,
    labelKey: row.labelKey,
    englishText: row.englishText,
    arabicText: row.arabicText,
    status: row.status,
    changeType: type,
    changeSummary: type === "Removed" ? "Removed from latest Strapi fetch" : details.map((item) => item.label).join(", "),
    changedAt,
    changedBy: changedBy(row) || "Captured from dashboard fetch",
    source: changedBy(row) ? "Strapi content history" : "Dashboard snapshot comparison",
    snapshotId,
    component: row.component,
    componentId: row.componentId,
    entryId: row.entryId,
    documentId: row.documentId,
    locale: row.locale,
    changeDetails: details,
  };
}

function historyFromSnapshots(snapshots = []) {
  const rows = [];
  const seen = new Set();
  for (const snapshot of snapshots) {
    const changedAt = snapshot.fetchedAt || "";
    const groups = [
      ["Created", snapshot.changes?.new ?? []],
      ["Modified", snapshot.changes?.modified ?? []],
      ["Removed", snapshot.changes?.removed ?? []],
    ];
    for (const [type, items] of groups) {
      for (const item of items) {
        const record = historyRecord(type, item, changedAt, item.changeDetails ?? [], snapshot.id ?? "");
        if (seen.has(record.historyId)) continue;
        seen.add(record.historyId);
        rows.push(record);
      }
    }
  }
  return rows.sort((a, b) => text(b.changedAt).localeCompare(text(a.changedAt)));
}

function extractLabelRows(entry) {
  const rows = [];
  const entryId = entry.id ?? "";
  const documentId = entry.documentId ?? "";
  const featureName = entry.featureName ?? entry.featureKey ?? "";
  const featureKey = entry.featureKey ?? entry.featureName ?? "";
  const locale = entry.locale ?? "";
  const sourceCreatedAt = entry.createdAt ?? "";
  const sourceUpdatedAt = entry.updatedAt ?? "";

  for (const [componentName, value] of Object.entries(entry)) {
    if (!Array.isArray(value)) continue;
    for (const item of value) {
      if (!item || typeof item !== "object") continue;
      const labelKey = item.key ?? item.labelKey ?? item.name ?? item.mediaLabel ?? "";
      const englishText = item.enValue ?? item.en ?? item.englishText ?? item.valueEn ?? "";
      const arabicText = item.arValue ?? item.ar ?? item.arabicText ?? item.valueAr ?? "";
      if (!labelKey && !englishText && !arabicText) continue;
      rows.push({
        featureName,
        featureKey,
        labelKey,
        englishText,
        arabicText,
        description: item.description ?? "",
        module: entry.module ?? "",
        screenName: entry.screenName ?? "",
        status: "Active",
        component: componentName,
        componentId: item.id ?? "",
        entryId,
        documentId,
        locale,
        sourceCreatedAt,
        sourceUpdatedAt,
        createdBy: entry.createdBy ?? item.createdBy ?? null,
        updatedBy: entry.updatedBy ?? item.updatedBy ?? null,
        comments: "",
      });
    }
  }
  return rows;
}

async function fetchLatestStrapiLabels() {
  const all = [];
  let page = 1;
  let pageCount = 1;
  do {
    const url = new URL(STRAPI_LABELS_API);
    url.searchParams.set("pagination[page]", String(page));
    url.searchParams.set("pagination[pageSize]", String(FETCH_PAGE_SIZE));
    url.searchParams.set("populate", "*");
    url.searchParams.set("status", "draft");
    const json = await getJson(url);
    all.push(...(json.data ?? []));
    pageCount = json.meta?.pagination?.pageCount ?? 1;
    page += 1;
  } while (page <= pageCount);

  const publicEntryCount = all.length;
  const supplement = await readAdminSupplement();
  const knownDocumentIds = new Set(all.map((entry) => entry.documentId).filter(Boolean));
  let supplementedEntryCount = 0;
  for (const entry of supplement.entries) {
    if (!entry?.documentId || knownDocumentIds.has(entry.documentId)) continue;
    all.push(entry);
    knownDocumentIds.add(entry.documentId);
    supplementedEntryCount += 1;
  }

  return {
    rows: all.flatMap(extractLabelRows),
    entryCount: Math.max(all.length, supplement.inventoryCount),
    publicEntryCount,
    supplementedEntryCount,
    supplementCapturedAt: supplement.capturedAt,
    source: supplementedEntryCount
      ? "Strapi latest draft-status API + authenticated admin supplement"
      : "Strapi latest draft-status API",
  };
}

function compareAndSaveLabels(state, currentRows, sourceInfo = {}) {
  const seenAt = nowIso();
  const previousRows = rowsFromState(state, true);
  const previousByStable = new Map();
  const previousByFallback = new Map();
  for (const row of previousRows) {
    previousByStable.set(stableLabelKey(row), row);
    if (!previousByFallback.has(fallbackLabelKey(row))) previousByFallback.set(fallbackLabelKey(row), row);
  }

  const nextLabels = {};
  const seenPreviousIds = new Set();
  const changes = { new: [], modified: [], removed: [] };

  for (const raw of currentRows) {
    let old = previousByStable.get(stableLabelKey(raw)) ?? previousByFallback.get(fallbackLabelKey(raw));
    if (old?.labelId && seenPreviousIds.has(old.labelId)) old = undefined;
    const labelId = old?.labelId ?? `LBL-${hash(stableLabelKey(raw) || fallbackLabelKey(raw), 12)}`;
    const details = changeDetails(old, raw);
    let changeType = "Existing";
    if (!old || old.changeType === "Removed") changeType = "New";
    else if (rowChanged(old, raw)) changeType = "Modified";

      const row = {
      ...old,
      ...raw,
      labelId,
      legacyLabelId: old?.legacyLabelId ?? "",
      firstSeenAt: old?.firstSeenAt ?? seenAt,
      lastSeenAt: seenAt,
      lastModifiedAt: changeType === "Modified" || changeType === "New" ? seenAt : (old?.lastModifiedAt ?? raw.sourceUpdatedAt ?? seenAt),
      changeType,
      changeDetails: details,
      changeSummary: details.map((item) => item.label).join(", "),
    };
    row.signature = labelSignature(row);
    nextLabels[labelId] = row;
    if (old?.labelId) seenPreviousIds.add(old.labelId);
    if (changeType === "New") changes.new.push(compact(row));
    if (changeType === "Modified") changes.modified.push(compact(row));
  }

  for (const old of previousRows) {
    if (old.changeType === "Removed" || seenPreviousIds.has(old.labelId)) {
      if (old.changeType === "Removed" && !nextLabels[old.labelId]) nextLabels[old.labelId] = old;
      continue;
    }
    const removed = {
      ...old,
      changeType: "Removed",
      lastSeenAt: seenAt,
      lastModifiedAt: seenAt,
      changeDetails: [],
      changeSummary: "Removed from latest Strapi fetch",
    };
    nextLabels[removed.labelId] = removed;
    changes.removed.push(compact(removed));
  }

  state.initializedAt ||= seenAt;
  state.labels = nextLabels;
  const allRows = rowsFromState(state, true);
  const counts = summarize(allRows, allRows);
  const snapshot = {
    id: `snap-${seenAt.replace(/[:.]/g, "-")}`,
    fetchedAt: seenAt,
    source: sourceInfo.source ?? "Strapi latest draft-status API",
    sourceTotals: {
      entries: sourceInfo.entryCount ?? new Set(currentRows.map((row) => row.documentId).filter(Boolean)).size,
      labels: currentRows.length,
      publicEntries: sourceInfo.publicEntryCount,
      adminSupplementEntries: sourceInfo.supplementedEntryCount,
      adminSupplementCapturedAt: sourceInfo.supplementCapturedAt,
    },
    counts: {
      existing: counts.existing,
      new: counts.new,
      modified: counts.modified,
      removed: changes.removed.length,
      total: counts.total,
      featureKeys: sourceInfo.entryCount ?? counts.featureKeys,
    },
    changes,
  };
  const existingHistoryIds = new Set((state.changeHistory ?? []).map((record) => record.historyId));
  const historyRows = [
    ...changes.new.map((row) => historyRecord("Created", row, seenAt, row.changeDetails ?? [], snapshot.id)),
    ...changes.modified.map((row) => historyRecord("Modified", row, seenAt, row.changeDetails ?? [], snapshot.id)),
    ...changes.removed.map((row) => historyRecord("Removed", row, seenAt, row.changeDetails ?? [], snapshot.id)),
  ].filter((record) => !existingHistoryIds.has(record.historyId));
  state.lastSnapshot = snapshot;
  state.snapshots = [snapshot, ...(state.snapshots ?? [])].slice(0, 25);
  state.changeHistory = [...historyRows, ...(state.changeHistory ?? [])].slice(0, 20000);
  return { state, snapshot, rows: allRows };
}

function languageIssues(row) {
  const issues = [];
  if (hasArabic(row.englishText)) issues.push("English has Arabic");
  if (hasEnglish(row.arabicText)) issues.push("Arabic has English");
  if (isFullEnglish(row.arabicText)) issues.push("Arabic full English");
  if (!text(row.englishText).trim()) issues.push("English Missing");
  if (!text(row.arabicText).trim()) issues.push("Arabic Missing");
  return issues;
}

function normalizedDisplayText(value) {
  return text(value).trim().replace(/\s+/g, " ");
}

function normalizedEnglishKey(value) {
  return normalizedDisplayText(value).toLocaleLowerCase("en");
}

function groupRows(rows, keyForRow) {
  const groups = new Map();
  for (const row of rows) {
    const key = keyForRow(row);
    if (!key) continue;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  }
  return groups;
}

function languageConsistencyRecords(rows) {
  const active = rows.filter((row) => row.changeType !== "Removed");
  const records = [];
  const emitted = new Set();

  const addRecord = (issueType, groupKey, row, occurrenceCount, details, variants = []) => {
    const recordKey = `${issueType}\u001f${groupKey}\u001f${row.labelId}`;
    if (emitted.has(recordKey)) return;
    emitted.add(recordKey);
    records.push({
      issueId: `ISS-${hash(recordKey, 14)}`,
      groupId: `GRP-${hash(`${issueType}\u001f${groupKey}`, 12)}`,
      issueType,
      occurrenceCount,
      details,
      variants,
      labelId: row.labelId,
      featureName: row.featureName,
      featureKey: row.featureKey,
      labelKey: row.labelKey,
      englishText: row.englishText,
      arabicText: row.arabicText,
      module: row.module,
      screenName: row.screenName,
      status: row.status,
      changeType: row.changeType,
      firstSeenAt: row.firstSeenAt,
      lastModifiedAt: row.lastModifiedAt,
      component: row.component,
      componentId: row.componentId,
      entryId: row.entryId,
      documentId: row.documentId,
      locale: row.locale,
      comments: row.comments,
    });
  };

  for (const row of active) {
    const english = normalizedDisplayText(row.englishText);
    const arabic = normalizedDisplayText(row.arabicText);
    if (!english) addRecord("English missing", row.labelId, row, 1, "English text is empty.");
    if (!arabic) addRecord("Arabic missing", row.labelId, row, 1, "Arabic text is empty.");
    if (hasArabic(english)) addRecord("English text contains Arabic", row.labelId, row, 1, "Arabic characters appear in the English field.");
    if (arabic && isFullEnglish(arabic)) {
      addRecord("Arabic text is full English", row.labelId, row, 1, "The Arabic field contains only English text.");
    } else if (arabic && hasEnglish(arabic)) {
      addRecord("Arabic text contains English", row.labelId, row, 1, "English characters appear in the Arabic field.");
    }
  }

  const byEnglish = groupRows(active, (row) => normalizedEnglishKey(row.englishText));
  for (const [englishKey, group] of byEnglish) {
    const englishVariants = [...new Set(group.map((row) => normalizedDisplayText(row.englishText)).filter(Boolean))];
    if (englishVariants.length > 1) {
      const details = `English casing variants: ${englishVariants.join(" | ")}`;
      for (const row of group) addRecord("English case mismatch", englishKey, row, group.length, details, englishVariants);
    }

    const arabicVariants = [...new Set(group.map((row) => normalizedDisplayText(row.arabicText) || "(missing)"))];
    if (arabicVariants.length > 1) {
      const details = `The same English text maps to ${arabicVariants.length} Arabic variants.`;
      for (const row of group) addRecord("Translation mismatch", englishKey, row, group.length, details, arabicVariants);
    }
  }

  const byArabic = groupRows(active, (row) => normalizedDisplayText(row.arabicText));
  for (const [arabicKey, group] of byArabic) {
    const englishVariants = [...new Set(group.map((row) => normalizedDisplayText(row.englishText) || "(missing)"))];
    const normalizedVariants = new Set(englishVariants.map(normalizedEnglishKey));
    if (normalizedVariants.size > 1) {
      const details = `The same Arabic text maps to ${englishVariants.length} English variants.`;
      for (const row of group) addRecord("Duplicate translation / Arabic inconsistency", arabicKey, row, group.length, details, englishVariants);
    }
  }

  const byLabelKey = groupRows(active, (row) => {
    const featureKey = normalizedDisplayText(row.featureKey);
    const labelKey = normalizedDisplayText(row.labelKey);
    return featureKey && labelKey ? `${featureKey}\u001f${labelKey}` : "";
  });
  for (const [labelGroupKey, group] of byLabelKey) {
    if (group.length < 2) continue;
    const contentVariants = [...new Set(group.map((row) => `${normalizedEnglishKey(row.englishText)}\u001f${normalizedDisplayText(row.arabicText)}`))];
    if (contentVariants.length > 1) {
      const details = `This Feature Key and Label Key combination has ${contentVariants.length} content variants.`;
      for (const row of group) addRecord("Label key content mismatch", labelGroupKey, row, group.length, details);
    }
  }

  return records.sort((a, b) =>
    a.issueType.localeCompare(b.issueType)
    || b.occurrenceCount - a.occurrenceCount
    || text(a.featureKey).localeCompare(text(b.featureKey))
    || text(a.labelKey).localeCompare(text(b.labelKey)));
}

function summarizeLanguageConsistency(records) {
  const groups = new Map();
  for (const record of records) groups.set(record.groupId, record.issueType);
  const countGroups = (issueType) => [...groups.values()].filter((value) => value === issueType).length;
  const directTypes = new Set([
    "English missing",
    "Arabic missing",
    "English text contains Arabic",
    "Arabic text contains English",
    "Arabic text is full English",
  ]);
  return {
    affectedLabels: new Set(records.map((record) => record.labelId)).size,
    issueGroups: groups.size,
    issueOccurrences: records.length,
    englishCaseMismatchGroups: countGroups("English case mismatch"),
    translationMismatchGroups: countGroups("Translation mismatch"),
    duplicateTranslationGroups: countGroups("Duplicate translation / Arabic inconsistency"),
    labelKeyMismatchGroups: countGroups("Label key content mismatch"),
    directLanguageIssueGroups: [...groups.entries()].filter(([, issueType]) => directTypes.has(issueType)).length,
  };
}

function filterLanguageConsistency(records, query) {
  const search = normalizedDisplayText(query.search).toLowerCase();
  return records.filter((record) => {
    if (query.issueType && query.issueType !== "all" && record.issueType !== query.issueType) return false;
    if (query.featureKey && query.featureKey !== "all" && record.featureKey !== query.featureKey) return false;
    if (query.status && query.status !== "all" && record.status !== query.status) return false;
    if (!search) return true;
    const haystack = [
      record.issueType,
      record.details,
      record.englishText,
      record.arabicText,
      record.featureName,
      record.featureKey,
      record.labelKey,
      record.module,
      record.screenName,
      record.status,
      ...(record.variants ?? []),
    ].map(text).join(" ").toLowerCase();
    return haystack.includes(search);
  });
}

function summarize(rows, allRows = rows) {
  const active = rows.filter((row) => row.changeType !== "Removed");
  const issueRows = active.filter((row) => languageIssues(row).length);
  const count = (type) => active.filter((row) => row.changeType === type).length;
  return {
    total: active.length,
    existing: count("Existing"),
    new: count("New"),
    modified: count("Modified"),
    removed: allRows.filter((row) => row.changeType === "Removed").length,
    languageIssues: issueRows.length,
    featureKeys: new Set(active.map((row) => row.featureKey).filter(Boolean)).size,
    englishHasArabic: active.filter((row) => hasArabic(row.englishText)).length,
    arabicHasEnglish: active.filter((row) => hasEnglish(row.arabicText)).length,
    arabicFullEnglish: active.filter((row) => isFullEnglish(row.arabicText)).length,
    englishMissing: active.filter((row) => !text(row.englishText).trim()).length,
    arabicMissing: active.filter((row) => !text(row.arabicText).trim()).length,
  };
}

function options(rows) {
  const uniq = (field) => [...new Set(rows.map((row) => row[field]).filter(Boolean))].sort((a, b) => text(a).localeCompare(text(b)));
  return {
    featureKeys: uniq("featureKey"),
    statuses: uniq("status"),
    changes: uniq("changeType"),
  };
}

function matchesLanguage(row, filter) {
  if (!filter || filter === "all") return true;
  const englishHasArabic = hasArabic(row.englishText);
  const arabicHasEnglish = hasEnglish(row.arabicText);
  const englishMissing = !text(row.englishText).trim();
  const arabicMissing = !text(row.arabicText).trim();
  if (filter === "any") return englishHasArabic || arabicHasEnglish || englishMissing || arabicMissing;
  if (filter === "englishHasArabic") return englishHasArabic;
  if (filter === "arabicHasEnglish") return arabicHasEnglish;
  if (filter === "arabicFullEnglish") return isFullEnglish(row.arabicText);
  if (filter === "englishMissing") return englishMissing;
  if (filter === "arabicMissing") return arabicMissing;
  return true;
}

function filterRows(rows, query) {
  const search = text(query.search).trim().toLowerCase();
  const advanced = ["labelId", "featureKey", "labelKey", "englishText", "arabicText", "status", "changeType", "firstSeenAt"]
    .reduce((acc, key) => ({ ...acc, [key]: text(query[key]).trim().toLowerCase() }), {});

  return rows.filter((row) => {
    const haystack = [
      row.labelId,
      row.legacyLabelId,
      row.featureKey,
      row.labelKey,
      row.englishText,
      row.arabicText,
      row.status,
      row.changeType,
      row.firstSeenAt,
      row.lastModifiedAt,
      row.entryId,
      row.documentId,
      languageIssues(row).join(" "),
    ].map(text).join(" ").toLowerCase();

    if (row.changeType === "Removed" && query.changeType !== "Removed") return false;
    if (search && !haystack.includes(search)) return false;
    if (query.featureKey && query.featureKey !== "all" && row.featureKey !== query.featureKey) return false;
    if (query.status && query.status !== "all" && row.status !== query.status) return false;
    if (query.changeType && query.changeType !== "all" && row.changeType !== query.changeType) return false;
    if (!matchesLanguage(row, query.languageIssue)) return false;
    for (const [key, value] of Object.entries(advanced)) {
      if (value && !text(row[key]).toLowerCase().includes(value)) return false;
    }
    if (query.languageIssueText && !languageIssues(row).join(" ").toLowerCase().includes(text(query.languageIssueText).toLowerCase())) return false;
    return true;
  });
}

function filterHistory(rows, query) {
  const search = text(query.search).trim().toLowerCase();
  return rows.filter((row) => {
    const haystack = [
      row.historyId,
      row.labelId,
      row.featureName,
      row.featureKey,
      row.labelKey,
      row.englishText,
      row.arabicText,
      row.status,
      row.changeType,
      row.changeSummary,
      row.changedAt,
      row.changedBy,
      row.source,
      row.entryId,
      row.documentId,
    ].map(text).join(" ").toLowerCase();
    if (search && !haystack.includes(search)) return false;
    if (query.featureKey && query.featureKey !== "all" && row.featureKey !== query.featureKey) return false;
    if (query.changeType && query.changeType !== "all" && row.changeType !== query.changeType) return false;
    if (query.status && query.status !== "all" && row.status !== query.status) return false;
    return true;
  });
}

function compact(row) {
  return {
    labelId: row.labelId,
    legacyLabelId: row.legacyLabelId,
    featureName: row.featureName,
    featureKey: row.featureKey,
    labelKey: row.labelKey,
    englishText: row.englishText,
    arabicText: row.arabicText,
    description: row.description,
    module: row.module,
    screenName: row.screenName,
    status: row.status,
    firstSeenAt: row.firstSeenAt,
    lastModifiedAt: row.lastModifiedAt,
    comments: row.comments,
    changeType: row.changeType,
    component: row.component,
    componentId: row.componentId,
    entryId: row.entryId,
    documentId: row.documentId,
    locale: row.locale,
    languageIssue: languageIssues(row).join(", ") || "None",
    changeSummary: row.changeSummary,
    changeDetails: row.changeDetails ?? [],
  };
}

function snapshotSummary(snapshot) {
  const changes = snapshot.changes ?? {};
  return {
    id: snapshot.id,
    fetchedAt: snapshot.fetchedAt,
    counts: snapshot.counts ?? {},
    changeCounts: {
      new: changes.new?.length ?? snapshot.counts?.new ?? 0,
      modified: changes.modified?.length ?? snapshot.counts?.modified ?? 0,
      removed: changes.removed?.length ?? snapshot.counts?.removed ?? 0,
    },
    source: snapshot.source ?? "",
    sourceTotals: snapshot.sourceTotals ?? {},
  };
}

async function labelData(query = {}) {
  const state = await readState();
  const allRows = rowsFromState(state, true);
  const filtered = filterRows(allRows, query);
  return { state, allRows, filtered };
}

function workbookCellText(cell) {
  return text(cell?.text ?? cell?.value).trim();
}

function normalizedWorkbookHeader(value) {
  return text(value).trim().toLowerCase().replace(/[^a-z0-9]/g, "");
}

function normalizedDictionaryEnglish(value) {
  return text(value)
    .normalize("NFKC")
    .replace(/[\u2018\u2019\u201B]/g, "'")
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/[\u2010-\u2015\u2212]/g, "-")
    .trim()
    .replace(/\s+/g, " ")
    .toLocaleLowerCase("en-US");
}

function compactDictionaryEnglish(value) {
  return normalizedDictionaryEnglish(value)
    .normalize("NFKD")
    .replace(/\p{M}/gu, "")
    .replace(/&/g, "and")
    .replace(/[^a-z0-9]+/g, "");
}

function englishFormattingQuality(value) {
  const clean = text(value);
  let score = 0;
  if (/\d\s+(?:am|pm)\b/i.test(clean)) score += 3;
  if (/\b[ap]\s+m\b/i.test(clean)) score -= 5;
  if (/\s[-:]\s/.test(clean)) score += 1;
  if (/\s{2,}/.test(clean)) score -= 2;
  return score;
}

function mostFrequent(values, preferEnglishFormatting = false) {
  const counts = new Map();
  for (const value of values.filter(Boolean)) counts.set(value, (counts.get(value) || 0) + 1);
  return [...counts.entries()]
    .sort((a, b) => b[1] - a[1]
      || (preferEnglishFormatting ? englishFormattingQuality(b[0]) - englishFormattingQuality(a[0]) : 0)
      || a[0].localeCompare(b[0]))[0]?.[0] || "";
}

function dictionaryTrigrams(value) {
  const compact = compactDictionaryEnglish(value);
  if (compact.length <= 3) return compact ? [compact] : [];
  const grams = [];
  for (let index = 0; index <= compact.length - 3; index += 1) grams.push(compact.slice(index, index + 3));
  return [...new Set(grams)];
}

function levenshteinDistance(left, right) {
  if (left === right) return 0;
  if (!left.length) return right.length;
  if (!right.length) return left.length;
  let previous = Array.from({ length: right.length + 1 }, (_, index) => index);
  for (let leftIndex = 1; leftIndex <= left.length; leftIndex += 1) {
    const current = [leftIndex];
    for (let rightIndex = 1; rightIndex <= right.length; rightIndex += 1) {
      current[rightIndex] = Math.min(
        current[rightIndex - 1] + 1,
        previous[rightIndex] + 1,
        previous[rightIndex - 1] + (left[leftIndex - 1] === right[rightIndex - 1] ? 0 : 1),
      );
    }
    previous = current;
  }
  return previous[right.length];
}

function mergeDictionaryEntries(entries) {
  const englishValues = entries.flatMap((entry) => entry.englishValues || entry.englishVariants || []);
  const arabicValues = entries.flatMap((entry) => entry.arabicValues || entry.arabicVariants || []);
  return {
    canonicalEnglish: mostFrequent(englishValues, true),
    canonicalArabic: mostFrequent(arabicValues),
    englishVariants: [...new Set(englishValues.filter(Boolean))],
    arabicVariants: [...new Set(arabicValues.filter(Boolean))],
    references: entries.reduce((sum, entry) => sum + Number(entry.references || 0), 0),
  };
}

function buildAllLabelsDictionary(rows) {
  const grouped = new Map();
  for (const row of rows) {
    const english = text(row.englishText).trim().replace(/\s+/g, " ");
    const key = normalizedDictionaryEnglish(english);
    if (!key) continue;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push({
      english,
      arabic: text(row.arabicText).trim().replace(/\s+/g, " "),
      featureKey: row.featureKey,
      labelKey: row.labelKey,
    });
  }

  const dictionary = new Map();
  for (const [key, matches] of grouped) {
    const englishVariants = [...new Set(matches.map((item) => item.english).filter(Boolean))];
    const arabicVariants = [...new Set(matches.map((item) => item.arabic).filter(Boolean))];
    dictionary.set(key, {
      normalizedKey: key,
      compactKey: compactDictionaryEnglish(key),
      canonicalEnglish: mostFrequent(matches.map((item) => item.english), true),
      canonicalArabic: mostFrequent(matches.map((item) => item.arabic)),
      englishVariants,
      arabicVariants,
      englishValues: matches.map((item) => item.english),
      arabicValues: matches.map((item) => item.arabic),
      references: matches.length,
    });
  }

  const compactIndex = new Map();
  const trigramIndex = new Map();
  for (const entry of dictionary.values()) {
    if (entry.compactKey) {
      if (!compactIndex.has(entry.compactKey)) compactIndex.set(entry.compactKey, []);
      compactIndex.get(entry.compactKey).push(entry);
    }
    for (const gram of dictionaryTrigrams(entry.compactKey)) {
      if (!trigramIndex.has(gram)) trigramIndex.set(gram, new Set());
      trigramIndex.get(gram).add(entry);
    }
  }
  return { exactIndex: dictionary, compactIndex, trigramIndex };
}

function resolveDictionaryMatch(originalEnglish, dictionary) {
  const normalized = normalizedDictionaryEnglish(originalEnglish);
  const exactEntry = dictionary.exactIndex.get(normalized);
  const compact = compactDictionaryEnglish(originalEnglish);
  const compactEntries = dictionary.compactIndex.get(compact) || [];
  if (exactEntry && compactEntries.length <= 1) {
    return { entry: exactEntry, matchMethod: "Exact", confidence: 100, requiresReview: false };
  }
  if (compactEntries.length) {
    const entry = mergeDictionaryEntries(compactEntries);
    return { entry, matchMethod: "Formatting normalized", confidence: 100, requiresReview: entry.arabicVariants.length > 1 };
  }

  if (compact.length < 6) return null;
  const candidateScores = new Map();
  for (const gram of dictionaryTrigrams(compact)) {
    for (const entry of dictionary.trigramIndex.get(gram) || []) {
      candidateScores.set(entry, (candidateScores.get(entry) || 0) + 1);
    }
  }
  const candidates = [...candidateScores.entries()]
    .filter(([entry]) => Math.abs(entry.compactKey.length - compact.length) <= Math.max(2, Math.floor(compact.length * 0.12)))
    .sort((a, b) => b[1] - a[1])
    .slice(0, 12)
    .map(([entry]) => {
      const distance = levenshteinDistance(compact, entry.compactKey);
      return { entry, similarity: 1 - (distance / Math.max(compact.length, entry.compactKey.length)) };
    })
    .sort((a, b) => b.similarity - a.similarity);
  const best = candidates[0];
  if (!best) return null;
  const threshold = compact.length < 10 ? 0.9 : 0.88;
  if (best.similarity < threshold) return null;
  return {
    entry: best.entry,
    matchMethod: "Close match",
    confidence: Math.round(best.similarity * 100),
    requiresReview: true,
  };
}

const ENGLISH_HEADER_ALIASES = new Map([
  ["englishtext", 100],
  ["englishlabel", 100],
  ["envalue", 100],
  ["enlabel", 95],
  ["labelenglish", 95],
  ["englishtranslation", 90],
  ["english", 90],
  ["en", 80],
  ["label", 60],
  ["labels", 60],
  ["text", 50],
  ["value", 50],
]);

function findEnglishColumn(sheet) {
  const maxHeaderRow = Math.min(20, Math.max(1, sheet.actualRowCount));
  const maxColumn = Math.max(1, sheet.actualColumnCount);
  let best = null;
  for (let rowNumber = 1; rowNumber <= maxHeaderRow; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    for (let columnNumber = 1; columnNumber <= maxColumn; columnNumber += 1) {
      const header = workbookCellText(row.getCell(columnNumber));
      const normalized = normalizedWorkbookHeader(header);
      let score = ENGLISH_HEADER_ALIASES.get(normalized) || 0;
      if (!score && normalized.includes("english") && !normalized.includes("arabic")) score = 75;
      if (!score) continue;
      const candidate = { rowNumber, columnNumber, header, score };
      if (!best || candidate.score > best.score || (candidate.score === best.score && candidate.rowNumber < best.rowNumber)) best = candidate;
    }
  }
  return best;
}

function generatedColumnMap(sheet, headerRowNumber, englishColumnNumber) {
  let headerRow = sheet.getRow(headerRowNumber);
  const originalColumnCount = Math.max(sheet.columnCount, sheet.actualColumnCount);
  let hasSuggestedColumn = false;
  for (let columnNumber = 1; columnNumber <= sheet.actualColumnCount; columnNumber += 1) {
    if (workbookCellText(headerRow.getCell(columnNumber)) === "Suggested English Text") hasSuggestedColumn = true;
  }
  if (!hasSuggestedColumn) sheet.spliceColumns(englishColumnNumber + 1, 0, []);

  headerRow = sheet.getRow(headerRowNumber);
  const knownColumnCount = Math.max(
    sheet.columnCount,
    sheet.actualColumnCount,
    originalColumnCount + (hasSuggestedColumn ? 0 : 1),
  );
  const wanted = [
    "Suggested English Text",
    "Arabic Translation",
    "Original English Text",
    "Match Status",
    "Match Method",
    "Dictionary Match",
    "Review Status",
  ];
  const existing = new Map();
  for (let columnNumber = 1; columnNumber <= knownColumnCount; columnNumber += 1) {
    existing.set(workbookCellText(headerRow.getCell(columnNumber)), columnNumber);
  }
  let nextColumn = knownColumnCount + 1;
  const columns = {};
  for (const header of wanted) {
    columns[header] = header === "Suggested English Text" && !hasSuggestedColumn
      ? englishColumnNumber + 1
      : existing.get(header) || nextColumn++;
    const cell = headerRow.getCell(columns[header]);
    cell.value = header;
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4F008C" } };
    cell.alignment = { vertical: "middle", wrapText: true };
    sheet.getColumn(columns[header]).width = ["Arabic Translation", "Suggested English Text", "Original English Text"].includes(header) ? 34 : 22;
  }
  return columns;
}

function classifyDictionaryLabel(originalEnglish, dictionaryMatch) {
  if (!dictionaryMatch) {
    return {
      suggestedEnglish: "",
      standardizedEnglish: "",
      arabicTranslation: "",
      matchStatus: "Not Found in Dictionary",
      matchMethod: "No match",
      matchConfidence: 0,
      dictionaryMatch: "No",
      reviewStatus: "Translation Required",
    };
  }

  const dictionaryEntry = dictionaryMatch.entry;
  const updated = originalEnglish !== dictionaryEntry.canonicalEnglish;
  const needsReview = dictionaryMatch.requiresReview || dictionaryEntry.arabicVariants.length > 1;
  if (!dictionaryEntry.canonicalArabic) {
    return {
      suggestedEnglish: dictionaryEntry.canonicalEnglish,
      standardizedEnglish: dictionaryEntry.canonicalEnglish,
      arabicTranslation: "",
      matchStatus: "Translation Required",
      matchMethod: dictionaryMatch.matchMethod,
      matchConfidence: dictionaryMatch.confidence,
      dictionaryMatch: "Yes",
      reviewStatus: "Translation Required",
    };
  }
  return {
    suggestedEnglish: dictionaryEntry.canonicalEnglish,
    standardizedEnglish: dictionaryEntry.canonicalEnglish,
    arabicTranslation: dictionaryEntry.canonicalArabic,
    matchStatus: updated ? "Updated" : "Matched",
    matchMethod: dictionaryMatch.matchMethod,
    matchConfidence: dictionaryMatch.confidence,
    dictionaryMatch: "Yes",
    reviewStatus: needsReview ? "Needs Review" : "Not Required",
  };
}

function summarizeDictionaryRows(rows, sheetSummaries = []) {
  return {
    totalLabels: rows.length,
    matchedLabels: rows.filter((row) => row.dictionaryMatch === "Yes").length,
    updatedLabels: rows.filter((row) => row.matchStatus === "Updated").length,
    notFound: rows.filter((row) => row.matchStatus === "Not Found in Dictionary").length,
    translationRequired: rows.filter((row) => row.reviewStatus === "Translation Required").length,
    needsReview: rows.filter((row) => row.reviewStatus === "Needs Review").length,
    processingProgress: 100,
    sheetsProcessed: sheetSummaries.length,
  };
}

async function writeDictionaryExceptionWorkbook(rows, filePath) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "AK stc labels Dictionary Generator";
  const sheet = workbook.addWorksheet("Unmatched Labels");
  sheet.columns = [
    { header: "Sheet", key: "sheetName", width: 24 },
    { header: "Row Number", key: "rowNumber", width: 14 },
    { header: "Original English Label", key: "originalEnglish", width: 48 },
    { header: "Suggested English Text", key: "suggestedEnglish", width: 48 },
    { header: "Match Status", key: "matchStatus", width: 28 },
    { header: "Match Method", key: "matchMethod", width: 24 },
    { header: "Review Status", key: "reviewStatus", width: 24 },
  ];
  sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  sheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4F008C" } };
  rows.filter((row) => row.dictionaryMatch === "No").forEach((row) => sheet.addRow(row));
  sheet.views = [{ state: "frozen", ySplit: 1 }];
  sheet.autoFilter = "A1:G1";
  await workbook.xlsx.writeFile(filePath);
}

async function writeDictionarySummaryWorkbook(summary, sheetSummaries, filePath) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "AK stc labels Dictionary Generator";
  const sheet = workbook.addWorksheet("Processing Summary");
  sheet.columns = [{ header: "Metric", key: "metric", width: 32 }, { header: "Count", key: "count", width: 18 }];
  [
    ["Total Labels", summary.totalLabels],
    ["Matched Labels", summary.matchedLabels],
    ["Updated Labels", summary.updatedLabels],
    ["Labels Not Found", summary.notFound],
    ["Translation Required", summary.translationRequired],
    ["Needs Review", summary.needsReview],
    ["Sheets Processed", summary.sheetsProcessed],
  ].forEach(([metric, count]) => sheet.addRow({ metric, count }));
  sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  sheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4F008C" } };

  const detail = workbook.addWorksheet("Sheet Coverage");
  detail.columns = [
    { header: "Sheet", key: "sheetName", width: 28 },
    { header: "Header Row", key: "headerRow", width: 14 },
    { header: "English Column", key: "englishHeader", width: 28 },
    { header: "Labels Processed", key: "labelsProcessed", width: 18 },
  ];
  sheetSummaries.forEach((row) => detail.addRow(row));
  detail.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  detail.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4F008C" } };
  await workbook.xlsx.writeFile(filePath);
}

async function processDictionaryWorkbook(buffer, originalFileName, state) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const dictionary = buildAllLabelsDictionary(rowsFromState(state, false));
  const results = [];
  const sheetSummaries = [];
  const yellow = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };

  for (const sheet of workbook.worksheets) {
    const englishColumn = findEnglishColumn(sheet);
    if (!englishColumn) continue;
    const outputColumns = generatedColumnMap(sheet, englishColumn.rowNumber, englishColumn.columnNumber);
    let labelsProcessed = 0;
    for (let rowNumber = englishColumn.rowNumber + 1; rowNumber <= sheet.actualRowCount; rowNumber += 1) {
      const row = sheet.getRow(rowNumber);
      const sourceEnglish = workbookCellText(row.getCell(englishColumn.columnNumber));
      const preservedOriginal = workbookCellText(row.getCell(outputColumns["Original English Text"]));
      const originalEnglish = preservedOriginal || sourceEnglish;
      if (!originalEnglish) continue;
      const dictionaryMatch = resolveDictionaryMatch(originalEnglish, dictionary);
      const classification = classifyDictionaryLabel(originalEnglish, dictionaryMatch);
      const result = {
        sheetName: sheet.name,
        rowNumber,
        originalEnglish,
        ...classification,
      };
      results.push(result);
      labelsProcessed += 1;
      row.getCell(outputColumns["Original English Text"]).value = originalEnglish;
      row.getCell(outputColumns["Suggested English Text"]).value = classification.suggestedEnglish;
      if (classification.dictionaryMatch === "Yes") row.getCell(englishColumn.columnNumber).value = classification.suggestedEnglish;
      row.getCell(outputColumns["Arabic Translation"]).value = classification.arabicTranslation;
      row.getCell(outputColumns["Arabic Translation"]).alignment = { horizontal: "right", readingOrder: "rtl", wrapText: true };
      row.getCell(outputColumns["Match Status"]).value = classification.matchStatus;
      row.getCell(outputColumns["Match Method"]).value = `${classification.matchMethod} (${classification.matchConfidence}%)`;
      row.getCell(outputColumns["Dictionary Match"]).value = classification.dictionaryMatch;
      row.getCell(outputColumns["Review Status"]).value = classification.reviewStatus;
      if (classification.matchStatus === "Not Found in Dictionary" || classification.reviewStatus !== "Not Required") {
        for (let columnNumber = 1; columnNumber <= Math.max(...Object.values(outputColumns)); columnNumber += 1) {
          row.getCell(columnNumber).fill = yellow;
        }
      }
    }
    sheetSummaries.push({
      sheetName: sheet.name,
      headerRow: englishColumn.rowNumber,
      englishHeader: englishColumn.header,
      labelsProcessed,
    });
  }

  if (!sheetSummaries.length) {
    throw new Error("No English label column was found. Use a header such as English Text, English Label, enValue, English, or Label.");
  }

  const jobId = crypto.randomUUID();
  const safeBaseName = path.basename(originalFileName, path.extname(originalFileName)).replace(/[^a-zA-Z0-9_-]+/g, "-").replace(/^-|-$/g, "") || "labels";
  const jobDir = path.join(DICTIONARY_EXPORT_DIR, jobId);
  await fs.mkdir(jobDir, { recursive: true });
  const updatedName = `${safeBaseName}-dictionary-updated.xlsx`;
  const exceptionsName = `${safeBaseName}-dictionary-exceptions.xlsx`;
  const summaryName = `${safeBaseName}-dictionary-summary.xlsx`;
  await workbook.xlsx.writeFile(path.join(jobDir, updatedName));
  const summary = summarizeDictionaryRows(results, sheetSummaries);
  await writeDictionaryExceptionWorkbook(results, path.join(jobDir, exceptionsName));
  await writeDictionarySummaryWorkbook(summary, sheetSummaries, path.join(jobDir, summaryName));
  return { jobId, results, summary, sheetSummaries, files: { updated: updatedName, exceptions: exceptionsName, summary: summaryName } };
}

app.get("/api/status", async (_req, res) => {
  const state = await readState();
  const allRows = rowsFromState(state, true);
  const totals = summarize(allRows, allRows);
  totals.featureEntries = state.lastSnapshot?.sourceTotals?.entries ?? totals.featureKeys;
  res.json({
    app: "AK stc labels",
    sourceMode: state.lastSnapshot?.source ?? "Strapi latest draft-status API",
    initializedAt: state.initializedAt,
    lastSnapshot: state.lastSnapshot ? snapshotSummary(state.lastSnapshot) : null,
    snapshots: state.snapshots.map(snapshotSummary),
    totals,
    historyTotals: summarizeHistory(state.changeHistory ?? []),
    options: options(allRows.filter((row) => row.changeType !== "Removed")),
  });
});

function summarizeHistory(rows = []) {
  const count = (type) => rows.filter((row) => row.changeType === type).length;
  return {
    total: rows.length,
    created: count("Created"),
    modified: count("Modified"),
    removed: count("Removed"),
    featureKeys: new Set(rows.map((row) => row.featureKey).filter(Boolean)).size,
  };
}

app.get("/api/labels", async (req, res) => {
  const { state, allRows, filtered } = await labelData(req.query);
  const page = Math.max(1, Number(req.query.page) || 1);
  const pageSize = Math.min(100, Math.max(1, Number(req.query.pageSize) || PAGE_SIZE));
  const start = (page - 1) * pageSize;
  const totals = summarize(filtered, allRows);
  totals.featureEntries = state.lastSnapshot?.sourceTotals?.entries ?? totals.featureKeys;
  res.json({
    rows: filtered.slice(start, start + pageSize).map(compact),
    total: filtered.length,
    page,
    pageSize,
    totals,
    options: options(allRows.filter((row) => row.changeType !== "Removed")),
    lastSnapshot: state.lastSnapshot ? snapshotSummary(state.lastSnapshot) : null,
    snapshots: state.snapshots.map(snapshotSummary),
  });
});

app.get("/api/language-consistency", async (req, res) => {
  const state = await readState();
  const activeRows = rowsFromState(state, false);
  const allRecords = languageConsistencyRecords(activeRows);
  const filtered = filterLanguageConsistency(allRecords, req.query);
  const page = Math.max(1, Number(req.query.page) || 1);
  const pageSize = Math.min(200, Math.max(1, Number(req.query.pageSize) || 50));
  const start = (page - 1) * pageSize;
  const uniq = (field) => [...new Set(allRecords.map((record) => record[field]).filter(Boolean))]
    .sort((a, b) => text(a).localeCompare(text(b)));
  res.json({
    rows: filtered.slice(start, start + pageSize),
    total: filtered.length,
    page,
    pageSize,
    totals: summarizeLanguageConsistency(filtered),
    allTotals: summarizeLanguageConsistency(allRecords),
    options: {
      issueTypes: uniq("issueType"),
      featureKeys: uniq("featureKey"),
      statuses: uniq("status"),
    },
    lastSnapshot: state.lastSnapshot ? snapshotSummary(state.lastSnapshot) : null,
  });
});

app.get("/api/history", async (req, res) => {
  const state = await readState();
  const rows = filterHistory(state.changeHistory ?? [], req.query);
  const page = Math.max(1, Number(req.query.page) || 1);
  const pageSize = Math.min(100, Math.max(1, Number(req.query.pageSize) || PAGE_SIZE));
  const start = (page - 1) * pageSize;
  res.json({
    rows: rows.slice(start, start + pageSize),
    total: rows.length,
    page,
    pageSize,
    totals: summarizeHistory(rows),
    allTotals: summarizeHistory(state.changeHistory ?? []),
    options: options(rowsFromState(state, false)),
  });
});

app.post("/api/fetch", async (_req, res) => {
  const state = await readState();
  const current = await fetchLatestStrapiLabels();
  const result = compareAndSaveLabels(state, current.rows, current);
  await writeState(result.state);
  const allRows = result.rows;
  const totals = summarize(allRows, allRows);
  totals.featureEntries = result.snapshot.sourceTotals.entries;
  res.json({
    rows: allRows.filter((row) => row.changeType !== "Removed").slice(0, PAGE_SIZE).map(compact),
    total: allRows.length,
    totals,
    snapshot: snapshotSummary(result.snapshot),
    snapshots: result.state.snapshots.map(snapshotSummary),
    sourceTotals: result.snapshot.sourceTotals,
    message: `Fetched ${result.snapshot.sourceTotals.labels} latest labels across ${result.snapshot.sourceTotals.entries} Strapi entries.`,
  });
});

app.get("/api/snapshots/:id/:changeType", async (req, res) => {
  const state = await readState();
  const snapshot = state.snapshots.find((item) => item.id === req.params.id);
  if (!snapshot) return res.status(404).json({ error: "Snapshot not found" });
  const key = text(req.params.changeType).toLowerCase();
  const rows = snapshot.changes?.[key] ?? [];
  res.json({
    snapshot: snapshotSummary(snapshot),
    changeType: key,
    rows: rows.map((row) => compact(row)),
    total: rows.length,
    totals: summarize(rows, rows),
  });
});

app.post("/api/dictionary/process", express.raw({ type: "application/octet-stream", limit: "50mb" }), async (req, res) => {
  try {
    const encodedName = text(req.headers["x-file-name"]);
    const originalFileName = decodeURIComponent(encodedName || "labels.xlsx");
    if (!/\.(xlsx|xlsm)$/i.test(originalFileName)) {
      return res.status(400).json({ error: "Upload an .xlsx or .xlsm Excel workbook." });
    }
    if (!Buffer.isBuffer(req.body) || !req.body.length) {
      return res.status(400).json({ error: "The uploaded workbook is empty." });
    }
    const state = await readState();
    const result = await processDictionaryWorkbook(req.body, originalFileName, state);
    res.json({
      ...result,
      uploadedFile: {
        name: originalFileName,
        size: req.body.length,
        uploadedAt: nowIso(),
      },
      dictionarySource: state.lastSnapshot ? snapshotSummary(state.lastSnapshot) : null,
    });
  } catch (error) {
    res.status(422).json({ error: error.message || "The workbook could not be processed." });
  }
});

app.get("/api/dictionary/jobs/:jobId/:fileType", async (req, res) => {
  const jobId = text(req.params.jobId);
  const fileType = text(req.params.fileType);
  if (!/^[0-9a-f-]{36}$/i.test(jobId)) return res.status(400).json({ error: "Invalid dictionary job." });
  const suffixes = {
    updated: "dictionary-updated.xlsx",
    exceptions: "dictionary-exceptions.xlsx",
    summary: "dictionary-summary.xlsx",
  };
  const suffix = suffixes[fileType];
  if (!suffix) return res.status(404).json({ error: "Dictionary download not found." });
  const jobDir = path.join(DICTIONARY_EXPORT_DIR, jobId);
  const files = await fs.readdir(jobDir).catch(() => []);
  const fileName = files.find((name) => name.endsWith(suffix));
  if (!fileName) return res.status(404).json({ error: "Dictionary download has expired or is unavailable." });
  res.download(path.join(jobDir, fileName), fileName);
});

async function buildWorkbook(rows) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "AK stc labels";
  workbook.created = new Date();
  const sheet = workbook.addWorksheet("All Labels");
  sheet.columns = [
    { header: "Feature Key", key: "featureKey", width: 28 },
    { header: "Label Key", key: "labelKey", width: 32 },
    { header: "English Text", key: "englishText", width: 54 },
    { header: "Arabic Text", key: "arabicText", width: 54 },
    { header: "Status", key: "status", width: 16 },
  ];
  sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  sheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4F008C" } };
  for (const row of rows) {
    sheet.addRow({
      featureKey: row.featureKey,
      labelKey: row.labelKey,
      englishText: row.englishText,
      arabicText: row.arabicText,
      status: row.changeType || "Existing",
    });
  }
  sheet.views = [{ state: "frozen", ySplit: 1 }];
  sheet.autoFilter = "A1:E1";
  return workbook;
}

function addHistorySheet(workbook, historyRows) {
  const sheet = workbook.addWorksheet("Change History");
  sheet.columns = [
    { header: "Date Time", key: "changedAt", width: 24 },
    { header: "Change Type", key: "changeType", width: 16 },
    { header: "Status", key: "status", width: 14 },
    { header: "Changed By", key: "changedBy", width: 28 },
    { header: "Feature Key", key: "featureKey", width: 28 },
    { header: "Label Key", key: "labelKey", width: 32 },
    { header: "English Text", key: "englishText", width: 54 },
    { header: "Arabic Text", key: "arabicText", width: 54 },
    { header: "Change Summary", key: "changeSummary", width: 32 },
    { header: "Label ID", key: "labelId", width: 22 },
    { header: "Component", key: "component", width: 24 },
    { header: "Component ID", key: "componentId", width: 16 },
    { header: "Entry ID", key: "entryId", width: 14 },
    { header: "Document ID", key: "documentId", width: 28 },
    { header: "Source", key: "source", width: 30 },
  ];
  sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } };
  sheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4F008C" } };
  for (const row of historyRows) {
    sheet.addRow({
      changedAt: row.changedAt,
      changeType: row.changeType,
      status: row.status,
      changedBy: row.changedBy,
      featureKey: row.featureKey,
      labelKey: row.labelKey,
      englishText: row.englishText,
      arabicText: row.arabicText,
      changeSummary: row.changeSummary,
      labelId: row.labelId,
      component: row.component,
      componentId: row.componentId,
      entryId: row.entryId,
      documentId: row.documentId,
      source: row.source,
    });
  }
  sheet.views = [{ state: "frozen", ySplit: 1 }];
  sheet.autoFilter = "A1:O1";
}

function excelDate(value) {
  if (!value) return null;
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

function languageIssueFill(issueType) {
  if (issueType.includes("missing")) return "FFFFF2CC";
  if (issueType.includes("contains") || issueType.includes("full English")) return "FFFFE4E9";
  if (issueType.includes("case")) return "FFF1E8FF";
  return "FFE7F5FA";
}

function addLanguageIssuesSheet(workbook, issueRows) {
  const sheet = workbook.addWorksheet("Language Issues");
  sheet.columns = [
    { header: "Issue Type", key: "issueType", width: 40 },
    { header: "Occurrences", key: "occurrenceCount", width: 14 },
    { header: "Issue Group ID", key: "groupId", width: 22 },
    { header: "Issue Details", key: "details", width: 56 },
    { header: "Observed Variants", key: "variants", width: 56 },
    { header: "English Text", key: "englishText", width: 54 },
    { header: "Arabic Text", key: "arabicText", width: 54 },
    { header: "Feature Name", key: "featureName", width: 28 },
    { header: "Feature Key", key: "featureKey", width: 28 },
    { header: "Label Key", key: "labelKey", width: 32 },
    { header: "Module", key: "module", width: 24 },
    { header: "Screen Name", key: "screenName", width: 24 },
    { header: "Status", key: "status", width: 16 },
    { header: "Change Type", key: "changeType", width: 16 },
    { header: "Label ID", key: "labelId", width: 22 },
    { header: "Component", key: "component", width: 28 },
    { header: "Component ID", key: "componentId", width: 16 },
    { header: "Entry ID", key: "entryId", width: 14 },
    { header: "Document ID", key: "documentId", width: 30 },
    { header: "First Seen", key: "firstSeenAt", width: 21 },
    { header: "Last Modified", key: "lastModifiedAt", width: 21 },
    { header: "Comments", key: "comments", width: 32 },
  ];

  const header = sheet.getRow(1);
  header.height = 28;
  header.font = { bold: true, color: { argb: "FFFFFFFF" } };
  header.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4F008C" } };
  header.alignment = { vertical: "middle", wrapText: true };

  for (const issue of issueRows) {
    const row = sheet.addRow({
      issueType: issue.issueType,
      occurrenceCount: issue.occurrenceCount,
      groupId: issue.groupId,
      details: issue.details,
      variants: (issue.variants || []).join(" | "),
      englishText: issue.englishText,
      arabicText: issue.arabicText,
      featureName: issue.featureName,
      featureKey: issue.featureKey,
      labelKey: issue.labelKey,
      module: issue.module,
      screenName: issue.screenName,
      status: issue.status,
      changeType: issue.changeType,
      labelId: issue.labelId,
      component: issue.component,
      componentId: issue.componentId,
      entryId: issue.entryId,
      documentId: issue.documentId,
      firstSeenAt: excelDate(issue.firstSeenAt),
      lastModifiedAt: excelDate(issue.lastModifiedAt),
      comments: issue.comments,
    });
    row.alignment = { vertical: "top", wrapText: true };
    row.getCell("issueType").fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: languageIssueFill(issue.issueType) },
    };
    row.getCell("issueType").font = { bold: true, color: { argb: "FF302444" } };
    row.getCell("occurrenceCount").numFmt = "#,##0";
    row.getCell("arabicText").alignment = { vertical: "top", wrapText: true, horizontal: "right", readingOrder: "rtl" };
    row.getCell("firstSeenAt").numFmt = "yyyy-mm-dd hh:mm";
    row.getCell("lastModifiedAt").numFmt = "yyyy-mm-dd hh:mm";
  }

  sheet.views = [{ state: "frozen", ySplit: 1 }];
  sheet.autoFilter = "A1:V1";
}

app.get("/api/download/xlsx", async (req, res) => {
  const { state, filtered } = await labelData(req.query);
  await fs.mkdir(EXPORT_DIR, { recursive: true });
  const workbook = await buildWorkbook(filtered);
  addHistorySheet(workbook, filterHistory(state.changeHistory ?? [], req.query));
  addLanguageIssuesSheet(workbook, languageConsistencyRecords(rowsFromState(state, false)));
  const fileName = `stc-labels-filtered-${filtered.length}-rows-${Date.now()}.xlsx`;
  const filePath = path.join(EXPORT_DIR, fileName);
  await workbook.xlsx.writeFile(filePath);
  res.download(filePath, fileName);
});

app.get([
  "/dashboard",
  "/labels",
  "/dictionary-generator",
  "/snapshots",
  "/language-issues",
  "/reports",
  "/settings",
  "/content-history",
  "/api-status",
  "/compare",
  "/language",
  "/audit",
  "/dahbaord",
], (_req, res) => {
  res.sendFile(path.join(__dirname, "public-labels", "index.html"));
});

const port = Number(process.env.PORT) || 4555;
app.listen(port, () => {
  console.log(`AK stc labels dashboard running at http://localhost:${port}`);
});
