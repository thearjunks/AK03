const fs = require("fs");
const https = require("https");
const ExcelJS = require("exceljs");

const sourceWorkbook =
  "C:/Users/thear/Downloads/stc-labels-filtered-5429-rows-1784034454063 (1) (1).xlsx";
const adminInventoryPath =
  "C:/Users/thear/OneDrive/Documents/New project/outputs/strapi-admin-68-entries-20260720.json";
const outputPath =
  "C:/Users/thear/OneDrive/Documents/New project/outputs/live-label-diff-20260720.json";

function norm(value) {
  return String(value ?? "").trim().replace(/\s+/g, " ");
}

function getJson(url) {
  return new Promise((resolve, reject) => {
    https
      .get(
        url,
        {
          headers: {
            Accept: "application/json",
            "User-Agent": "AK-stc-label-live-diff/1.0",
          },
        },
        (res) => {
          let body = "";
          res.setEncoding("utf8");
          res.on("data", (chunk) => (body += chunk));
          res.on("end", () => {
            try {
              resolve(JSON.parse(body));
            } catch (error) {
              reject(new Error(`${res.statusCode}: ${body.slice(0, 300)}`));
            }
          });
        },
      )
      .on("error", reject);
  });
}

async function fetchLive() {
  const live = [];
  let page = 1;
  while (true) {
    const url = `https://content.stc.com.kw/api/stc-labels?populate=*&publicationState=preview&pagination[page]=${page}&pagination[pageSize]=100`;
    const json = await getJson(url);
    live.push(...(json.data || []));
    const pagination = json.meta?.pagination;
    if (!pagination || page >= pagination.pageCount) break;
    page += 1;
  }
  return live;
}

function indexLive(entries) {
  const map = new Map();
  for (const entry of entries) {
    const featureKey = norm(entry.featureKey);
    for (const row of entry.KeyValuePairComponent || []) {
      const key = `${featureKey}\u0000${norm(row.key)}`;
      if (!map.has(key)) map.set(key, []);
      map.get(key).push({
        entryId: entry.id,
        documentId: entry.documentId,
        featureName: entry.featureName,
        featureKey,
        labelKey: row.key,
        enValue: row.enValue ?? "",
        arValue: row.arValue ?? "",
        componentId: row.id,
        updatedAt: entry.updatedAt,
        publishedAt: entry.publishedAt,
      });
    }
  }
  return map;
}

(async () => {
  const adminEntries = JSON.parse(fs.readFileSync(adminInventoryPath, "utf8"));
  const adminFeatureKeys = new Set(adminEntries.map((e) => norm(e.featureKey)).filter(Boolean));
  const adminByFeature = new Map(adminEntries.map((e) => [norm(e.featureKey), e]));
  const liveEntries = await fetchLive();
  const liveMap = indexLive(liveEntries);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(sourceWorkbook);
  const ws = workbook.getWorksheet("All Labels") || workbook.worksheets[0];

  const rows = [];
  for (let rowNumber = 2; rowNumber <= ws.rowCount; rowNumber += 1) {
    const featureKey = norm(ws.getCell(rowNumber, 1).value);
    const labelKey = norm(ws.getCell(rowNumber, 2).value);
    const englishText = norm(ws.getCell(rowNumber, 3).value);
    const arabicText = norm(ws.getCell(rowNumber, 4).value);
    const key = `${featureKey}\u0000${labelKey}`;
    const matches = liveMap.get(key) || [];
    let status = "Not Included in This Update";
    let reason = "";
    let live = null;

    if (!featureKey || !labelKey) {
      reason = "Missing Feature Key or Label Key in Excel";
    } else if (!adminFeatureKeys.has(featureKey)) {
      status = "New Field";
      reason = "Feature Key is not present in the 68-entry admin list";
    } else if (!matches.length) {
      status = "New Field";
      reason = "Label Key is not present in live public API for this Feature Key";
    } else if (matches.length > 1) {
      reason = "Duplicate live Feature Key + Label Key; manual review required";
      live = matches;
    } else {
      live = matches[0];
      const enSame = norm(live.enValue) === englishText;
      const arSame = norm(live.arValue) === arabicText;
      if (enSame && arSame) {
        status = "No Change";
        reason = "Excel matches live Strapi values";
      } else {
        status = "Update Needed";
        reason = [
          !enSame ? "English Text differs" : null,
          !arSame ? "Arabic Text differs" : null,
        ].filter(Boolean).join(", ");
      }
    }

    rows.push({
      rowNumber,
      featureKey,
      labelKey,
      englishText,
      arabicText,
      status,
      reason,
      adminEntry: adminByFeature.get(featureKey) || null,
      live,
    });
  }

  const counts = rows.reduce((acc, row) => {
    acc[row.status] = (acc[row.status] || 0) + 1;
    return acc;
  }, {});
  const updateNeededByFeature = rows
    .filter((r) => r.status === "Update Needed")
    .reduce((acc, row) => {
      acc[row.featureKey] = (acc[row.featureKey] || 0) + 1;
      return acc;
    }, {});
  const publicFeatures = new Set(liveEntries.map((e) => norm(e.featureKey)));
  const adminOnlyEntries = adminEntries.filter((e) => !publicFeatures.has(norm(e.featureKey)));

  const output = {
    generatedAt: new Date().toISOString(),
    sourceWorkbook,
    adminEntryCount: adminEntries.length,
    publicApiEntryCount: liveEntries.length,
    publicApiLabelCount: liveEntries.reduce((sum, e) => sum + (e.KeyValuePairComponent?.length || 0), 0),
    adminStatusCounts: adminEntries.reduce((acc, e) => {
      acc[e.status || "Unknown"] = (acc[e.status || "Unknown"] || 0) + 1;
      return acc;
    }, {}),
    adminOnlyEntries,
    counts,
    updateNeededByFeature,
    rows,
  };

  fs.writeFileSync(outputPath, JSON.stringify(output, null, 2), "utf8");
  console.log(JSON.stringify({
    outputPath,
    adminEntryCount: output.adminEntryCount,
    publicApiEntryCount: output.publicApiEntryCount,
    publicApiLabelCount: output.publicApiLabelCount,
    adminStatusCounts: output.adminStatusCounts,
    adminOnlyEntries,
    counts,
    updateNeededByFeature,
  }, null, 2));
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
