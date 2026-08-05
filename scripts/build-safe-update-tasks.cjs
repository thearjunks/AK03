const fs = require("fs");
const https = require("https");

const diffPath = "C:/Users/thear/OneDrive/Documents/New project/outputs/live-label-diff-20260720.json";
const outputPath = "C:/Users/thear/OneDrive/Documents/New project/outputs/safe-non-empty-update-tasks-20260720.json";

function getJson(url) {
  return new Promise((resolve, reject) => {
    https
      .get(url, { headers: { Accept: "application/json", "User-Agent": "AK-stc-label-indexer/1.0" } }, (res) => {
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
      })
      .on("error", reject);
  });
}

(async () => {
  const diff = JSON.parse(fs.readFileSync(diffPath, "utf8"));
  const raw = diff.rows.filter((r) => r.status === "Update Needed").filter((r) => {
    const liveEn = String(r.live?.enValue || "").trim();
    const liveAr = String(r.live?.arValue || "").trim();
    const desiredEn = String(r.englishText || "").trim();
    const desiredAr = String(r.arabicText || "").trim();
    const enDiff = liveEn !== desiredEn;
    const arDiff = liveAr !== desiredAr;
    return (!enDiff || desiredEn) && (!arDiff || desiredAr);
  });

  const byFeature = {};
  for (const row of raw) (byFeature[row.featureKey] ||= []).push(row);

  const output = {};
  for (const [featureKey, rows] of Object.entries(byFeature)) {
    const json = await getJson(`https://content.stc.com.kw/api/stc-labels?filters[featureKey][$eq]=${encodeURIComponent(featureKey)}&populate=*&publicationState=preview`);
    const entry = json.data?.[0];
    const components = entry?.KeyValuePairComponent || [];
    output[featureKey] = rows.map((row) => {
      let componentIndex = components.findIndex((c) => c.id === row.live?.componentId);
      if (componentIndex < 0) componentIndex = components.findIndex((c) => c.key === row.labelKey);
      return {
        featureKey,
        href: row.adminEntry?.href,
        rowNumber: row.rowNumber,
        labelKey: row.labelKey,
        desiredEn: row.englishText,
        desiredAr: row.arabicText,
        liveEn: row.live?.enValue || "",
        liveAr: row.live?.arValue || "",
        componentId: row.live?.componentId,
        componentIndex,
      };
    });
  }

  fs.writeFileSync(outputPath, JSON.stringify(output, null, 2), "utf8");
  console.log(JSON.stringify({
    outputPath,
    total: Object.values(output).reduce((sum, rows) => sum + rows.length, 0),
    byFeature: Object.fromEntries(Object.entries(output).map(([k, rows]) => [k, rows.length])),
    missingIndex: Object.values(output).flat().filter((r) => r.componentIndex < 0).length,
  }, null, 2));
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
