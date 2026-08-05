const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

const sourceWorkbook =
  "C:/Users/thear/Downloads/stc-labels-filtered-5429-rows-1784034454063 (1) (1).xlsx";
const diffPath =
  "C:/Users/thear/OneDrive/Documents/New project/outputs/live-label-diff-20260720.json";
const outputDir = "C:/Users/thear/OneDrive/Documents/New project/outputs";

const auditFiles = [
  "bulk-ui-update-lineManagement-20260720-v2.json",
  "bulk-ui-update-Qitaf-20260720.json",
  "bulk-ui-safe-update-audit-remaining-20260720.json",
  "bulk-ui-safe-update-audit-remaining2-20260720.json",
  "bulk-ui-safe-update-audit-remaining3-20260720.json",
  "bulk-ui-safe-update-audit-remaining5-20260720.json",
].map((file) => path.join(outputDir, file));

const manualVerifiedRows = new Set([
  4114, // gnlPlans / uthAgeRestrictionMessage, verified published on Jul 20
]);

const STATUS_UPDATED = "updated";
const STATUS_NO_CHANGE = "no changes";
const STATUS_NEW_CONFIRM =
  "New label is there in the 68 entries need to confirm the new label";

function cloneStyle(style) {
  return style ? JSON.parse(JSON.stringify(style)) : {};
}

function collectUpdatedRows() {
  const updated = new Set(manualVerifiedRows);

  function addApplied(applied) {
    for (const row of applied || []) {
      if (row.rowNumber) updated.add(Number(row.rowNumber));
    }
  }

  for (const file of auditFiles) {
    if (!fs.existsSync(file)) continue;
    const json = JSON.parse(fs.readFileSync(file, "utf8"));
    const entries = Array.isArray(json) ? json : [json];
    for (const entry of entries) {
      if (entry.ok === false) continue;
      addApplied(entry.applied);
      addApplied(entry.result?.applied);
    }
  }

  return updated;
}

(async () => {
  const diff = JSON.parse(fs.readFileSync(diffPath, "utf8"));
  const diffByRow = new Map(diff.rows.map((row) => [Number(row.rowNumber), row]));
  const updatedRows = collectUpdatedRows();

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(sourceWorkbook);
  const sheet = workbook.getWorksheet("All Labels") || workbook.worksheets[0];
  const rowCount = sheet.rowCount;

  const statusCol = 5;
  const detailCol = 6;
  sheet.getCell(1, statusCol).value = "Update Status";
  sheet.getCell(1, detailCol).value = "Review Detail";

  for (const col of [statusCol, detailCol]) {
    sheet.getCell(1, col).style = cloneStyle(sheet.getCell(1, 4).style);
    sheet.getCell(1, col).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF4F008C" },
    };
    sheet.getCell(1, col).font = {
      ...(sheet.getCell(1, col).font || {}),
      bold: true,
      color: { argb: "FFFFFFFF" },
    };
  }

  const counts = {};
  const detailCounts = {};

  for (let rowNumber = 2; rowNumber <= rowCount; rowNumber += 1) {
    const diffRow = diffByRow.get(rowNumber);
    let status;
    let detail;

    if (updatedRows.has(rowNumber)) {
      status = STATUS_UPDATED;
      detail = "Published and verified in Strapi";
    } else if (diffRow?.status === "No Change") {
      status = STATUS_NO_CHANGE;
      detail = "Excel matches current live Strapi value";
    } else {
      status = STATUS_NEW_CONFIRM;
      detail = diffRow
        ? `${diffRow.status}: ${diffRow.reason || "Needs confirmation"}`
        : "Needs confirmation";
    }

    const statusCell = sheet.getCell(rowNumber, statusCol);
    const detailCell = sheet.getCell(rowNumber, detailCol);
    statusCell.value = status;
    detailCell.value = detail;
    statusCell.style = cloneStyle(sheet.getCell(rowNumber, 4).style);
    detailCell.style = cloneStyle(sheet.getCell(rowNumber, 4).style);

    if (status === STATUS_UPDATED) {
      statusCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8F7EF" } };
      statusCell.font = { ...(statusCell.font || {}), color: { argb: "FF087443" }, bold: true };
    } else if (status === STATUS_NEW_CONFIRM) {
      statusCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF4DB" } };
      statusCell.font = { ...(statusCell.font || {}), color: { argb: "FF8A5A00" }, bold: true };
    }

    counts[status] = (counts[status] || 0) + 1;
    detailCounts[detail.split(":")[0]] = (detailCounts[detail.split(":")[0]] || 0) + 1;
  }

  sheet.getColumn(statusCol).width = 64;
  sheet.getColumn(detailCol).width = 46;
  sheet.views = [{ state: "frozen", ySplit: 1 }];
  sheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: rowCount, column: detailCol },
  };

  fs.mkdirSync(outputDir, { recursive: true });
  const timestamp = new Date().toISOString().replace(/[-:TZ.]/g, "").slice(0, 14);
  const outputPath = path.join(outputDir, `stc-labels-three-status-${timestamp}.xlsx`);
  await workbook.xlsx.writeFile(outputPath);

  const verify = new ExcelJS.Workbook();
  await verify.xlsx.readFile(outputPath);
  const verifySheet = verify.getWorksheet(sheet.name);
  const verifyCounts = {};
  verifySheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const value = row.getCell(statusCol).value || "";
    verifyCounts[value] = (verifyCounts[value] || 0) + 1;
  });

  console.log(JSON.stringify({
    outputPath,
    rowCount: verifySheet.rowCount - 1,
    counts: verifyCounts,
    updatedRowCount: updatedRows.size,
    headers: verifySheet.getRow(1).values.slice(1, 7),
  }, null, 2));
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
