import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const sourceWorkbook =
  "C:/Users/thear/Downloads/stc-labels-filtered-5429-rows-1784034454063 (1) (1).xlsx";
const classificationPath =
  "C:/Users/thear/OneDrive/Documents/New project/outputs/bulk-label-classification.json";
const outputDir = "C:/Users/thear/OneDrive/Documents/New project/outputs";

const verifiedUpdatedRows = new Set([
  2947, // estore / brand
  2952, // estore / memory
  2975, // estore / digitalVouchers
  3342, // autoPay / debitCreditCard
  3985, // gnlPlans / activationMethodDescription
  4065, // gnlPlans / dueToday
  4072, // gnlPlans / youPayNow
  4111, // gnlPlans / uthEligibilityInfo
  4608, // dealerApp / recharge
]);

const classification = JSON.parse(await fs.readFile(classificationPath, "utf8"));
const byRow = new Map(classification.classified.map((row) => [row.rowNumber, row]));

function workbookStatus(rowNumber) {
  if (verifiedUpdatedRows.has(rowNumber)) return "Updated";

  const item = byRow.get(rowNumber);
  if (!item) return "Not Included in This Update";
  if (item.status === "No Change") return "No Change";
  if (item.status === "New Field") return "New Field";
  return "Not Included in This Update";
}

const input = await FileBlob.load(sourceWorkbook);
const workbook = await SpreadsheetFile.importXlsx(input);
const sheet = workbook.worksheets.getItem("All Labels");

const rowCount = classification.classified.length + 1;
const statuses = [["Update Status"]];
for (let rowNumber = 2; rowNumber <= rowCount; rowNumber += 1) {
  statuses.push([workbookStatus(rowNumber)]);
}

sheet.getRange(`E1:E${rowCount}`).copyFrom(sheet.getRange(`D1:D${rowCount}`), "formats");
sheet.getRange(`E1:E${rowCount}`).values = statuses;
sheet.getRange("E1").format = {
  fill: "#4F008C",
  font: { bold: true, color: "#FFFFFF" },
};
sheet.getRange(`E2:E${rowCount}`).dataValidation = {
  rule: {
    type: "list",
    values: ["Updated", "No Change", "New Field", "Not Included in This Update"],
  },
};
sheet.getRange(`A1:E${rowCount}`).format.autofitColumns();
sheet.getRange("A1:E1").format.borders = {
  preset: "outside",
  style: "thin",
  color: "#D8CBEA",
};
sheet.freezePanes.freezeRows(1);

const timestamp = new Date().toISOString().replace(/[-:TZ.]/g, "").slice(0, 14);
await fs.mkdir(outputDir, { recursive: true });
const outputPath = path.join(outputDir, `stc-labels-update-status-${timestamp}.xlsx`);
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);

const counts = {};
for (const [status] of statuses.slice(1)) counts[status] = (counts[status] || 0) + 1;

const inspect = await workbook.inspect({
  kind: "table",
  sheetId: "All Labels",
  range: "A1:E8",
  include: "values",
  tableMaxRows: 8,
  tableMaxCols: 5,
});

console.log(JSON.stringify({ outputPath, rowCount: rowCount - 1, counts, preview: inspect.ndjson }, null, 2));
