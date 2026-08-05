const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

const sourceWorkbook =
  "C:/Users/thear/Downloads/stc-labels-filtered-5429-rows-1784034454063 (1) (1).xlsx";
const classificationPath =
  "C:/Users/thear/OneDrive/Documents/New project/outputs/bulk-label-classification.json";
const outputDir = "C:/Users/thear/OneDrive/Documents/New project/outputs";

const verifiedUpdatedRows = new Set([
  2947,
  2952,
  2975,
  3342,
  3985,
  4065,
  4072,
  4111,
  4608,
]);

const classification = JSON.parse(fs.readFileSync(classificationPath, "utf8"));
const byRow = new Map(classification.classified.map((row) => [row.rowNumber, row]));

function getStatus(rowNumber) {
  if (verifiedUpdatedRows.has(rowNumber)) return "Updated";
  const item = byRow.get(rowNumber);
  if (!item) return "Not Included in This Update";
  if (item.status === "No Change") return "No Change";
  if (item.status === "New Field") return "New Field";
  return "Not Included in This Update";
}

function cloneStyle(style) {
  return style ? JSON.parse(JSON.stringify(style)) : {};
}

(async () => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(sourceWorkbook);
  const sheet = workbook.getWorksheet("All Labels") || workbook.worksheets[0];

  const statusCol = 5;
  const rowCount = classification.classified.length + 1;

  sheet.getColumn(statusCol).header = "Update Status";
  sheet.getCell(1, statusCol).value = "Update Status";
  sheet.getCell(1, statusCol).style = cloneStyle(sheet.getCell(1, 4).style);
  sheet.getCell(1, statusCol).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF4F008C" },
  };
  sheet.getCell(1, statusCol).font = { ...(sheet.getCell(1, statusCol).font || {}), bold: true, color: { argb: "FFFFFFFF" } };

  const counts = {};
  for (let rowNumber = 2; rowNumber <= rowCount; rowNumber += 1) {
    const cell = sheet.getCell(rowNumber, statusCol);
    cell.style = cloneStyle(sheet.getCell(rowNumber, 4).style);
    const status = getStatus(rowNumber);
    cell.value = status;
    counts[status] = (counts[status] || 0) + 1;
  }

  sheet.getColumn(statusCol).width = 28;
  sheet.views = [{ state: "frozen", ySplit: 1 }];
  sheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: rowCount, column: statusCol },
  };

  for (let rowNumber = 2; rowNumber <= rowCount; rowNumber += 1) {
    const cell = sheet.getCell(rowNumber, statusCol);
    if (cell.value === "Updated") {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE8F7EF" } };
      cell.font = { ...(cell.font || {}), color: { argb: "FF087443" } };
    } else if (cell.value === "New Field") {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF4DB" } };
      cell.font = { ...(cell.font || {}), color: { argb: "FF8A5A00" } };
    } else if (cell.value === "Not Included in This Update") {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF3EEF8" } };
      cell.font = { ...(cell.font || {}), color: { argb: "FF4F008C" } };
    }
  }

  fs.mkdirSync(outputDir, { recursive: true });
  const timestamp = new Date().toISOString().replace(/[-:TZ.]/g, "").slice(0, 14);
  const outputPath = path.join(outputDir, `stc-labels-update-status-excel-${timestamp}.xlsx`);
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
    sheetName: sheet.name,
    rowCount: verifySheet.rowCount - 1,
    counts: verifyCounts,
    headers: verifySheet.getRow(1).values.slice(1, 6),
  }, null, 2));
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
