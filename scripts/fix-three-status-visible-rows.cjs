const ExcelJS = require("exceljs");
const path = require("path");

const input =
  "C:/Users/thear/OneDrive/Documents/New project/outputs/stc-labels-three-status-20260720064353.xlsx";
const outputDir = "C:/Users/thear/OneDrive/Documents/New project/outputs";

(async () => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(input);
  const sheet = workbook.getWorksheet("All Labels") || workbook.worksheets[0];

  for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    row.hidden = false;
    row.collapsed = false;
    row.outlineLevel = 0;
  }

  // Keep the data plain-visible. User can reapply filters manually in Excel.
  sheet.autoFilter = null;
  sheet.views = [{ state: "frozen", ySplit: 1 }];

  const outputPath = path.join(outputDir, "stc-labels-three-status-all-rows-visible.xlsx");
  await workbook.xlsx.writeFile(outputPath);

  const verify = new ExcelJS.Workbook();
  await verify.xlsx.readFile(outputPath);
  const verifySheet = verify.getWorksheet(sheet.name);
  let hidden = 0;
  verifySheet.eachRow((row) => {
    if (row.hidden) hidden += 1;
  });

  console.log(JSON.stringify({
    outputPath,
    sheet: verifySheet.name,
    rowCount: verifySheet.rowCount,
    dataRows: verifySheet.rowCount - 1,
    hiddenRows: hidden,
    headers: verifySheet.getRow(1).values.slice(1, 7),
  }, null, 2));
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
