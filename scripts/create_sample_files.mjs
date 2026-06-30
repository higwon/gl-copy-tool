import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, "..");
const outputDir = path.join(rootDir, "samples");

const reviewHeaders = ["날짜", "계정코드", "차변(EUR)", "대변(EUR)", "거래처명", "적요", "Review Check"];
const exportHeaders = ["적요", "거래처명", "대변(EUR)", "차변(EUR)", "계정코드", "날짜"];
const templateRows = [
  ["2026-06-01", "400100", 1000, null, "Old Customer", "Old sales entry", null],
  ["2026-06-02", "500200", null, 250, "Old Vendor", "Old expense entry", null],
  ["2026-06-03", "210300", null, 100, "Old Vendor", "Old accrual entry", null],
];
const exportRows = [
  ["ERP sales invoice A-1001", "Customer A", null, 1250000, "400100", "2026-06-10"],
  ["ERP travel expense T-210", "Vendor B", 340000, null, "500200", "2026-06-11"],
  ["ERP accrued liability L-300", "Vendor C", 780000, null, "210300", "2026-06-12"],
  ["ERP bank receipt B-090", "Customer D", null, 450000, "110100", "2026-06-13"],
];

function styleSheet(sheet, title, rangeAddress) {
  sheet.showGridLines = false;
  sheet.getRange("A1:G1").merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A1").format = {
    fill: "#1F4E78",
    font: { bold: true, color: "#FFFFFF", size: 14 },
  };
  sheet.getRange("A3:G3").format = {
    fill: "#D9EAF7",
    font: { bold: true, color: "#1F2937" },
  };
  sheet.getRange(rangeAddress).format.borders = {
    preset: "all",
    style: "thin",
    color: "#C8D2DC",
  };
  sheet.getRange("A:A").format.columnWidth = 13;
  sheet.getRange("B:B").format.columnWidth = 14;
  sheet.getRange("C:D").format.columnWidth = 14;
  sheet.getRange("E:E").format.columnWidth = 18;
  sheet.getRange("F:F").format.columnWidth = 34;
  sheet.getRange("G:G").format.columnWidth = 18;
  sheet.getRange("A4:A20").format.numberFormat = "yyyy-mm-dd";
  sheet.getRange("C4:D20").format.numberFormat = "#,##0";
  sheet.freezePanes.freezeRows(3);
}

async function saveWorkbook(workbook, fileName) {
  const xlsx = await SpreadsheetFile.exportXlsx(workbook);
  await xlsx.save(path.join(outputDir, fileName));
}

async function buildTemplate() {
  const workbook = Workbook.create();
  const cover = workbook.worksheets.add("Instructions");
  cover.showGridLines = false;
  cover.getRange("A1:D1").merge();
  cover.getRange("A1").values = [["Sample Review Template"]];
  cover.getRange("A1").format = {
    fill: "#1F4E78",
    font: { bold: true, color: "#FFFFFF", size: 14 },
  };
  cover.getRange("A3").values = [["Use this file as the Review template input."]];
  cover.getRange("A4").values = [["The app copies ERP Export data into the 2. GL Input sheet."]];
  cover.getRange("A5").values = [["Column G contains formulas and should remain unchanged."]];
  cover.getRange("A:A").format.columnWidth = 72;

  const sheet = workbook.worksheets.add("2. GL Input");
  styleSheet(sheet, "Review Template - 2. GL Input", "A3:G6");
  sheet.getRange("A3:G6").values = [reviewHeaders, ...templateRows];
  sheet.getRange("G4").formulas = [["=IF(C4>0,\"Debit\",IF(D4>0,\"Credit\",\"Check\"))"]];
  sheet.getRange("G4:G6").fillDown();

  await saveWorkbook(workbook, "Review_Template_Sample.xlsx");
}

async function buildExport() {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("ERP Export");
  styleSheet(sheet, "ERP Export - First Sheet", "A3:F7");
  sheet.getRange("A3:F7").values = [exportHeaders, ...exportRows];
  await saveWorkbook(workbook, "ERP_Export_Sample.xlsx");
}

await fs.mkdir(outputDir, { recursive: true });
await buildTemplate();
await buildExport();

console.log(`Sample files created in ${outputDir}`);
