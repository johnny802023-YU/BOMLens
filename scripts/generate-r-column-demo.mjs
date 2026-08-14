import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import ExcelJS from "exceljs";
import { compareBom, parseCompanyBomMatrix } from "../app/bom-logic.ts";
import { buildBomReportWithOriginals } from "../app/export-report.ts";
import { buildBomHtmlReport } from "../app/export-html-report.ts";
import { mapCustomerBom, parseCustomerBomMatrix } from "../app/supplemental-logic.ts";

const outputDir = path.resolve("outputs/bomlens-r-column-demo");
await mkdir(outputDir, { recursive: true });

const companyHeader = Array(18).fill("");
companyHeader[0] = "項次";
companyHeader[1] = "主件料號";
companyHeader[2] = "品名";
companyHeader[4] = "組成用量";
companyHeader[5] = "插件位置";
companyHeader[10] = "製造廠商";
companyHeader[15] = "製造廠商料號";
companyHeader[17] = "對應客戶料號";

function companyRow(ref, part, description, qty, positions, manufacturer, mpn, customerPartNumbers) {
  const row = Array(18).fill("");
  row[0] = ref;
  row[1] = part;
  row[2] = description;
  row[4] = `${qty}.0/1`;
  row[5] = positions;
  row[10] = manufacturer;
  row[15] = mpn;
  row[17] = customerPartNumbers;
  return row;
}

const beforeMatrix = [
  companyHeader,
  companyRow("00A", "DEMO-PART-A01", "R欄版本變更示範", 2, "U1,U2", "DEMO MAKER A", "MPN-A-100", "CUST-100"),
  companyRow("00B", "DEMO-PART-B01", "新版 R欄漏維護示範", 1, "U3", "DEMO MAKER B", "MPN-B-200", "CUST-200"),
  companyRow("00C", "DEMO-PART-C01", "新版 R欄不一致示範", 1, "U4", "DEMO MAKER C", "MPN-C-300", "CUST-OLD"),
  companyRow("00D", "DEMO-PART-D01", "正常一致與位置差異", 1, "U5", "DEMO MAKER D", "MPN-D-400", "CUST-400"),
  companyRow("00E", "DEMO-PART-E01", "新版完全移除", 1, "U8", "DEMO MAKER E", "MPN-E-500", "CUST-500"),
];

const afterMatrix = [
  companyHeader,
  companyRow("00A", "DEMO-PART-A01", "R欄版本變更示範", 2, "U1,U2", "DEMO MAKER A", "MPN-A-100", "CUST-101"),
  companyRow("00B", "DEMO-PART-B01", "新版 R欄漏維護示範", 1, "U3", "DEMO MAKER B", "MPN-B-200", ""),
  companyRow("00C", "DEMO-PART-C01", "新版 R欄不一致示範", 1, "U4", "DEMO MAKER C", "MPN-C-300", "CUST-OLD"),
  companyRow("00D", "DEMO-PART-D01", "正常一致與位置差異", 2, "U5,U7", "DEMO MAKER D", "MPN-D-400", "CUST-400"),
  companyRow("00F", "DEMO-PART-F01", "新版完全新料", 1, "U6", "DEMO MAKER F", "MPN-F-600", "CUST-600"),
];

const customerHeader = Array(17).fill("");
customerHeader[1] = "PartNumber";
customerHeader[6] = "MfgPNos";
customerHeader[16] = "ReferenceDesignator";

function customerRow(customerPartNumber, mpn, positions) {
  const row = Array(17).fill("");
  row[1] = customerPartNumber;
  row[6] = mpn;
  row[16] = positions;
  return row;
}

const customerBeforeMatrix = [
  customerHeader,
  customerRow("CUST-100", "MPN-A-100", "U2|U1"),
  customerRow("CUST-200", "MPN-B-200", "U3"),
  customerRow("CUST-OLD", "MPN-C-300", "U4"),
  customerRow("CUST-400", "MPN-D-400", "U5"),
  customerRow("CUST-500", "MPN-E-500", "U8"),
];

const customerAfterMatrix = [
  customerHeader,
  customerRow("CUST-101", "MPN-A-100", "U1|U2"),
  customerRow("CUST-200", "MPN-B-200", "U3"),
  customerRow("CUST-300", "MPN-C-300", "U4"),
  customerRow("CUST-400", "MPN-D-400", "U7|U5"),
  customerRow("CUST-600", "MPN-F-600", "U6"),
];

function styleWorksheet(sheet, columnCount) {
  sheet.views = [{ state: "frozen", ySplit: 1, showGridLines: false }];
  sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: sheet.rowCount, column: columnCount } };
  sheet.getRow(1).height = 28;
  sheet.getRow(1).eachCell((cell) => {
    cell.font = { name: "Microsoft JhengHei", bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2F6BCE" } };
    cell.alignment = { vertical: "middle", horizontal: "center" };
  });
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    row.height = 24;
    row.eachCell((cell) => {
      cell.font = { name: "Microsoft JhengHei", size: 10, color: { argb: "FF344054" } };
      cell.alignment = { vertical: "middle", wrapText: true };
      cell.border = { bottom: { style: "thin", color: { argb: "FFE5E9EF" } } };
    });
  }
}

async function createInputWorkbook(matrix, fileName, sheetName, widths) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(sheetName);
  matrix.forEach((values) => sheet.addRow(values));
  styleWorksheet(sheet, matrix[0].length);
  widths.forEach((width, index) => {
    sheet.getColumn(index + 1).width = width;
  });
  const filePath = path.join(outputDir, fileName);
  await workbook.xlsx.writeFile(filePath);
  const buffer = await workbook.xlsx.writeBuffer();
  return { filePath, buffer };
}

const companyWidths = [10, 22, 25, 4, 12, 18, 4, 4, 4, 4, 20, 4, 4, 4, 4, 22, 4, 28];
const customerWidths = [4, 20, 4, 4, 4, 4, 24, 4, 4, 4, 4, 4, 4, 4, 4, 4, 28];
const beforeFile = await createInputWorkbook(beforeMatrix, "Demo_舊版公司BOM.xlsx", "舊版 BOM", companyWidths);
const afterFile = await createInputWorkbook(afterMatrix, "Demo_新版公司BOM.xlsx", "新版 BOM", companyWidths);
await createInputWorkbook(customerBeforeMatrix, "Demo_舊版客戶BOM.xlsx", "舊版客戶 BOM", customerWidths);
await createInputWorkbook(customerAfterMatrix, "Demo_新版客戶BOM.xlsx", "新版客戶 BOM", customerWidths);

const beforeItems = parseCompanyBomMatrix(beforeMatrix);
const afterItems = parseCompanyBomMatrix(afterMatrix);
const customerBefore = mapCustomerBom(beforeItems, parseCustomerBomMatrix(customerBeforeMatrix, 0, {
  customerPartNumber: 1,
  manufacturerParts: 6,
  positions: 16,
}));
const customerAfter = mapCustomerBom(afterItems, parseCustomerBomMatrix(customerAfterMatrix, 0, {
  customerPartNumber: 1,
  manufacturerParts: 6,
  positions: 16,
}));
const diffs = compareBom(customerBefore.items, customerAfter.items).filter((diff) => diff.kind !== "same");
const context = {
  customerBefore,
  customerAfter,
  mvaBefore: {
    smtTop: 126,
    smtBottom: 42,
    dipTop: 18,
    dipBottom: 6,
    included: [],
    excluded: [],
  },
  mvaAfter: {
    smtTop: 127,
    smtBottom: 42,
    dipTop: 18,
    dipBottom: 6,
    included: [],
    excluded: [],
  },
  originalBefore: {
    fileName: "Demo_舊版公司BOM.xlsx",
    sheetName: "舊版 BOM",
    data: beforeFile.buffer,
    matrix: beforeMatrix,
  },
  originalAfter: {
    fileName: "Demo_新版公司BOM.xlsx",
    sheetName: "新版 BOM",
    data: afterFile.buffer,
    matrix: afterMatrix,
  },
};

const report = await buildBomReportWithOriginals(
  diffs,
  "Demo_舊版公司BOM.xlsx",
  "Demo_新版公司BOM.xlsx",
  context,
);
await report.xlsx.writeFile(path.join(outputDir, "Demo_BOMLens_R欄客戶料號差異報告.xlsx"));
await writeFile(
  path.join(outputDir, "Demo_BOMLens_R欄客戶料號差異報告.html"),
  buildBomHtmlReport(diffs, "Demo_舊版公司BOM.xlsx", "Demo_新版公司BOM.xlsx", context),
  "utf8",
);

console.log(JSON.stringify({
  outputDir,
  diffCount: diffs.length,
  beforeMapping: customerBefore.counts,
  afterMapping: customerAfter.counts,
}, null, 2));
