import ExcelJS from "exceljs";
import { type BomAlternative, type BomDiff, type ImportAudit } from "./bom-logic.ts";

export type ReportSource = { fileName: string; sheetName: string; importedAt: string; audit: ImportAudit };
export type OriginalBomSource = { fileName: string; sheetName: string; data: ArrayBuffer; matrix: unknown[][] };
export type ReportContext = {
  before?: ReportSource | null;
  after?: ReportSource | null;
  originalBefore?: OriginalBomSource | null;
  originalAfter?: OriginalBomSource | null;
};

const colors = {
  navy: "1F3A5F",
  blue: "2F6BCE",
  paleBlue: "EAF1FD",
  green: "16835D",
  paleGreen: "E8F7F0",
  red: "B14949",
  paleRed: "FFF0F0",
  amber: "9A651B",
  paleAmber: "FFF3DC",
  gray: "667085",
  paleGray: "F4F6F8",
  white: "FFFFFF",
  border: "D9E0E8",
};

function partLabel(part: BomAlternative) {
  return part.part || part.manufacturerPart || "未提供料號";
}

function listPartNumbers(parts: BomAlternative[]) {
  return parts.map(partLabel).join("\n");
}

function listManufacturerParts(parts: BomAlternative[]) {
  return parts.map((part) => part.manufacturerPart).filter(Boolean).join("\n");
}

function listManufacturers(parts: BomAlternative[]) {
  return parts.map((part) => part.manufacturerName).filter(Boolean).join("\n");
}

function primaryLabel(diff: BomDiff) {
  if (diff.primaryType === "componentAdded" || diff.primaryType === "substituteAdded") return "新增";
  if (diff.primaryType === "componentRemoved" || diff.primaryType === "substituteRemoved") return "刪除";
  return "變更";
}

function partChanges(diff: BomDiff) {
  return [
    ...diff.addedParts.map((part) => `＋ ${partLabel(part)}`),
    ...diff.removedParts.map((part) => `－ ${partLabel(part)}`),
  ].join("\n");
}

function positionChanges(diff: BomDiff) {
  return [
    ...diff.addedPositions.map((position) => `＋ ${position}`),
    ...diff.removedPositions.map((position) => `－ ${position}`),
    ...diff.replacementPositions.map((position) => `${position} 換料`),
  ].join("\n");
}

function fill(color: string): ExcelJS.Fill {
  return { type: "pattern", pattern: "solid", fgColor: { argb: color } };
}

function border(): Partial<ExcelJS.Borders> {
  const side = { style: "thin" as const, color: { argb: colors.border } };
  return { top: side, left: side, bottom: side, right: side };
}

function styleTitle(sheet: ExcelJS.Worksheet, range: string, title: string) {
  sheet.mergeCells(range);
  const cell = sheet.getCell(range.split(":")[0]);
  cell.value = title;
  cell.font = { name: "Microsoft JhengHei", size: 18, bold: true, color: { argb: colors.white } };
  cell.fill = fill(colors.navy);
  cell.alignment = { vertical: "middle", horizontal: "left" };
}

function styleHeader(row: ExcelJS.Row) {
  row.height = 32;
  row.eachCell((cell) => {
    cell.font = { name: "Microsoft JhengHei", size: 10, bold: true, color: { argb: colors.white } };
    cell.fill = fill(colors.blue);
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = border();
  });
}

function saveBuffer(buffer: ExcelJS.Buffer, filename: string) {
  const blob = new Blob([buffer as BlobPart], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function cloneExcelValue<T>(value: T): T {
  if (value == null) return value;
  return structuredClone(value);
}

function copyWorksheet(sourceWorkbook: ExcelJS.Workbook, source: ExcelJS.Worksheet, targetWorkbook: ExcelJS.Workbook, targetName: string) {
  const target = targetWorkbook.addWorksheet(targetName, {
    properties: cloneExcelValue(source.properties),
    pageSetup: cloneExcelValue(source.pageSetup),
    views: cloneExcelValue(source.views),
    headerFooter: cloneExcelValue(source.headerFooter),
  });
  target.state = source.state;
  target.autoFilter = cloneExcelValue(source.autoFilter);

  for (let columnNumber = 1; columnNumber <= source.columnCount; columnNumber += 1) {
    const sourceColumn = source.getColumn(columnNumber);
    const targetColumn = target.getColumn(columnNumber);
    if (sourceColumn.width != null) targetColumn.width = sourceColumn.width;
    targetColumn.hidden = sourceColumn.hidden;
    targetColumn.outlineLevel = sourceColumn.outlineLevel;
    targetColumn.style = cloneExcelValue(sourceColumn.style);
  }

  source.eachRow({ includeEmpty: true }, (sourceRow, rowNumber) => {
    const targetRow = target.getRow(rowNumber);
    if (sourceRow.height != null) targetRow.height = sourceRow.height;
    targetRow.hidden = sourceRow.hidden;
    targetRow.outlineLevel = sourceRow.outlineLevel;
    targetRow.style = cloneExcelValue(sourceRow.style);
    sourceRow.eachCell({ includeEmpty: true }, (sourceCell, columnNumber) => {
      const targetCell = targetRow.getCell(columnNumber);
      targetCell.value = cloneExcelValue(sourceCell.value);
      targetCell.style = cloneExcelValue(sourceCell.style);
      targetCell.note = cloneExcelValue(sourceCell.note);
      targetCell.dataValidation = cloneExcelValue(sourceCell.dataValidation);
    });
  });

  source.model.merges.forEach((range) => target.mergeCells(range));
  const sourceModel = source.model as ExcelJS.WorksheetModel & { conditionalFormattings?: ExcelJS.ConditionalFormattingOptions[] };
  sourceModel.conditionalFormattings?.forEach((formatting) => target.addConditionalFormatting(cloneExcelValue(formatting)));
  sourceModel.rowBreaks.forEach((rowBreak) => target.getRow(rowBreak.id).addPageBreak(rowBreak.min, rowBreak.max));
  source.getImages().forEach((image) => {
    const sourceImage = sourceWorkbook.getImage(Number(image.imageId));
    if (!sourceImage) return;
    target.addImage(targetWorkbook.addImage(sourceImage), image.range);
  });
  const backgroundImageId = source.getBackgroundImageId();
  if (backgroundImageId) {
    const sourceImage = sourceWorkbook.getImage(Number(backgroundImageId));
    if (sourceImage) target.addBackgroundImage(targetWorkbook.addImage(sourceImage));
  }
  return target;
}

function copyMatrix(matrix: unknown[][], targetWorkbook: ExcelJS.Workbook, targetName: string) {
  const target = targetWorkbook.addWorksheet(targetName);
  matrix.forEach((values, index) => { target.getRow(index + 1).values = values as ExcelJS.CellValue[]; });
  return target;
}

async function appendOriginalBom(targetWorkbook: ExcelJS.Workbook, source: OriginalBomSource, targetName: string) {
  try {
    const originalWorkbook = new ExcelJS.Workbook();
    await originalWorkbook.xlsx.load(source.data as ExcelJS.Buffer);
    const originalSheet = originalWorkbook.getWorksheet(source.sheetName) ?? originalWorkbook.worksheets[0];
    if (originalSheet) return copyWorksheet(originalWorkbook, originalSheet, targetWorkbook, targetName);
  } catch {
    // Older or unusual workbooks still retain all original cell values through the import matrix.
  }
  return copyMatrix(source.matrix, targetWorkbook, targetName);
}

export function buildBomReport(diffs: BomDiff[], beforeName: string, afterName: string, context: ReportContext = {}) {
  void context;
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "BOMLens Offline";
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.title = "BOM 差異清單";
  workbook.subject = `${beforeName} → ${afterName}`;

  const sheet = workbook.addWorksheet("差異清單", { properties: { defaultRowHeight: 28 } });
  styleTitle(sheet, "A1:M1", "BOM 差異明細");
  sheet.getRow(1).height = 34;
  sheet.getCell("A2").value = "舊版 BOM";
  sheet.getCell("B2").value = beforeName;
  sheet.getCell("D2").value = "新版 BOM";
  sheet.getCell("E2").value = afterName;
  ["A2", "D2"].forEach((address) => {
    sheet.getCell(address).font = { name: "Microsoft JhengHei", bold: true, color: { argb: colors.gray } };
  });

  const header = sheet.getRow(4);
  header.values = [
    "主要異動", "影響標籤", "料號新增／刪除", "插件位置差異",
    "舊版主件料號", "舊版製造廠商料號", "舊版製造廠商",
    "新版主件料號", "新版製造廠商料號", "新版製造廠商",
    "舊版數量", "新版數量", "數量變化",
  ];
  styleHeader(header);

  diffs.forEach((diff, index) => {
    const beforeParts = diff.before?.alternatives ?? [];
    const afterParts = diff.after?.alternatives ?? [];
    const beforeQty = diff.before?.qty ?? 0;
    const afterQty = diff.after?.qty ?? 0;
    const quantityTrend = afterQty > beforeQty ? "Increase" : afterQty < beforeQty ? "Decrease" : "Same";
    const row = sheet.getRow(5 + index);
    row.values = [
      primaryLabel(diff), diff.fields.join("\n"), partChanges(diff), positionChanges(diff),
      listPartNumbers(beforeParts), listManufacturerParts(beforeParts), listManufacturers(beforeParts),
      listPartNumbers(afterParts), listManufacturerParts(afterParts), listManufacturers(afterParts),
      beforeQty, afterQty, quantityTrend,
    ];
    row.height = Math.max(beforeParts.length, afterParts.length) > 1 ? 54 : 36;
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.font = { name: "Microsoft JhengHei", size: 10, color: { argb: "344054" } };
      cell.alignment = { vertical: "middle", wrapText: true };
      cell.border = border();
      if (index % 2 === 1) cell.fill = fill("FAFBFC");
    });
    const typeCell = row.getCell(1);
    const label = primaryLabel(diff);
    const tone = label === "新增"
      ? { font: colors.green, fill: colors.paleGreen }
      : label === "刪除"
        ? { font: colors.red, fill: colors.paleRed }
        : { font: colors.blue, fill: colors.paleBlue };
    typeCell.font = { name: "Microsoft JhengHei", size: 10, bold: true, color: { argb: tone.font } };
    typeCell.fill = fill(tone.fill);
    row.getCell(11).numFmt = "#,##0";
    row.getCell(12).numFmt = "#,##0";
    const trendCell = row.getCell(13);
    trendCell.font = { name: "Microsoft JhengHei", size: 10, bold: true, color: { argb: quantityTrend === "Increase" ? colors.green : quantityTrend === "Decrease" ? colors.red : colors.gray } };
    trendCell.fill = fill(quantityTrend === "Increase" ? colors.paleGreen : quantityTrend === "Decrease" ? colors.paleRed : colors.paleGray);
  });

  if (!diffs.length) {
    sheet.mergeCells("A5:M6");
    sheet.getCell("A5").value = "目前篩選結果沒有差異。";
    sheet.getCell("A5").font = { name: "Microsoft JhengHei", size: 11, color: { argb: colors.gray } };
    sheet.getCell("A5").alignment = { horizontal: "center", vertical: "middle" };
    sheet.getCell("A5").fill = fill(colors.paleGray);
    sheet.getCell("A5").border = border();
  }

  sheet.columns = [13, 22, 24, 24, 22, 26, 20, 22, 26, 20, 12, 12, 14].map((width) => ({ width }));
  sheet.autoFilter = { from: "A4", to: "M4" };
  sheet.views = [{ state: "frozen", ySplit: 4, xSplit: 4 }];
  sheet.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.25, right: 0.25, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } };

  return workbook;
}

export async function buildBomReportWithOriginals(diffs: BomDiff[], beforeName: string, afterName: string, context: ReportContext = {}) {
  const workbook = buildBomReport(diffs, beforeName, afterName, context);
  if (context.originalBefore) {
    await appendOriginalBom(workbook, context.originalBefore, "舊版原始 BOM");
    workbook.getWorksheet("差異清單")!.getCell("B2").value = { text: beforeName, hyperlink: "#'舊版原始 BOM'!A1" };
  }
  if (context.originalAfter) {
    await appendOriginalBom(workbook, context.originalAfter, "新版原始 BOM");
    workbook.getWorksheet("差異清單")!.getCell("E2").value = { text: afterName, hyperlink: "#'新版原始 BOM'!A1" };
  }
  return workbook;
}

export async function exportBomReport(diffs: BomDiff[], beforeName: string, afterName: string, context: ReportContext = {}) {
  const workbook = await buildBomReportWithOriginals(diffs, beforeName, afterName, context);
  const buffer = await workbook.xlsx.writeBuffer();
  const stamp = new Date().toISOString().slice(0, 10).replaceAll("-", "");
  saveBuffer(buffer, `BOM_diff_report_${stamp}.xlsx`);
}
