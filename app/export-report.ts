import ExcelJS from "exceljs";
import { bomDiffDisplayFields, type BomAlternative, type BomDiff, type ImportAudit } from "./bom-logic.ts";

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
  oldVersion: "58677C",
  paleOldVersion: "F2F4F7",
  newVersion: "2F6BCE",
  paleNewVersion: "EDF4FF",
  white: "FFFFFF",
  border: "D9E0E8",
};

const exportCategoryOrder = ["新版完全新料", "新增替代", "新版完全移除", "刪除替代", "數量差異", "製程別放置異常"] as const;
type ExportCategory = (typeof exportCategoryOrder)[number];
const exportSectionOrder = [...exportCategoryOrder, "僅插件位置差異"] as const;
type ExportSection = (typeof exportSectionOrder)[number];
const detailHeaders = [
  "主要異動", "差異項目", "料號異動", "插件位置差異",
  "舊版主件料號", "舊版製造廠商料號", "舊版製造廠商",
  "新版主件料號", "新版製造廠商料號", "新版製造廠商",
  "舊版數量", "新版數量", "數量變化",
] as const;
const dataHeaders = ["分類", ...detailHeaders] as const;
const summaryColumnGroups = [
  { label: "主要異動", start: 1, end: 2 },
  { label: "差異項目", start: 3, end: 4 },
  { label: "料號異動", start: 5, end: 6 },
  { label: "插件位置差異", start: 7, end: 8 },
  { label: "舊版料號資訊", start: 9, end: 11, tone: "old" },
  { label: "→", start: 12, end: 12 },
  { label: "新版料號資訊", start: 13, end: 15, tone: "new" },
  { label: "數量", start: 16, end: 17 },
] as const;
const summarySubHeaders = [
  { label: "新增／刪除／變更", start: 1, end: 2 },
  { label: "差異標籤", start: 3, end: 4 },
  { label: "料號新增／刪除", start: 5, end: 6 },
  { label: "插件位置", start: 7, end: 8 },
  { label: "舊版料號", start: 9, end: 9, tone: "old" },
  { label: "舊版製造商料號", start: 10, end: 10, tone: "old" },
  { label: "舊版製造商", start: 11, end: 11, tone: "old" },
  { label: "→", start: 12, end: 12 },
  { label: "新版料號", start: 13, end: 13, tone: "new" },
  { label: "新版製造商料號", start: 14, end: 14, tone: "new" },
  { label: "新版製造商", start: 15, end: 15, tone: "new" },
  { label: "舊版 → 新版", start: 16, end: 17 },
] as const;

function partLabel(part: BomAlternative) {
  return part.part || part.manufacturerPart || "未提供料號";
}

function partIdentity(part: BomAlternative) {
  return part.part.trim().toUpperCase();
}

function includesPart(parts: BomAlternative[], candidate?: BomAlternative) {
  if (!candidate) return false;
  const identity = partIdentity(candidate);
  return Boolean(identity) && parts.some((part) => partIdentity(part) === identity);
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

export function exportCategories(diff: BomDiff): ExportCategory[] {
  const categories: ExportCategory[] = [];
  if ((diff.primaryType === "componentAdded" || diff.primaryType === "partReplaced") && diff.newParts.length > 0) {
    categories.push("新版完全新料");
  }
  if ((diff.primaryType === "componentRemoved" || diff.primaryType === "partReplaced") && diff.deletedParts.length > 0) {
    categories.push("新版完全移除");
  }
  if (diff.primaryType === "substituteAdded") categories.push("新增替代");
  if (diff.primaryType === "substituteRemoved") categories.push("刪除替代");
  if (diff.processChange) categories.push("製程別放置異常");
  if (diff.before && diff.after && diff.before.qty !== diff.after.qty) categories.push("數量差異");
  return categories;
}

function belongsToSection(diff: BomDiff, section: ExportSection) {
  const categories = exportCategories(diff);
  return section === "僅插件位置差異" ? categories.length === 0 : categories.includes(section);
}

function exportCategoryLabel(diff: BomDiff) {
  const categories = exportCategories(diff);
  return categories.length ? categories.join("、") : "僅插件位置差異";
}

function sortForData(diffs: BomDiff[]) {
  return diffs
    .map((diff, index) => ({ diff, index, categories: exportCategories(diff) }))
    .sort((left, right) => {
      const leftRank = left.categories.length ? exportCategoryOrder.indexOf(left.categories[0]) : exportCategoryOrder.length;
      const rightRank = right.categories.length ? exportCategoryOrder.indexOf(right.categories[0]) : exportCategoryOrder.length;
      return leftRank - rightRank || left.index - right.index;
    })
    .map(({ diff }) => diff);
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

function styleChangedPartCells(row: ExcelJS.Row, columns: number[], tone: "removed" | "added") {
  const foreground = tone === "removed" ? colors.red : colors.green;
  const background = tone === "removed" ? colors.paleRed : colors.paleGreen;
  columns.forEach((column, index) => {
    const cell = row.getCell(column);
    cell.fill = fill(background);
    cell.font = {
      name: "Microsoft JhengHei",
      size: 10,
      bold: index === 0,
      color: { argb: foreground },
    };
    if (index === 0) {
      cell.border = {
        ...border(),
        left: { style: "medium", color: { argb: foreground } },
      };
    }
  });
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

function sectionTone(section: ExportSection) {
  if (section === "新版完全新料" || section === "新增替代") {
    return { foreground: colors.green, background: colors.paleGreen };
  }
  if (section === "新版完全移除" || section === "刪除替代") {
    return { foreground: colors.red, background: colors.paleRed };
  }
  if (section === "數量差異") {
    return { foreground: colors.amber, background: colors.paleAmber };
  }
  if (section === "製程別放置異常") {
    return { foreground: colors.amber, background: colors.paleAmber };
  }
  return { foreground: colors.blue, background: colors.paleBlue };
}

function styleSummarySectionTitle(sheet: ExcelJS.Worksheet, rowNumber: number, section: ExportSection, count: number) {
  sheet.mergeCells(rowNumber, 1, rowNumber, 17);
  const row = sheet.getRow(rowNumber);
  row.height = 24;
  const cell = row.getCell(1);
  const tone = sectionTone(section);
  cell.value = `${section}（${count}）`;
  cell.font = { name: "Microsoft JhengHei", size: 11, bold: true, color: { argb: tone.foreground } };
  cell.fill = fill(tone.background);
  cell.alignment = { vertical: "middle", horizontal: "left" };
  cell.border = border();
}

function detailRowValues(diff: BomDiff): ExcelJS.CellValue[] {
  const beforeParts = diff.before?.alternatives ?? [];
  const afterParts = diff.after?.alternatives ?? [];
  const beforeQty = diff.before?.qty ?? 0;
  const afterQty = diff.after?.qty ?? 0;
  const quantityTrend = afterQty > beforeQty ? "Increase" : afterQty < beforeQty ? "Decrease" : "Same";
  return [
    primaryLabel(diff), bomDiffDisplayFields(diff).join("\n"), partChanges(diff), positionChanges(diff),
    listPartNumbers(beforeParts), listManufacturerParts(beforeParts), listManufacturers(beforeParts),
    listPartNumbers(afterParts), listManufacturerParts(afterParts), listManufacturers(afterParts),
    beforeQty, afterQty, quantityTrend,
  ];
}

function styleDataRow(row: ExcelJS.Row, diff: BomDiff, striped: boolean) {
  const beforeParts = diff.before?.alternatives ?? [];
  const afterParts = diff.after?.alternatives ?? [];
  const beforeQty = diff.before?.qty ?? 0;
  const afterQty = diff.after?.qty ?? 0;
  const quantityTrend = afterQty > beforeQty ? "Increase" : afterQty < beforeQty ? "Decrease" : "Same";
  row.height = Math.max(beforeParts.length, afterParts.length) > 1 ? 54 : 36;
  row.eachCell({ includeEmpty: true }, (cell) => {
    cell.font = { name: "Microsoft JhengHei", size: 10, color: { argb: "344054" } };
    cell.alignment = { vertical: "middle", wrapText: true };
    cell.border = border();
    if (striped) cell.fill = fill("FAFBFC");
  });
  const categoryCell = row.getCell(1);
  const categories = exportCategories(diff);
  const section = (categories[0] ?? "僅插件位置差異") as ExportSection;
  const categoryTone = sectionTone(section);
  categoryCell.font = { name: "Microsoft JhengHei", size: 10, bold: true, color: { argb: categoryTone.foreground } };
  categoryCell.fill = fill(categoryTone.background);
  const typeCell = row.getCell(2);
  const label = primaryLabel(diff);
  const tone = label === "新增"
    ? { font: colors.green, fill: colors.paleGreen }
    : label === "刪除"
      ? { font: colors.red, fill: colors.paleRed }
      : { font: colors.blue, fill: colors.paleBlue };
  typeCell.font = { name: "Microsoft JhengHei", size: 10, bold: true, color: { argb: tone.font } };
  typeCell.fill = fill(tone.fill);
  [6, 7, 8, 12].forEach((column) => { row.getCell(column).fill = fill(colors.paleOldVersion); });
  [9, 10, 11, 13].forEach((column) => { row.getCell(column).fill = fill(colors.paleNewVersion); });
  if (beforeParts.length > 0 && beforeParts.every((part) => includesPart(diff.removedParts, part))) {
    styleChangedPartCells(row, [6, 7, 8], "removed");
  }
  if (afterParts.length > 0 && afterParts.every((part) => includesPart(diff.addedParts, part))) {
    styleChangedPartCells(row, [9, 10, 11], "added");
  }
  row.getCell(12).numFmt = "#,##0";
  row.getCell(13).numFmt = "#,##0";
  const trendCell = row.getCell(14);
  trendCell.font = {
    name: "Microsoft JhengHei",
    size: 10,
    bold: true,
    color: { argb: quantityTrend === "Increase" ? colors.green : quantityTrend === "Decrease" ? colors.red : colors.gray },
  };
  trendCell.fill = fill(quantityTrend === "Increase" ? colors.paleGreen : quantityTrend === "Decrease" ? colors.paleRed : colors.paleGray);
}

function styleSummaryHeaderRow(
  sheet: ExcelJS.Worksheet,
  rowNumber: number,
  headers: ReadonlyArray<{ label: string; start: number; end: number; tone?: "old" | "new" }>,
  color: string,
) {
  headers.forEach((group) => {
    if (group.end > group.start) sheet.mergeCells(rowNumber, group.start, rowNumber, group.end);
    const cell = sheet.getRow(rowNumber).getCell(group.start);
    cell.value = group.label;
    cell.font = { name: "Microsoft JhengHei", size: 10, bold: true, color: { argb: colors.white } };
    const groupColor = group.tone === "old" ? colors.oldVersion : group.tone === "new" ? colors.newVersion : color;
    cell.fill = fill(groupColor);
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = border();
  });
  sheet.getRow(rowNumber).height = 28;
}

function writeSummaryDiffRows(
  sheet: ExcelJS.Worksheet,
  startRow: number,
  diff: BomDiff,
  section: ExportSection,
  striped: boolean,
) {
  const beforeParts = diff.before?.alternatives ?? [];
  const afterParts = diff.after?.alternatives ?? [];
  const partRowCount = Math.max(beforeParts.length, afterParts.length, 1);
  const endRow = startRow + partRowCount - 1;
  for (let rowNumber = startRow; rowNumber <= endRow; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    row.height = 34;
    for (let column = 1; column <= 17; column += 1) {
      const cell = row.getCell(column);
      cell.font = { name: "Microsoft JhengHei", size: 10, color: { argb: "344054" } };
      cell.alignment = { vertical: "middle", horizontal: column === 12 || column >= 16 ? "center" : "left", wrapText: true };
      cell.border = border();
      if (striped) cell.fill = fill("FAFBFC");
    }
  }

  [
    { start: 1, end: 2 },
    { start: 3, end: 4 },
    { start: 5, end: 6 },
    { start: 7, end: 8 },
    { start: 12, end: 12 },
    { start: 16, end: 17 },
  ].forEach((group) => sheet.mergeCells(startRow, group.start, endRow, group.end));

  const row = sheet.getRow(startRow);
  row.getCell(1).value = primaryLabel(diff);
  row.getCell(3).value = bomDiffDisplayFields(diff).join("\n") || primaryLabel(diff);
  row.getCell(5).value = partChanges(diff) || "料號無增減";
  row.getCell(7).value = positionChanges(diff) || "—";
  row.getCell(12).value = "→";
  const beforeQty = diff.before?.qty ?? 0;
  const afterQty = diff.after?.qty ?? 0;
  row.getCell(16).value = `${beforeQty} → ${afterQty}`;

  for (let index = 0; index < partRowCount; index += 1) {
    const partRow = sheet.getRow(startRow + index);
    const beforePart = beforeParts[index];
    const afterPart = afterParts[index];
    partRow.getCell(9).value = beforePart ? partLabel(beforePart) : index === 0 ? "—" : "";
    partRow.getCell(10).value = beforePart?.manufacturerPart ?? "";
    partRow.getCell(11).value = beforePart?.manufacturerName ?? "";
    partRow.getCell(13).value = afterPart ? partLabel(afterPart) : index === 0 ? "—" : "";
    partRow.getCell(14).value = afterPart?.manufacturerPart ?? "";
    partRow.getCell(15).value = afterPart?.manufacturerName ?? "";
    [9, 10, 11].forEach((column) => { partRow.getCell(column).fill = fill(colors.paleOldVersion); });
    [13, 14, 15].forEach((column) => { partRow.getCell(column).fill = fill(colors.paleNewVersion); });
    partRow.getCell(9).font = { name: "Microsoft JhengHei", size: 10, bold: index === 0, color: { argb: "344054" } };
    partRow.getCell(13).font = { name: "Microsoft JhengHei", size: 10, bold: index === 0, color: { argb: "344054" } };
    if (includesPart(diff.removedParts, beforePart)) styleChangedPartCells(partRow, [9, 10, 11], "removed");
    if (includesPart(diff.addedParts, afterPart)) styleChangedPartCells(partRow, [13, 14, 15], "added");
  }

  const tone = sectionTone(section);
  const typeCell = row.getCell(1);
  typeCell.font = { name: "Microsoft JhengHei", size: 10, bold: true, color: { argb: tone.foreground } };
  typeCell.fill = fill(tone.background);
  const arrowCell = row.getCell(12);
  arrowCell.font = { name: "Microsoft JhengHei", size: 16, bold: true, color: { argb: tone.foreground } };
  const quantityCell = row.getCell(16);
  quantityCell.font = {
    name: "Microsoft JhengHei",
    size: 11,
    bold: true,
    color: { argb: afterQty > beforeQty ? colors.green : afterQty < beforeQty ? colors.red : colors.gray },
  };
  quantityCell.fill = fill(afterQty > beforeQty ? colors.paleGreen : afterQty < beforeQty ? colors.paleRed : colors.paleGray);
  return endRow + 1;
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

  const summarySheet = workbook.addWorksheet("差異摘要", { properties: { defaultRowHeight: 26 } });
  styleTitle(summarySheet, "A1:Q1", "BOM 差異比較報告");
  summarySheet.getRow(1).height = 36;
  summarySheet.mergeCells("A2:B2");
  summarySheet.mergeCells("C2:H2");
  summarySheet.mergeCells("I2:J2");
  summarySheet.mergeCells("K2:Q2");
  summarySheet.getCell("A2").value = "舊版 BOM";
  summarySheet.getCell("C2").value = beforeName;
  summarySheet.getCell("I2").value = "新版 BOM";
  summarySheet.getCell("K2").value = afterName;
  ["A2", "I2"].forEach((address) => {
    summarySheet.getCell(address).font = { name: "Microsoft JhengHei", bold: true, color: { argb: colors.gray } };
  });
  for (let column = 1; column <= 8; column += 1) summarySheet.getCell(2, column).fill = fill(colors.paleOldVersion);
  for (let column = 9; column <= 17; column += 1) summarySheet.getCell(2, column).fill = fill(colors.paleNewVersion);

  const cardRanges = ["A4:B5", "C4:D5", "E4:F5", "G4:I5", "J4:L5", "M4:O5", "P4:Q5"];
  exportSectionOrder.forEach((section, index) => {
    const sectionDiffs = diffs.filter((diff) => belongsToSection(diff, section));
    const range = cardRanges[index];
    summarySheet.mergeCells(range);
    const cell = summarySheet.getCell(range.split(":")[0]);
    const tone = sectionTone(section);
    cell.value = `${section}\n${sectionDiffs.length}`;
    cell.font = { name: "Microsoft JhengHei", size: 11, bold: true, color: { argb: tone.foreground } };
    cell.fill = fill(tone.background);
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = border();
  });
  summarySheet.mergeCells("A6:Q6");
  summarySheet.getCell("A6").value = "顏色說明｜淡紅：舊版刪除／停用料號　淡綠：新版新增／改用料號　灰／藍：料號未異動";
  summarySheet.getCell("A6").font = { name: "Microsoft JhengHei", size: 9, color: { argb: colors.gray } };
  summarySheet.getCell("A6").alignment = { horizontal: "right", vertical: "middle" };

  styleSummaryHeaderRow(summarySheet, 7, summaryColumnGroups, colors.navy);
  styleSummaryHeaderRow(summarySheet, 8, summarySubHeaders, colors.blue);
  let summaryRow = 9;
  exportSectionOrder.forEach((section) => {
    const sectionDiffs = diffs.filter((diff) => belongsToSection(diff, section));
    styleSummarySectionTitle(summarySheet, summaryRow, section, sectionDiffs.length);
    summaryRow += 1;
    if (sectionDiffs.length) {
      sectionDiffs.forEach((diff, index) => {
        summaryRow = writeSummaryDiffRows(summarySheet, summaryRow, diff, section, index % 2 === 1);
      });
    } else {
      summarySheet.mergeCells(summaryRow, 1, summaryRow, 17);
      const emptyCell = summarySheet.getRow(summaryRow).getCell(1);
      emptyCell.value = "此分類無差異";
      emptyCell.font = { name: "Microsoft JhengHei", size: 10, italic: true, color: { argb: colors.gray } };
      emptyCell.alignment = { horizontal: "center", vertical: "middle" };
      emptyCell.fill = fill(colors.paleGray);
      emptyCell.border = border();
      summarySheet.getRow(summaryRow).height = 24;
      summaryRow += 1;
    }
  });
  summarySheet.mergeCells(summaryRow + 1, 1, summaryRow + 1, 17);
  summarySheet.getCell(summaryRow + 1, 1).value = "閱讀版｜完整獨立欄位與篩選請至「差異資料」工作表";
  summarySheet.getCell(summaryRow + 1, 1).font = { name: "Microsoft JhengHei", size: 9, italic: true, color: { argb: colors.gray } };
  summarySheet.getCell(summaryRow + 1, 1).alignment = { horizontal: "right", vertical: "middle" };
  summarySheet.columns = [10, 10, 13, 13, 15, 15, 12, 12, 15, 15, 15, 6, 15, 15, 15, 9, 9].map((width) => ({ width }));
  summarySheet.views = [{ state: "frozen", ySplit: 8, showGridLines: false }];
  summarySheet.pageSetup = {
    orientation: "landscape",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    margins: { left: 0.25, right: 0.25, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 },
  };

  const dataSheet = workbook.addWorksheet("差異資料", { properties: { defaultRowHeight: 28 } });
  styleTitle(dataSheet, "A1:N1", "BOM 差異資料");
  dataSheet.getRow(1).height = 34;
  dataSheet.getCell("A2").value = "舊版 BOM";
  dataSheet.getCell("B2").value = beforeName;
  dataSheet.getCell("D2").value = "新版 BOM";
  dataSheet.getCell("E2").value = afterName;
  ["A2", "D2"].forEach((address) => {
    dataSheet.getCell(address).font = { name: "Microsoft JhengHei", bold: true, color: { argb: colors.gray } };
  });
  dataSheet.mergeCells("A3:N3");
  dataSheet.getCell("A3").value = "顏色說明｜淡紅：舊版刪除／停用料號　淡綠：新版新增／改用料號　灰／藍：料號未異動";
  dataSheet.getCell("A3").font = { name: "Microsoft JhengHei", size: 9, color: { argb: colors.gray } };
  dataSheet.getCell("A3").alignment = { horizontal: "right", vertical: "middle" };
  const sortedDiffs = sortForData(diffs);
  const dataRows = sortedDiffs.map((diff) => [exportCategoryLabel(diff), ...detailRowValues(diff)]);
  dataSheet.addTable({
    name: "BomDiffData",
    ref: "A4",
    headerRow: true,
    totalsRow: false,
    style: { theme: "TableStyleMedium2", showRowStripes: true },
    columns: dataHeaders.map((name) => ({ name, filterButton: true })),
    rows: dataRows,
  });
  styleHeader(dataSheet.getRow(4));
  [6, 7, 8, 12].forEach((column) => { dataSheet.getRow(4).getCell(column).fill = fill(colors.oldVersion); });
  [9, 10, 11, 13].forEach((column) => { dataSheet.getRow(4).getCell(column).fill = fill(colors.newVersion); });
  sortedDiffs.forEach((diff, index) => styleDataRow(dataSheet.getRow(5 + index), diff, index % 2 === 1));
  dataSheet.columns = [20, 13, 22, 24, 24, 22, 26, 20, 22, 26, 20, 12, 12, 14].map((width) => ({ width }));
  dataSheet.views = [{ state: "frozen", ySplit: 4, xSplit: 5, showGridLines: false }];
  dataSheet.pageSetup = {
    orientation: "landscape",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    margins: { left: 0.25, right: 0.25, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 },
  };

  return workbook;
}

export async function buildBomReportWithOriginals(diffs: BomDiff[], beforeName: string, afterName: string, context: ReportContext = {}) {
  const workbook = buildBomReport(diffs, beforeName, afterName, context);
  if (context.originalBefore) {
    await appendOriginalBom(workbook, context.originalBefore, "舊版原始 BOM");
    workbook.getWorksheet("差異摘要")!.getCell("C2").value = { text: beforeName, hyperlink: "#'舊版原始 BOM'!A1" };
    workbook.getWorksheet("差異資料")!.getCell("B2").value = { text: beforeName, hyperlink: "#'舊版原始 BOM'!A1" };
  }
  if (context.originalAfter) {
    await appendOriginalBom(workbook, context.originalAfter, "新版原始 BOM");
    workbook.getWorksheet("差異摘要")!.getCell("K2").value = { text: afterName, hyperlink: "#'新版原始 BOM'!A1" };
    workbook.getWorksheet("差異資料")!.getCell("E2").value = { text: afterName, hyperlink: "#'新版原始 BOM'!A1" };
  }
  return workbook;
}

export async function exportBomReport(diffs: BomDiff[], beforeName: string, afterName: string, context: ReportContext = {}) {
  const workbook = await buildBomReportWithOriginals(diffs, beforeName, afterName, context);
  const buffer = await workbook.xlsx.writeBuffer();
  const stamp = new Date().toISOString().slice(0, 10).replaceAll("-", "");
  saveBuffer(buffer, `BOM_diff_report_${stamp}.xlsx`);
}
