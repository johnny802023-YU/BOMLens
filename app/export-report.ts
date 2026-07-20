import ExcelJS from "exceljs";
import { bomStructureLabel, companyColumnLabels, type BomAlternative, type BomDiff, type CompanyColumnKey, type DiffPrimaryType, type ImportAudit } from "./bom-logic.ts";

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

const primaryLabels: Record<DiffPrimaryType, string> = {
  componentAdded: "新增元件",
  componentRemoved: "移除元件",
  substituteAdded: "新增替料",
  substituteRemoved: "刪除替料",
  partReplaced: "同位置換料",
  positionChanged: "位置變更",
  same: "相同",
};

function partLabel(part: BomAlternative) {
  return part.part || part.manufacturerPart || "未提供料號";
}

function listParts(parts: BomAlternative[]) {
  return parts.map(partLabel).join("\n");
}

function listMpn(parts: BomAlternative[]) {
  return parts.map((part) => part.manufacturerPart).filter(Boolean).join("\n");
}

function listManufacturerNames(parts: BomAlternative[]) {
  return parts.map((part) => part.manufacturerName).filter(Boolean).join("\n");
}

function diffStructureLabel(diff: BomDiff) {
  const before = bomStructureLabel(diff.before);
  const after = bomStructureLabel(diff.after);
  if (before && after && before !== after) return `${before} → ${after}`;
  return after || before;
}

function primaryWithStructure(diff: BomDiff) {
  const structure = diffStructureLabel(diff);
  return `${primaryLabels[diff.primaryType]}${structure ? `\n所屬架構：${structure}` : ""}`;
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

function primaryColors(type: DiffPrimaryType) {
  if (type === "componentAdded" || type === "substituteAdded") return { font: colors.green, fill: colors.paleGreen };
  if (type === "componentRemoved" || type === "substituteRemoved") return { font: colors.red, fill: colors.paleRed };
  return { font: colors.amber, fill: colors.paleAmber };
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
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "BOMLens Offline";
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.title = "BOM 版本差異報告";
  workbook.subject = `${beforeName} → ${afterName}`;

  const summary = workbook.addWorksheet("差異摘要", { properties: { defaultRowHeight: 20 } });
  summary.columns = [{ width: 24 }, { width: 18 }, { width: 24 }, { width: 18 }, { width: 24 }, { width: 18 }];
  styleTitle(summary, "A1:F2", "BOM 版本差異報告");
  summary.getRow(1).height = 30;
  summary.getRow(2).height = 20;
  summary.getCell("A4").value = "前版 BOM";
  summary.getCell("B4").value = beforeName;
  summary.getCell("D4").value = "後版 BOM";
  summary.getCell("E4").value = afterName;
  summary.getCell("A5").value = "產生時間";
  summary.getCell("B5").value = new Date();
  summary.getCell("B5").numFmt = "yyyy-mm-dd hh:mm";
  summary.getCell("D5").value = "工作表";
  summary.getCell("E5").value = `${context.before?.sheetName ?? "—"} → ${context.after?.sheetName ?? "—"}`;
  ["A4", "D4", "A5", "D5"].forEach((address) => {
    const cell = summary.getCell(address);
    cell.font = { name: "Microsoft JhengHei", bold: true, color: { argb: colors.gray } };
  });

  const counts = new Map<DiffPrimaryType, number>();
  diffs.forEach((diff) => counts.set(diff.primaryType, (counts.get(diff.primaryType) ?? 0) + 1));
  const newPartCount = new Set(diffs.flatMap((diff) => diff.newParts.map((part) => part.part))).size;
  const deletedPartCount = new Set(diffs.flatMap((diff) => diff.deletedParts.map((part) => part.part))).size;
  const cards = [
    ["差異群組", diffs.length, colors.paleBlue, colors.blue],
    ["新版完全新料", newPartCount, colors.paleGreen, colors.green],
    ["新版完全移除", deletedPartCount, colors.paleRed, colors.red],
  ] as const;
  cards.forEach(([label, value, background, foreground], index) => {
    const col = 1 + index * 2;
    const labelCell = summary.getCell(8, col);
    const valueCell = summary.getCell(9, col);
    summary.mergeCells(8, col, 8, col + 1);
    summary.mergeCells(9, col, 10, col + 1);
    labelCell.value = label;
    valueCell.value = value;
    [labelCell, valueCell].forEach((cell) => {
      cell.fill = fill(background);
      cell.border = border();
      cell.alignment = { horizontal: "center", vertical: "middle" };
    });
    labelCell.font = { name: "Microsoft JhengHei", size: 10, bold: true, color: { argb: foreground } };
    valueCell.font = { name: "Microsoft JhengHei", size: 22, bold: true, color: { argb: foreground } };
  });
  summary.getRow(9).height = 28;
  summary.getRow(10).height = 22;

  summary.mergeCells("A13:F13");
  summary.getCell("A13").value = "主要異動分類";
  summary.getCell("A13").font = { name: "Microsoft JhengHei", bold: true, color: { argb: colors.white } };
  summary.getCell("A13").fill = fill(colors.blue);
  const typeRows = ["componentAdded", "substituteAdded", "componentRemoved", "substituteRemoved", "positionChanged", "partReplaced"] as DiffPrimaryType[];
  typeRows.forEach((type, index) => {
    const row = summary.getRow(14 + index);
    row.values = [primaryLabels[type], counts.get(type) ?? 0, "", "", "", ""];
    row.getCell(1).font = { name: "Microsoft JhengHei", bold: true, color: { argb: primaryColors(type).font } };
    row.getCell(1).fill = fill(primaryColors(type).fill);
    row.getCell(2).numFmt = "#,##0";
    row.eachCell({ includeEmpty: true }, (cell) => { cell.border = border(); });
  });
  summary.mergeCells("A22:F24");
  summary.getCell("A22").value = "判斷說明：料號取最右 12 碼；依 69 → VB-D／60／VB-T／08 PCB 的順序架構分組；項次只切分同一架構內的主替料且不跨版比較；數量優先採用插件位置數，插件位置空白時才採用數量欄；製造商名稱與製造商料號只供顯示，客戶料號暫不使用。";
  summary.getCell("A22").alignment = { vertical: "top", wrapText: true };
  summary.getCell("A22").font = { name: "Microsoft JhengHei", size: 10, color: { argb: colors.gray } };
  summary.getCell("A22").fill = fill(colors.paleGray);
  summary.getCell("A22").border = border();
  summary.views = [{ state: "frozen", ySplit: 3 }];
  summary.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 1 };

  const details = workbook.addWorksheet("差異明細", { properties: { defaultRowHeight: 28 } });
  const detailHeaders = [
    "主要異動", "影響標籤", "新版完全新料", "新版完全移除", "前版料號", "後版料號",
    "前版製造商料號", "後版製造商料號", "前版製造商名稱", "後版製造商名稱", "新增料號", "刪除料號",
    "新增插件位置", "移除插件位置", "同位置換料", "前版數量", "後版數量", "前版項次（追溯）", "後版項次（追溯）",
    "配對可信度", "配對依據", "待人工確認", "前版原始列", "後版原始列",
  ];
  styleTitle(details, `A1:X1`, "BOM 差異明細");
  details.mergeCells("A2:X2");
  details.getCell("A2").value = `${beforeName}  →  ${afterName}`;
  details.getCell("A2").font = { name: "Microsoft JhengHei", size: 10, color: { argb: colors.gray } };
  details.getCell("A2").alignment = { vertical: "middle" };
  const headerRow = details.getRow(4);
  headerRow.values = detailHeaders;
  styleHeader(headerRow);

  diffs.forEach((diff, index) => {
    const row = details.getRow(5 + index);
    row.values = [
      primaryWithStructure(diff), diff.fields.join("、"), listParts(diff.newParts), listParts(diff.deletedParts),
      diff.before ? listParts(diff.before.alternatives) : "", diff.after ? listParts(diff.after.alternatives) : "",
      diff.before ? listMpn(diff.before.alternatives) : "", diff.after ? listMpn(diff.after.alternatives) : "",
      diff.before ? listManufacturerNames(diff.before.alternatives) : "", diff.after ? listManufacturerNames(diff.after.alternatives) : "",
      listParts(diff.addedParts), listParts(diff.removedParts), diff.addedPositions.join(", "), diff.removedPositions.join(", "),
      diff.replacementPositions.join(", "), diff.before?.qty ?? 0, diff.after?.qty ?? 0, diff.before?.ref ?? "", diff.after?.ref ?? "",
      diff.matchConfidence === "low" ? "低可信" : "", diff.matchConfidence === "low" ? diff.matchReason : "", diff.needsReview ? "是" : "否",
      diff.before?.sourceRows?.join(", ") ?? "", diff.after?.sourceRows?.join(", ") ?? "",
    ];
    row.height = diffStructureLabel(diff) ? 54 : 40;
    row.eachCell({ includeEmpty: true }, (cell, column) => {
      cell.font = { name: "Microsoft JhengHei", size: 9, color: { argb: "344054" } };
      cell.alignment = { vertical: "middle", wrapText: true };
      cell.border = border();
      if (index % 2 === 1) cell.fill = fill("FAFBFC");
      if (column === 16 || column === 17) cell.numFmt = "#,##0";
    });
    const typeColor = primaryColors(diff.primaryType);
    row.getCell(1).font = { name: "Microsoft JhengHei", size: 9, bold: true, color: { argb: typeColor.font } };
    row.getCell(1).fill = fill(typeColor.fill);
    if (diff.newParts.length) row.getCell(3).fill = fill(colors.paleGreen);
    if (diff.deletedParts.length) row.getCell(4).fill = fill(colors.paleRed);
  });
  details.columns = [13, 24, 20, 20, 20, 20, 24, 24, 20, 20, 18, 18, 18, 18, 16, 11, 11, 14, 14, 12, 24, 14, 18, 18].map((width) => ({ width }));
  details.autoFilter = { from: "A4", to: "X4" };
  details.views = [{ state: "frozen", ySplit: 4, xSplit: 2 }];
  details.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.25, right: 0.25, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } };

  const lifecycle = workbook.addWorksheet("料號生命週期", { properties: { defaultRowHeight: 25 } });
  styleTitle(lifecycle, "A1:H1", "料號生命週期清單");
  lifecycle.mergeCells("A2:H2");
  lifecycle.getCell("A2").value = "集中列出新版完全新料與新版完全移除料號；製造商料號與製造商名稱只供識別，不參與差異判斷。";
  lifecycle.getCell("A2").font = { name: "Microsoft JhengHei", size: 10, color: { argb: colors.gray } };
  const lifecycleHeader = lifecycle.getRow(4);
  lifecycleHeader.values = ["生命週期", "料號", "製造商料號", "製造商名稱", "插件位置", "數量", "主要異動", "項次（追溯）"];
  styleHeader(lifecycleHeader);
  const lifecycleRows: Array<[string, BomAlternative, string, number, string, string]> = [];
  diffs.forEach((diff) => {
    diff.newParts.forEach((part) => lifecycleRows.push(["新版完全新料", part, diff.after?.positions.join(", ") ?? "", diff.after?.qty ?? 0, primaryWithStructure(diff), diff.after?.ref ?? ""]));
    diff.deletedParts.forEach((part) => lifecycleRows.push(["新版完全移除", part, diff.before?.positions.join(", ") ?? "", diff.before?.qty ?? 0, primaryWithStructure(diff), diff.before?.ref ?? ""]));
  });
  lifecycleRows.forEach(([status, part, positions, qty, primary, ref], index) => {
    const row = lifecycle.getRow(5 + index);
    row.values = [status, partLabel(part), part.manufacturerPart, part.manufacturerName ?? "", positions, qty, primary, ref];
    row.height = primary.includes("所屬架構：") ? 42 : 25;
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.font = { name: "Microsoft JhengHei", size: 9, color: { argb: "344054" } };
      cell.alignment = { vertical: "middle", wrapText: true };
      cell.border = border();
    });
    const isNew = status === "新版完全新料";
    row.getCell(1).font = { name: "Microsoft JhengHei", size: 9, bold: true, color: { argb: isNew ? colors.green : colors.red } };
    row.getCell(1).fill = fill(isNew ? colors.paleGreen : colors.paleRed);
    row.getCell(6).numFmt = "#,##0";
  });
  lifecycle.columns = [18, 22, 26, 20, 28, 10, 16, 15].map((width) => ({ width }));
  lifecycle.autoFilter = { from: "A4", to: "H4" };
  lifecycle.views = [{ state: "frozen", ySplit: 4 }];
  lifecycle.pageSetup = { orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0 };

  const review = workbook.addWorksheet("待人工確認", { properties: { defaultRowHeight: 25 } });
  styleTitle(review, "A1:H1", "待人工確認清單");
  review.mergeCells("A2:H2");
  review.getCell("A2").value = "列出配對可信度偏低或存在多個合理配對候選的群組，請回到原始 BOM 人工確認。";
  review.getCell("A2").font = { name: "Microsoft JhengHei", size: 10, color: { argb: colors.gray } };
  const reviewHeader = review.getRow(4);
  reviewHeader.values = ["主要異動", "前版料號", "新版料號", "插件位置", "可信度", "配對依據", "前版原始列", "新版原始列"];
  styleHeader(reviewHeader);
  diffs.filter((diff) => diff.needsReview).forEach((diff, index) => {
    const row = review.getRow(5 + index);
    row.values = [primaryWithStructure(diff), diff.before ? listParts(diff.before.alternatives) : "", diff.after ? listParts(diff.after.alternatives) : "", diff.after?.positions.join(", ") || diff.before?.positions.join(", ") || "", diff.matchConfidence, diff.matchReason, diff.before?.sourceRows?.join(", ") ?? "", diff.after?.sourceRows?.join(", ") ?? ""];
    row.height = diffStructureLabel(diff) ? 48 : 25;
    row.eachCell({ includeEmpty: true }, (cell) => { cell.font = { name: "Microsoft JhengHei", size: 9, color: { argb: "344054" } }; cell.alignment = { vertical: "middle", wrapText: true }; cell.border = border(); });
  });
  review.columns = [15, 24, 24, 24, 12, 28, 18, 18].map((width) => ({ width }));
  review.autoFilter = { from: "A4", to: "H4" };
  review.views = [{ state: "frozen", ySplit: 4 }];

  const audit = workbook.addWorksheet("匯入稽核", { properties: { defaultRowHeight: 23 } });
  styleTitle(audit, "A1:F1", "匯入稽核紀錄");
  const auditHeader = audit.getRow(3);
  auditHeader.values = ["版本", "檔案", "工作表", "項目", "內容", "原始列"];
  styleHeader(auditHeader);
  const auditRows: Array<[string, string, string, string, string, string]> = [];
  ([{ label: "前版", source: context.before }, { label: "後版", source: context.after }] as const).forEach(({ label, source }) => {
    if (!source) return;
    auditRows.push([label, source.fileName, source.sheetName, "標題列", String(source.audit.headerRow), ""]);
    (Object.entries(source.audit.mappingLabels) as Array<[CompanyColumnKey, string]>).forEach(([key, value]) => auditRows.push([label, source.fileName, source.sheetName, `欄位對應：${companyColumnLabels[key]}`, value, ""]));
    source.audit.issues.forEach((issue) => auditRows.push([label, source.fileName, source.sheetName, issue.severity === "error" ? "錯誤" : "警告", issue.message, issue.rows?.join(", ") ?? ""]));
    if (!source.audit.issues.length) auditRows.push([label, source.fileName, source.sheetName, "資料檢查", "通過，沒有警告", ""]);
  });
  auditRows.forEach((values, index) => {
    const row = audit.getRow(4 + index); row.values = values;
    row.eachCell({ includeEmpty: true }, (cell) => { cell.font = { name: "Microsoft JhengHei", size: 9, color: { argb: "344054" } }; cell.alignment = { vertical: "middle", wrapText: true }; cell.border = border(); });
  });
  audit.columns = [12, 28, 20, 24, 70, 18].map((width) => ({ width }));
  audit.autoFilter = { from: "A3", to: "F3" };
  audit.views = [{ state: "frozen", ySplit: 3 }];

  return workbook;
}

export async function buildBomReportWithOriginals(diffs: BomDiff[], beforeName: string, afterName: string, context: ReportContext = {}) {
  const workbook = buildBomReport(diffs, beforeName, afterName, context);
  if (context.originalBefore) {
    await appendOriginalBom(workbook, context.originalBefore, "前版原始 BOM");
    workbook.getWorksheet("差異摘要")!.getCell("B4").value = { text: beforeName, hyperlink: "#'前版原始 BOM'!A1" };
  }
  if (context.originalAfter) {
    await appendOriginalBom(workbook, context.originalAfter, "後版原始 BOM");
    workbook.getWorksheet("差異摘要")!.getCell("E4").value = { text: afterName, hyperlink: "#'後版原始 BOM'!A1" };
  }
  return workbook;
}

export async function exportBomReport(diffs: BomDiff[], beforeName: string, afterName: string, context: ReportContext = {}) {
  const workbook = await buildBomReportWithOriginals(diffs, beforeName, afterName, context);
  const buffer = await workbook.xlsx.writeBuffer();
  const stamp = new Date().toISOString().slice(0, 10).replaceAll("-", "");
  saveBuffer(buffer, `BOM_diff_report_${stamp}.xlsx`);
}
