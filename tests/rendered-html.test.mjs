import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { analyzeCompanyBomMatrix, bomDiffDisplayFields, bomProcessKind, bomQuantityChanged, canonicalPartNumber, compareBom, detectCompanyColumns, findCompanyHeader, parseCompanyBomMatrix, parseQuantity, sortBomDiffsForAll } from "../app/bom-logic.ts";
import { buildPageReferenceIndex, centeredPdfHitScroll, centeredRenderedHitScroll, findReferenceHits, lookupReferenceHits, normalizeReference } from "../app/pdf-search.ts";
import { calculateMva, detectCustomerColumns, detectPlacementColumns, mapCustomerBom, parseCustomerBomMatrix, parsePlacementMatrix } from "../app/supplemental-logic.ts";
import { createSchematicReportPlan } from "../app/schematic-report-logic.ts";
import { parseWorkbookData } from "../app/workbook-import.ts";
import ExcelJS from "exceljs";
import * as XLSX from "xlsx";

async function render() {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);
  return worker.fetch(
    new Request("http://localhost/", { headers: { accept: "text/html" } }),
    { ASSETS: { fetch: async () => new Response("Not found", { status: 404 }) } },
    { waitUntil() {}, passThroughOnException() {} },
  );
}

test("server-renders the BOM comparison workspace", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);
  const html = await response.text();
  assert.match(html, /<html lang="zh-Hant">/i);
  assert.match(html, /<title>BOMLens｜BOM 與線路圖版本比對<\/title>/i);
  assert.match(html, /版本比對/);
  assert.match(html, /BOM 差異/);
  assert.match(html, /料號異動/);
  assert.match(html, /插件位置差異/);
  assert.match(html, /比對規則/);
  assert.match(html, /S／R 欄自動辨識/);
  assert.match(html, /主要異動/);
  assert.match(html, /差異項目/);
  assert.match(html, /只存在新版，需要確認採購與插件位置/);
  assert.match(html, /插件位置、數量、同位置換料或所屬架構異動/);
  assert.match(html, /新版完全新料/);
  assert.match(html, /新版完全移除/);
  assert.match(html, /新增料號/);
  assert.match(html, /新增替料/);
  assert.match(html, /刪除替料/);
  assert.match(html, /刪除料號/);
  assert.match(html, /新增插件位置/);
  assert.match(html, /移除插件位置/);
  assert.match(html, /更換料號/);
  assert.doesNotMatch(html, /前版|後版|新增元件|移除元件/);
  assert.match(html, /新版完全新料/);
  assert.doesNotMatch(html, /需確認下單/);
  assert.match(html, /新版完全移除/);
  assert.doesNotMatch(html, /主替料變更|主體料變更/);
  assert.doesNotMatch(html, /來源 A/);
  assert.match(html, /線路圖比對/);
  assert.doesNotMatch(html, /高可信|中可信/);
  assert.doesNotMatch(html, /structure-context|match-reason/);
  assert.match(html, /離線隱私模式/);
  assert.match(html, /不會上傳、同步或儲存/);
  assert.match(html, /PCB_Main_v1\.3\.xlsx/);
  assert.match(html, /匯出報表語言/);
  assert.match(html, /中文版/);
  assert.match(html, /English/);
  assert.doesNotMatch(html, /codex-preview|Your site is taking shape/);
});

test("ships real BOM parsing, comparison, and export behavior", async () => {
  const [page, layout, styles, packageJson] = await Promise.all([
    readFile(new URL("../app/page.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/layout.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/globals.css", import.meta.url), "utf8"),
    readFile(new URL("../package.json", import.meta.url), "utf8"),
  ]);
  assert.match(page, /parseWorkbookData/);
  assert.match(page, /analyzeCompanyBomMatrix/);
  assert.match(page, /const \{ book, sheets \} = parseWorkbookData\(data\)/);
  assert.match(page, /待人工確認/);
  assert.match(page, /selectedFields\.some\(\(field\) => item\.fields\.includes\(field\)\)/);
  assert.match(page, /差異項目（可複選）/);
  assert.match(page, /diffGroupMeta/);
  assert.match(page, /DiffDetailDrawer/);
  assert.match(page, /aria-label={`查看\$\{primaryTypeLabel\(item\.primaryType\)\}差異詳細資料`}/);
  assert.match(page, /<th>差異項目<\/th><th>料號異動<\/th>/);
  assert.match(page, /item\.primaryType === "substituteAdded" \? "替料"/);
  assert.match(page, /item\.primaryType === "substituteRemoved" \? "替料"/);
  assert.doesNotMatch(page, /新增料號 · 新版完全新料|刪除料號 · 新版完全移除/);
  assert.doesNotMatch(page, /"數量差異"/);
  assert.match(page, /matchesFilter && matchesImpact/);
  assert.match(page, /fieldFilterOptions/);
  assert.match(page, /componentAdded: "新增"/);
  assert.match(page, /substituteAdded: "新增"/);
  assert.match(page, /componentRemoved: "刪除"/);
  assert.match(page, /substituteRemoved: "刪除"/);
  assert.match(page, /partReplaced: "變更"/);
  assert.match(page, /showConfidence = item\.matchConfidence === "low"/);
  assert.match(page, /exportBomReport/);
  assert.match(page, /exportBomReport\(visible/);
  assert.match(page, /exportLanguage/);
  assert.match(page, /value="zh-TW">中文版/);
  assert.doesNotMatch(page, /exportBomHtmlReport\(visible|匯出 HTML/);
  assert.match(page, /getImageData/);
  assert.match(page, /application\/pdf/);
  assert.match(layout, /lang="zh-Hant"/);
  assert.match(styles, /confidence-popover:focus-within/);
  assert.match(styles, /\.diff-table thead th:nth-child\(1\).*position: sticky/);
  assert.match(styles, /\.detail-drawer/);
  assert.match(packageJson, /"xlsx"/);
  assert.match(packageJson, /"exceljs"/);
  assert.match(packageJson, /"offline:serve"/);
  assert.doesNotMatch(packageJson, /react-loading-skeleton/);
});

test("builds a concise formatted Excel difference report", async () => {
  const item = (part, position) => ({
    ref: position,
    part,
    manufacturerPart: `${part}-MPN`,
    value: "",
    description: "",
    qty: 1,
    positions: [position],
    alternatives: [{ part, manufacturerPart: `${part}-MPN`, manufacturerName: `${part}-MAKER`, description: "", spec: "" }],
    structureKind: "vb-t",
    structurePath: ["69-ROOT", "VB-BOARD-T"],
    structureKey: "1:vb-t:1",
  });
  const replacementDiffs = compareBom([item("OLD-PART", "U20")], [item("NEW-PART", "U20")]);
  const quantityBefore = { ...item("SAME-PART", "U30"), structureKey: "1:vb-t:2" };
  const quantityAfter = { ...item("SAME-PART", "U30"), structureKey: "1:vb-t:2", qty: 2, positions: ["U30", "U31"] };
  const diffs = [...replacementDiffs, ...compareBom([quantityBefore], [quantityAfter])];
  const { buildBomReport, exportCategories } = await import("../app/export-report.ts");
  const workbook = buildBomReport(diffs, "before.xlsx", "after.xlsx");
  const summarySheet = workbook.getWorksheet("差異摘要");
  const dataSheet = workbook.getWorksheet("差異資料");
  assert.equal(summarySheet.views[0].state, "frozen");
  assert.equal(summarySheet.views[0].ySplit, 8);
  assert.equal(summarySheet.views[0].showGridLines, false);
  assert.equal(summarySheet.autoFilter, null);
  assert.equal(summarySheet.getCell("A4").fill.fgColor.argb, "E8F7F0");
  assert.equal(summarySheet.getCell("A7").value, "主要異動");
  assert.equal(summarySheet.getCell("C7").value, "差異項目");
  assert.equal(summarySheet.getCell("E7").value, "料號異動");
  assert.equal(summarySheet.getCell("G7").value, "插件位置差異");
  assert.equal(summarySheet.getCell("I7").value, "舊版料號資訊");
  assert.equal(summarySheet.getCell("N7").value, "新版料號資訊");
  assert.equal(summarySheet.getCell("A8").value, "新增／刪除／變更");
  assert.equal(summarySheet.getCell("E8").value, "料號新增／刪除");
  assert.equal(summarySheet.getCell("I8").value, "舊版料號");
  assert.equal(summarySheet.getCell("J8").value, "舊版製造商料號");
  assert.equal(summarySheet.getCell("K8").value, "舊版製造商");
  assert.equal(summarySheet.getCell("L8").value, "舊版 BOM R欄 TPN");
  assert.equal(summarySheet.getCell("N8").value, "新版料號");
  assert.equal(summarySheet.getCell("O8").value, "新版製造商料號");
  assert.equal(summarySheet.getCell("P8").value, "新版製造商");
  assert.equal(summarySheet.getCell("Q8").value, "新版 BOM R欄 TPN");
  assert.equal(summarySheet.getCell("I7").fill.fgColor.argb, "58677C");
  assert.equal(summarySheet.getCell("N7").fill.fgColor.argb, "2F6BCE");
  assert.match(summarySheet.getCell("A6").value, /淡紅：舊版刪除／停用料號/);
  assert.equal(summarySheet.getCell("I10").fill.fgColor.argb, "FFF0F0");
  assert.equal(summarySheet.getCell("I10").font.color.argb, "B14949");
  assert.equal(summarySheet.getCell("I10").border.left.style, "medium");
  assert.equal(summarySheet.getCell("N10").fill.fgColor.argb, "E8F7F0");
  assert.equal(summarySheet.getCell("N10").font.color.argb, "16835D");
  assert.equal(summarySheet.getCell("N10").border.left.style, "medium");
  assert.equal(dataSheet.views[0].ySplit, 3);
  assert.equal(dataSheet.views[0].showGridLines, false);
  assert.ok(Array.from({ length: 19 }, (_, index) => summarySheet.getColumn(index + 1).width).every((width) => Number.isFinite(width)));
  assert.ok(Array.from({ length: 18 }, (_, index) => dataSheet.getColumn(index + 1).width).every((width) => Number.isFinite(width)));
  assert.ok(dataSheet.getColumn(7).width >= 22 && dataSheet.getColumn(7).width <= 40);
  assert.ok(dataSheet.getColumn(12).width >= 22 && dataSheet.getColumn(12).width <= 40);
  assert.deepEqual(
    dataSheet.model.tables.map((table) => table.name),
    ["BomDiffNewParts", "BomDiffRemovedParts", "BomDiffQuantity"],
  );
  assert.deepEqual(exportCategories(replacementDiffs[0]), ["新版完全新料", "新版完全移除"]);
  assert.deepEqual(exportCategories(diffs[1]), ["數量差異"]);
  const buffer = await workbook.xlsx.writeBuffer();
  const parsed = XLSX.read(buffer, { type: "buffer" });

  assert.deepEqual(parsed.SheetNames, ["差異摘要", "差異資料"]);
  assert.equal(parsed.Sheets["差異摘要"].A1.v, "BOM 差異比較報告");
  assert.equal(parsed.Sheets["差異摘要"].A4.v, "新版完全新料\n1");
  assert.equal(parsed.Sheets["差異摘要"].A9.v, "新版完全新料（1）");
  assert.equal(parsed.Sheets["差異摘要"].A10.v, "變更");
  assert.equal(parsed.Sheets["差異摘要"].E10.v, "＋ NEW-PART\n－ OLD-PART");
  assert.equal(parsed.Sheets["差異摘要"].G10.v, "U20 換料");
  assert.equal(parsed.Sheets["差異摘要"].I10.v, "OLD-PART");
  assert.equal(parsed.Sheets["差異摘要"].J10.v, "OLD-PART-MPN");
  assert.equal(parsed.Sheets["差異摘要"].K10.v, "OLD-PART-MAKER");
  assert.equal(parsed.Sheets["差異摘要"].N10.v, "NEW-PART");
  assert.equal(parsed.Sheets["差異摘要"].O10.v, "NEW-PART-MPN");
  assert.equal(parsed.Sheets["差異摘要"].P10.v, "NEW-PART-MAKER");
  assert.equal(parsed.Sheets["差異摘要"].R10.v, "1 → 1");
  assert.equal(parsed.Sheets["差異摘要"].A13.v, "新版完全移除（1）");
  assert.equal(parsed.Sheets["差異摘要"].A17.v, "數量差異（1）");
  assert.equal(parsed.Sheets["差異摘要"].A19.v, "TPN 差異（0）");
  assert.equal(parsed.Sheets["差異摘要"].A21.v, "架構／製程變更（0）");
  assert.equal(parsed.Sheets["差異摘要"].A23.v, "僅插件位置差異（0）");
  assert.equal(parsed.Sheets["差異資料"].A1.v, "BOM 差異資料｜依分類分表");
  assert.equal(parsed.Sheets["差異資料"].A4.v, "新版完全新料（1）");
  assert.equal(parsed.Sheets["差異資料"].A5.v, "分類");
  assert.equal(parsed.Sheets["差異資料"].B5.v, "主要異動");
  assert.equal(parsed.Sheets["差異資料"].F5.v, "舊版主件料號");
  assert.equal(parsed.Sheets["差異資料"].I5.v, "舊版 BOM R欄 TPN");
  assert.equal(parsed.Sheets["差異資料"].K5.v, "新版主件料號");
  assert.equal(parsed.Sheets["差異資料"].R5.v, "數量變化");
  assert.equal(parsed.Sheets["差異資料"].A6.v, "新版完全新料");
  assert.equal(parsed.Sheets["差異資料"].D6.v, "＋ NEW-PART");
  assert.equal(parsed.Sheets["差異資料"].E6.v, "U20 換料");
  assert.equal(parsed.Sheets["差異資料"].R6.v, "Same");
  assert.equal(parsed.Sheets["差異資料"].A8.v, "新版完全移除（1）");
  assert.equal(parsed.Sheets["差異資料"].A10.v, "新版完全移除");
  assert.equal(parsed.Sheets["差異資料"].D10.v, "－ OLD-PART");
  assert.equal(parsed.Sheets["差異資料"].R14.v, "Increase");
  assert.equal(dataSheet.getCell("F5").fill.fgColor.argb, "58677C");
  assert.equal(dataSheet.getCell("K5").fill.fgColor.argb, "2F6BCE");
  assert.equal(dataSheet.getCell("F6").fill.fgColor.argb, "F2F4F7");
  assert.equal(dataSheet.getCell("K6").fill.fgColor.argb, "E8F7F0");
  assert.equal(dataSheet.getCell("K6").font.color.argb, "16835D");
  assert.equal(dataSheet.getCell("F10").fill.fgColor.argb, "FFF0F0");
  assert.equal(dataSheet.getCell("F10").font.color.argb, "B14949");
  assert.equal(dataSheet.getCell("K10").fill.fgColor.argb, "EDF4FF");
  assert.ok(buffer.byteLength > 5_000);

  const englishWorkbook = buildBomReport(diffs, "before.xlsx", "after.xlsx", {}, "en");
  const englishSummary = englishWorkbook.getWorksheet("Difference Summary");
  const englishData = englishWorkbook.getWorksheet("Difference Data");
  assert.equal(englishSummary.getCell("A1").value, "BOM Difference Comparison Report");
  assert.equal(englishSummary.getCell("A4").value, "Completely New Parts\n1");
  assert.equal(englishSummary.getCell("A7").value, "Primary Change");
  assert.equal(englishSummary.getCell("I7").value, "Old Part Information");
  assert.equal(englishSummary.getCell("N7").value, "New Part Information");
  assert.equal(englishSummary.getCell("A10").value, "Changed");
  assert.equal(englishSummary.getCell("G10").value, "U20 Replacement");
  assert.equal(englishData.getCell("A1").value, "BOM Difference Data | Grouped by Category");
  assert.equal(englishData.getCell("A5").value, "Category");
  assert.equal(englishData.getCell("F5").value, "Old Company Part Number");
  assert.equal(englishData.getCell("K5").value, "New Company Part Number");
  assert.equal(englishData.getCell("A6").value, "Completely New Parts");
});

test("classifies substitute-only export rows without calling them completely new or removed", async () => {
  const item = (parts, positions = ["U20"]) => ({
    ref: "00A",
    part: parts[0],
    manufacturerPart: `${parts[0]}-MPN`,
    value: "",
    description: "",
    qty: positions.length,
    positions,
    alternatives: parts.map((part) => ({ part, manufacturerPart: `${part}-MPN`, manufacturerName: "MAKER", description: "", spec: "" })),
    structureKind: "vb-t",
    structurePath: ["69-ROOT", "VB-BOARD-T"],
    structureKey: "1:vb-t:1",
  });
  const added = compareBom([item(["MAIN"])], [item(["MAIN", "ALT"])])[0];
  const removed = compareBom([item(["MAIN", "ALT"])], [item(["MAIN"])])[0];
  const entirelyNew = compareBom([], [item(["NEW"], ["U10"])])[0];
  const entirelyRemoved = compareBom([item(["OLD"], ["U11"])], [])[0];
  const quantity = compareBom([item(["QTY"], ["U12"])], [item(["QTY"], ["U12", "U13"])])[0];
  const positionOnly = compareBom([item(["MOVE"], ["U14"])], [item(["MOVE"], ["U15"])])[0];
  const { buildBomReport, exportCategories } = await import("../app/export-report.ts");

  assert.deepEqual(exportCategories(added), ["新增替代"]);
  assert.deepEqual(exportCategories(removed), ["刪除替代"]);
  const workbook = buildBomReport(
    [positionOnly, entirelyRemoved, removed, quantity, entirelyNew, added],
    "before.xlsx",
    "after.xlsx",
  );
  const summarySheet = workbook.getWorksheet("差異摘要");
  assert.deepEqual(
    summarySheet.getColumn(1).values.filter((value) => typeof value === "string" && /（\d+）$/.test(value)),
    ["新版完全新料（1）", "新增替代（1）", "新版完全移除（1）", "刪除替代（1）", "數量差異（1）", "TPN 差異（0）", "架構／製程變更（0）", "僅插件位置差異（1）"],
  );
  const substituteSectionRow = summarySheet.getColumn(1).values.findIndex((value) => value === "新增替代（1）");
  assert.equal(summarySheet.getCell(`I${substituteSectionRow + 1}`).value, "MAIN");
  assert.equal(summarySheet.getCell(`N${substituteSectionRow + 1}`).value, "MAIN");
  assert.equal(summarySheet.getCell(`N${substituteSectionRow + 2}`).value, "ALT");
  assert.equal(summarySheet.getCell(`O${substituteSectionRow + 2}`).value, "ALT-MPN");
  assert.equal(summarySheet.getCell(`N${substituteSectionRow + 1}`).fill.fgColor.argb, "EDF4FF");
  assert.equal(summarySheet.getCell(`N${substituteSectionRow + 2}`).fill.fgColor.argb, "E8F7F0");
  assert.equal(summarySheet.getCell(`N${substituteSectionRow + 2}`).font.color.argb, "16835D");
  const dataSheet = workbook.getWorksheet("差異資料");
  const dataSectionRows = dataSheet.getColumn(1).values
    .map((value, row) => ({ value, row }))
    .filter(({ value }) => typeof value === "string" && /（\d+）$/.test(value));
  assert.deepEqual(
    dataSectionRows.map(({ value }) => value),
    ["新版完全新料（1）", "新增替代（1）", "新版完全移除（1）", "刪除替代（1）", "數量差異（1）", "僅插件位置差異（1）"],
  );
  assert.deepEqual(
    dataSectionRows.map(({ row }) => dataSheet.getCell(`A${row + 2}`).value),
    ["新版完全新料", "新增替代", "新版完全移除", "刪除替代", "數量差異", "僅插件位置差異"],
  );
  const addedSubstituteDataRow = dataSectionRows.find(({ value }) => value === "新增替代（1）").row + 2;
  assert.ok(dataSheet.getRow(addedSubstituteDataRow).height > 36);
  assert.ok(dataSheet.getRow(addedSubstituteDataRow).height <= 96);
  assert.deepEqual(
    dataSheet.model.tables.map((table) => table.name),
    [
      "BomDiffNewParts",
      "BomDiffAddedSubstitutes",
      "BomDiffRemovedParts",
      "BomDiffRemovedSubstitutes",
      "BomDiffQuantity",
      "BomDiffPositions",
    ],
  );
});

test("includes unchanged before and after BOM worksheets in the Excel report", async () => {
  const sourceWorkbook = new ExcelJS.Workbook();
  const sourceSheet = sourceWorkbook.addWorksheet("ProductStructure", { views: [{ state: "frozen", ySplit: 1 }] });
  sourceSheet.getCell("A1").value = "項次";
  sourceSheet.getCell("A1").font = { name: "Microsoft JhengHei", bold: true, color: { argb: "FFFFFF" } };
  sourceSheet.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "1F4E78" } };
  sourceSheet.getCell("A2").value = "00A";
  sourceSheet.getCell("B2").value = { formula: "1+1", result: 2 };
  sourceSheet.mergeCells("C1:D1");
  sourceSheet.getCell("C1").value = "原始合併標題";
  sourceSheet.getColumn(2).width = 24;
  sourceSheet.getRow(2).height = 31;
  sourceSheet.getCell("A3").value = "隱藏原始列";
  sourceSheet.getRow(3).hidden = true;
  sourceSheet.addConditionalFormatting({ ref: "B2", rules: [{ type: "cellIs", operator: "greaterThan", formulae: [1], style: { font: { color: { argb: "FFFF0000" } } } }] });
  const sourceBuffer = await sourceWorkbook.xlsx.writeBuffer();
  const originalData = new Uint8Array(sourceBuffer).slice().buffer;
  const source = { fileName: "source.xlsx", sheetName: "ProductStructure", data: originalData, matrix: [["項次", "數量"], ["00A", 2]] };

  const { buildBomReportWithOriginals } = await import("../app/export-report.ts");
  const workbook = await buildBomReportWithOriginals([], "before.xlsx", "after.xlsx", { originalBefore: source, originalAfter: source });
  assert.deepEqual(workbook.worksheets.map((sheet) => sheet.name), ["差異摘要", "差異資料", "舊版原始 BOM", "新版原始 BOM"]);

  const beforeSheet = workbook.getWorksheet("舊版原始 BOM");
  assert.equal(beforeSheet.getCell("A2").value, "00A");
  assert.equal(beforeSheet.getCell("B2").formula, "1+1");
  assert.equal(beforeSheet.getCell("A1").fill.fgColor.argb, "1F4E78");
  assert.equal(beforeSheet.getColumn(2).width, 24);
  assert.equal(beforeSheet.getRow(2).height, 31);
  assert.equal(beforeSheet.getRow(3).hidden, true);
  assert.equal(beforeSheet.getCell("D1").isMerged, true);
  assert.equal(beforeSheet.views[0].state, "frozen");
  assert.equal(beforeSheet.model.conditionalFormattings.length, 1);
  assert.equal(workbook.getWorksheet("差異摘要").getCell("C2").value.hyperlink, "#'舊版原始 BOM'!A1");
  assert.equal(workbook.getWorksheet("差異摘要").getCell("K2").value.hyperlink, "#'新版原始 BOM'!A1");
  assert.equal(workbook.getWorksheet("差異資料").getCell("B2").value.hyperlink, "#'舊版原始 BOM'!A1");

  const reportBuffer = await workbook.xlsx.writeBuffer();
  const parsed = XLSX.read(reportBuffer, { type: "buffer" });
  assert.equal(parsed.Sheets["舊版原始 BOM"].A2.v, "00A");
  assert.equal(parsed.Sheets["新版原始 BOM"].A2.v, "00A");
});

test("builds a standalone offline HTML report", async () => {
  const item = (part, position) => ({
    ref: position,
    part,
    manufacturerPart: `${part}-MPN`,
    value: "",
    description: "",
    qty: 1,
    positions: [position],
    alternatives: [{ part, manufacturerPart: `${part}-MPN`, manufacturerName: `${part}-MAKER`, description: "", spec: "" }],
    structureKind: "vb-t",
    structurePath: ["69-ROOT", "VB-BOARD-T"],
    structureKey: "1:vb-t:1",
  });
  const diffs = compareBom([item("OLD<&", "U20")], [item("NEW", "U20")]);
  const { buildBomHtmlReport } = await import("../app/export-html-report.ts");
  const html = buildBomHtmlReport(diffs, "before.xlsx", "after.xlsx");

  assert.match(html, /^<!doctype html>/);
  assert.match(html, /BOM 版本差異報告/);
  assert.match(html, /差異明細/);
  assert.match(html, /新版新料／移除清單/);
  assert.doesNotMatch(html, /<h2>待人工確認<\/h2>|<h2>匯入警告<\/h2>/);
  assert.match(html, /製造廠商/);
  assert.match(html, /所屬架構/);
  assert.match(html, /69-ROOT › VB-BOARD-T/);
  assert.match(html, /class="type replacement">變更/);
  assert.match(html, /舊版 BOM/);
  assert.match(html, /新版 BOM/);
  assert.doesNotMatch(html, /前版|後版|新增元件|移除元件/);
  assert.match(html, /Content-Security-Policy/);
  assert.match(html, /OLD&lt;&amp;/);
  assert.doesNotMatch(html, /class="confidence high"|class="confidence medium"/);
  assert.doesNotMatch(html, /https?:\/\//);
});

test("finds exact schematic reference designators without partial matches", () => {
  const boxes = [
    { page: 1, text: "U45", x: 10, y: 20, width: 30, height: 12 },
    { page: 1, text: "U450", x: 50, y: 20, width: 35, height: 12 },
    { page: 2, text: "R10, U45", x: 10, y: 40, width: 60, height: 12 },
    { page: 3, text: "U 45", x: 10, y: 60, width: 30, height: 12 },
  ];
  assert.equal(normalizeReference(" u 45 "), "U45");
  assert.deepEqual(findReferenceHits(boxes, "U45").map((hit) => hit.page), [1, 2, 3]);
});

test("builds a direct reference index and merges adjacent PDF text blocks", () => {
  const boxes = [
    { page: 1, text: "R10, U45", x: 10, y: 10, width: 60, height: 10 },
    { page: 1, text: "U", x: 10, y: 30, width: 7, height: 10 },
    { page: 1, text: "102", x: 19, y: 30, width: 18, height: 10 },
    { page: 1, text: "A", x: 39, y: 30, width: 7, height: 10 },
    { page: 1, text: "U", x: 10, y: 50, width: 7, height: 10 },
    { page: 1, text: "99", x: 100, y: 50, width: 13, height: 10 },
  ];
  const index = buildPageReferenceIndex(boxes);
  assert.equal(lookupReferenceHits(index, "R10").length, 1);
  assert.equal(lookupReferenceHits(index, "U45").length, 1);
  assert.equal(lookupReferenceHits(index, "U102A").length, 1);
  assert.equal(lookupReferenceHits(index, "U99").length, 0);
  assert.ok(lookupReferenceHits(index, "U102A")[0].width > 30);
});

test("prefers exact reference labels over references embedded in net names", () => {
  const exact = { page: 1, text: "U45", x: 20, y: 30, width: 22, height: 14, rotation: 0 };
  const netName = { page: 1, text: "NET_U45_SIGNAL", x: 20, y: 52, width: 90, height: 12, rotation: 0 };
  const index = buildPageReferenceIndex([exact, netName]);
  assert.deepEqual(lookupReferenceHits(index, "U45"), [exact]);
});

test("merges vertically split and rotated schematic references", () => {
  const boxes = [
    { page: 2, text: "U", x: 20, y: 10, width: 10, height: 7, rotation: 90 },
    { page: 2, text: "45", x: 20, y: 19, width: 10, height: 14, rotation: 90 },
    { page: 2, text: "A", x: 50, y: 10, width: 10, height: 7, rotation: 270 },
    { page: 2, text: "102", x: 50, y: 19, width: 10, height: 18, rotation: 270 },
    { page: 2, text: "U", x: 50, y: 39, width: 10, height: 7, rotation: 270 },
  ];
  const index = buildPageReferenceIndex(boxes);
  assert.equal(lookupReferenceHits(index, "U45").length, 1);
  assert.equal(lookupReferenceHits(index, "U102A").length, 1);
  assert.equal(lookupReferenceHits(index, "U45")[0].rotation, 90);
});

test("centers schematic reference hits and clamps page-edge positions", () => {
  const center = centeredPdfHitScroll(
    { page: 1, text: "U45", x: 900, y: 600, width: 20, height: 10 },
    1.35, 400, 300, 1200, 900,
  );
  assert.deepEqual(center, { left: 710, top: 455 });

  const topLeft = centeredPdfHitScroll(
    { page: 1, text: "U1", x: 5, y: 5, width: 10, height: 10 },
    1.35, 400, 300, 1200, 900,
  );
  assert.deepEqual(topLeft, { left: 0, top: 0 });

  const bottomRight = centeredPdfHitScroll(
    { page: 1, text: "U99", x: 1180, y: 880, width: 20, height: 20 },
    1.35, 400, 300, 1200, 900,
  );
  assert.deepEqual(bottomRight, { left: 800, top: 600 });
});

test("centers from the rendered highlight rectangle after layout and scaling", () => {
  const centered = centeredRenderedHitScroll(
    { left: 120, top: 80 },
    { left: 30, top: 20, width: 500, height: 320 },
    { left: 610, top: 390, width: 24, height: 16 },
    { left: 900, top: 700 },
  );
  assert.deepEqual(centered, { left: 462, top: 298 });
});

test("builds a schematic report plan from BOM placement differences", () => {
  const item = (positions, part = "PART-A") => ({
    ref: "001", part, manufacturerPart: `${part}-MPN`, manufacturerName: "", value: "", description: "", qty: positions.length,
    positions,
    alternatives: [{ part, manufacturerPart: `${part}-MPN`, description: "", spec: "" }],
  });
  const relocation = compareBom([item(["U1"])], [item(["U2"])])[0];
  const replacement = compareBom([item(["U3"], "PART-OLD")], [item(["U3"], "PART-NEW")])[0];
  const plan = createSchematicReportPlan([relocation, replacement]);
  assert.deepEqual(plan.map((entry) => entry.reference), ["U1", "U2", "U3"]);
  assert.ok(plan.find((entry) => entry.reference === "U1").changeLabels.includes("移除插件位置"));
  assert.ok(plan.find((entry) => entry.reference === "U2").changeLabels.includes("新增插件位置"));
  assert.ok(plan.find((entry) => entry.reference === "U3").changeLabels.includes("更換料號"));
});

test("progressively indexes and caches schematic PDFs", async () => {
  const viewer = await readFile(new URL("../app/pdf-schematic-viewer.tsx", import.meta.url), "utf8");
  const styles = await readFile(new URL("../app/globals.css", import.meta.url), "utf8");
  assert.match(viewer, /Map<File, IndexEntry>/);
  assert.match(viewer, /已處理.*processedPages/);
  assert.match(viewer, /背景索引中/);
  assert.match(viewer, /已使用索引快取/);
  assert.match(viewer, /notify\(entry\)/);
  assert.match(viewer, /Promise\.all\(pageNumbers/);
  assert.match(viewer, /lookupReferenceHits/);
  assert.match(viewer, /MAX_CACHED_PDFS = 2/);
  assert.match(viewer, /entry\.document\?\.destroy/);
  assert.match(viewer, /syncState\.source === side/);
  assert.match(viewer, /onScaleChange/);
  assert.match(viewer, /centeredPdfHitScroll/);
  assert.match(viewer, /centeredRenderedHitScroll/);
  assert.match(viewer, /highlightRef/);
  assert.match(viewer, /currentHit\?\.page === pageNumber/);
  assert.match(viewer, /scheduleCenterCurrentHit/);
  assert.match(viewer, /重新置中/);
  assert.match(viewer, /pageElement\.style\.width/);
  assert.match(styles, /\.sheet-preview\s*\{[^}]*overflow:\s*hidden/);
  assert.match(styles, /\.pdf-viewer\s*\{[^}]*min-height:\s*0[^}]*overflow:\s*hidden/);
  assert.match(styles, /\.pdf-page-scroll\s*\{[^}]*min-height:\s*0[^}]*overflow:\s*auto/);
});

test("maps company columns by header name and groups substitute parts", () => {
  const header = Array(18).fill("");
  header[0] = "項次"; header[1] = "料號"; header[4] = "組成用量"; header[5] = "插件位置"; header[10] = "製造商名稱"; header[15] = "製造商料號"; header[17] = "客戶料號";
  const main = Array(18).fill("");
  main[0] = "00K"; main[1] = "PREFIX-123456789012"; main[4] = "2.0/1"; main[5] = "U20, U31"; main[10] = "Texas Instruments"; main[15] = "MPN-MAIN"; main[17] = "CUSTOMER-OLD";
  const substitute = Array(18).fill("");
  substitute[1] = "INTERNAL-ALT"; substitute[4] = "2.0/1"; substitute[10] = "Nexperia"; substitute[15] = "MPN-ALT"; substitute[17] = "CUSTOMER-NEW";
  const matrix = [header, main, substitute];

  assert.equal(findCompanyHeader(matrix), 0);
  assert.equal(parseQuantity("4.0/1"), 4);
  assert.equal(parseQuantity("197.0/100000"), 0.00197);
  assert.equal(parseQuantity("1.97/1000"), 0.00197);
  assert.equal(canonicalPartNumber("PREFIX-123456789012"), "123456789012");
  const parsed = parseCompanyBomMatrix(matrix);
  assert.equal(parsed.length, 1);
  assert.equal(parsed[0].ref, "00K");
  assert.equal(parsed[0].qty, 2);
  assert.deepEqual(parsed[0].positions, ["U20", "U31"]);
  assert.deepEqual(parsed[0].alternatives.map((part) => part.part), ["123456789012", "INTERNAL-ALT"]);
  assert.deepEqual(parsed[0].alternatives.map((part) => part.manufacturerName), ["Texas Instruments", "Nexperia"]);
  assert.deepEqual(parsed[0].alternatives.map((part) => part.rdCustomerPartNumbers), [["CUSTOMER-OLD"], ["CUSTOMER-NEW"]]);
  assert.deepEqual(parsed[0].rdCustomerPartNumbers, ["CUSTOMER-OLD", "CUSTOMER-NEW"]);
});

test("normalizes 1G usage quantities by their denominator instead of counting material labels as placements", async () => {
  const header = ["項次", "主件料號", "組成用量", "插件位置", "製造廠商料號"];
  const before = parseCompanyBomMatrix([
    header,
    ["05M", "1G03-002B000", "197.0/100000", "SOLDERPASTE", "INDIUM10.8HF"],
  ]);
  const equivalentAfter = parseCompanyBomMatrix([
    header,
    ["05M", "1G03-002B000", "1.97/1000", "SOLDERPASTE", "INDIUM10.8HF"],
  ]);
  const changedAfter = parseCompanyBomMatrix([
    header,
    ["05M", "1G03-002B000", "250.0/100000", "SOLDERPASTE", "INDIUM10.8HF"],
  ]);

  assert.equal(before[0].qty, 0.00197);
  assert.equal(before[0].declaredQty, 0.00197);
  assert.equal(compareBom(before, equivalentAfter)[0].kind, "same");
  assert.equal(bomQuantityChanged(before[0], equivalentAfter[0]), false);
  assert.equal(bomQuantityChanged(before[0], changedAfter[0]), true);
  const changedDiff = compareBom(before, changedAfter)[0];
  assert.ok(changedDiff.categories.includes("changed"));

  const { buildBomReport, exportCategories } = await import("../app/export-report.ts");
  assert.deepEqual(exportCategories(changedDiff), ["數量差異"]);
  const report = buildBomReport([changedDiff], "before.xlsx", "after.xlsx");
  const dataSheet = report.getWorksheet("差異資料");
  const sectionRow = dataSheet.getColumn(1).values.findIndex((value) => value === "數量差異（1）");
  assert.equal(dataSheet.getCell(`P${sectionRow + 2}`).value, 0.00197);
  assert.equal(dataSheet.getCell(`Q${sectionRow + 2}`).value, 0.0025);
  assert.equal(dataSheet.getCell(`P${sectionRow + 2}`).numFmt, "#,##0.##########");
  assert.equal(dataSheet.getCell(`Q${sectionRow + 2}`).numFmt, "#,##0.##########");

  const audit = analyzeCompanyBomMatrix([
    header,
    ["05M", "1G03-002B000", "197.0/100000", "SOLDERPASTE", "INDIUM10.8HF"],
  ], 0);
  assert.equal(audit.audit.issues.some((issue) => issue.code === "quantity-mismatch"), false);
});

test("defaults to the numbered company BOM headers", () => {
  const header = ["1.項次", "2.主件料號", "5組成用量", "6插件位置", "11製造廠商", "16製造廠商料號"];
  const mapping = detectCompanyColumns(header);
  assert.equal(mapping.part, 1);
  assert.equal(mapping.manufacturerName, 4);
  assert.equal(mapping.manufacturerPart, 5);
});

test("defaults to the company field names without numeric prefixes", () => {
  const header = ["項次", "主件料號", "組成用量", "插件位置", "製造廠商", "製造廠商料號"];
  const mapping = detectCompanyColumns(header);
  assert.equal(mapping.part, 1);
  assert.equal(mapping.manufacturerName, 4);
  assert.equal(mapping.manufacturerPart, 5);
});

test("automatically selects the S column named 廠商/料號及型號", () => {
  const header = Array(19).fill("");
  header[0] = "項次";
  header[1] = "主件料號";
  header[4] = "組成用量";
  header[5] = "插件位置";
  header[17] = "對應客戶料號";
  header[18] = "廠商/料號及型號";
  const mapping = detectCompanyColumns(header);
  assert.equal(mapping.customerPartNumbers, 17);
  assert.equal(mapping.customerPartAssociations, 18);

  header[18] = "廠商／料號及型號";
  assert.equal(detectCompanyColumns(header).customerPartAssociations, 18);
});

test("reads HTML content saved with an .xls extension and still detects the S column", () => {
  const html = `<!doctype html><html><head><meta charset="utf-8"></head><body><table>
    <tr><th>項次</th><th>主件料號</th><th>組成用量</th><th>插件位置</th><th>製造廠商料號</th><th>對應客戶料號</th><th>廠商/料號及型號</th></tr>
    <tr><td>001</td><td>0704-06HL0ZY</td><td>1</td><td>Q33</td><td>SSM3K361R_LXHF</td><td>2156321-00-A</td><td>TOSHIBA/SSM3K361R_LXHF</td></tr>
  </table></body></html>`;
  const parsedWorkbook = parseWorkbookData(new TextEncoder().encode(html).buffer);
  assert.equal(parsedWorkbook.sourceFormat, "html-xls");
  assert.equal(parsedWorkbook.sheets.length, 1);
  const header = parsedWorkbook.sheets[0].matrix[0];
  assert.equal(detectCompanyColumns(header).customerPartAssociations, 6);
  const items = parseCompanyBomMatrix(parsedWorkbook.sheets[0].matrix);
  assert.equal(items[0].alternatives[0].part, "0704-06HL0ZY");
  assert.equal(items[0].alternatives[0].rdCustomerPartAssociations[0].manufacturerPart, "SSM3K361R_LXHF");
});

test("reads a real legacy BIFF8 .xls workbook", () => {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    ["項次", "主件料號", "組成用量", "插件位置", "製造廠商料號", "對應客戶料號", "廠商／料號及型號"],
    ["001", "0704-06HL0ZY", 1, "Q33", "SSM3K361R_LXHF", "2156321-00-A", "TOSHIBA/SSM3K361R_LXHF"],
  ]), "ProductStructureReport");
  const binary = XLSX.write(workbook, { type: "array", bookType: "biff8" });
  const parsedWorkbook = parseWorkbookData(binary);
  assert.equal(parsedWorkbook.sourceFormat, "excel-binary");
  assert.equal(parsedWorkbook.sheets[0].name, "ProductStructureReport");
  assert.equal(detectCompanyColumns(parsedWorkbook.sheets[0].matrix[0]).customerPartAssociations, 6);
  assert.equal(parseCompanyBomMatrix(parsedWorkbook.sheets[0].matrix)[0].positions[0], "Q33");
});

test("compresses setup controls and strengthens added and removed states", async () => {
  const [page, styles] = await Promise.all([
    readFile(new URL("../app/page.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/globals.css", import.meta.url), "utf8"),
  ]);
  assert.match(page, /comparison-setup/);
  assert.match(page, /setRulesOpen/);
  assert.match(page, /setMvaOpen/);
  assert.doesNotMatch(page, /className="new-compare"/);
  assert.match(page, /舊版不存在/);
  assert.match(page, /新版已移除/);
  assert.match(page, /part-absence/);
  assert.match(styles, /\.comparison-setup/);
  assert.match(styles, /\.group-status-card/);
  assert.match(styles, /\.part-absence/);
  assert.match(styles, /\.part-absence \{[^}]*border: 1px dashed #cfd6df/);
  assert.doesNotMatch(styles, /\.part-absence\.not-existed[^}]*#c8e7d9/);
  assert.doesNotMatch(styles, /\.part-absence\.removed[^}]*#efcece/);
  assert.match(styles, /\.primary-badge[^}]*min-width: 60px/);
});

test("shows explicit manual review and enlarged schematic controls", async () => {
  const page = await readFile(new URL("../app/page.tsx", import.meta.url), "utf8");
  const styles = await readFile(new URL("../app/globals.css", import.meta.url), "utf8");
  assert.match(page, /人工待審核/);
  assert.match(page, /存在多個接近候選/);
  assert.match(page, /系統判斷依據/);
  assert.match(page, /不確認的風險/);
  assert.match(page, /建議確認順序/);
  assert.match(page, /顯示全部 \$\{positions\.length\} 個位置/);
  assert.match(page, /previewLimit = 6/);
  assert.match(styles, /\.position-overflow/);
  assert.match(styles, /\.review-guidance-grid/);
  assert.match(page, /reviewDiffs\.length/);
  assert.match(page, /setImpactFilter\(impactFilter === "review" \? "all" : "review"\)/);
  assert.match(page, /放大顯示/);
  assert.match(page, /退出放大/);
  assert.match(page, /匯出線路圖報告/);
  assert.match(page, /exportSchematicPdfReport/);
  assert.match(styles, /\.schematic-panel\.expanded/);
  assert.match(styles, /\.schematic-report-progress/);
});

test("filters and sorts customer BOM TPN mapping rows by status", async () => {
  const [page, styles] = await Promise.all([
    readFile(new URL("../app/page.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/globals.css", import.meta.url), "utf8"),
  ]);
  assert.match(page, /CustomerStatusFilter/);
  assert.match(page, /全部狀態/);
  assert.match(page, /需處理/);
  assert.match(page, /異常優先/);
  assert.match(page, /依狀態名稱/);
  assert.match(page, /依公司料號/);
  assert.match(page, /customerStatusPriority/);
  assert.match(page, /匯入 TPN/);
  assert.doesNotMatch(page, /> 客戶 BOM TPN 對應 <span/);
  assert.match(styles, /\.customer-table-controls/);
  assert.match(styles, /\.customer-filter-summary/);
});

test("blocks customer BOM import until the matching company BOM is imported", async () => {
  const page = await readFile(new URL("../app/page.tsx", import.meta.url), "utf8");
  assert.match(page, /function hasImportedCompanyBom/);
  assert.match(page, /beforeAudit && before\.length/);
  assert.match(page, /afterAudit && after\.length/);
  assert.match(page, /請先匯入\$\{version\}公司 BOM，才能匯入\$\{version\}客戶 BOM。/);
  assert.match(page, /requestCustomerBomImport\("before"\)/);
  assert.match(page, /requestCustomerBomImport\("after"\)/);
});

test("keeps parsing after columns are inserted, removed, or reordered", () => {
  const header = ["備註", "製造商名稱", "插件位置", "料號", "項次", "組成用量", "製造商料號", "品名", "規格"];
  const main = ["保留", "Texas Instruments", "U20, U31", "PREFIX-123456789012", "00K", "2.0/1", "MPN-MAIN", "Logic", "QFN"];
  const substitute = ["", "Nexperia", "", "INTERNAL-ALT", "", "2.0/1", "MPN-ALT", "Alternative", ""];
  const parsed = parseCompanyBomMatrix([["BOM 匯出報表"], header, main, substitute]);

  assert.equal(findCompanyHeader([["BOM 匯出報表"], header]), 1);
  assert.equal(parsed.length, 1);
  assert.equal(parsed[0].ref, "00K");
  assert.equal(parsed[0].qty, 2);
  assert.deepEqual(parsed[0].positions, ["U20", "U31"]);
  assert.deepEqual(parsed[0].alternatives.map((part) => part.part), ["123456789012", "INTERNAL-ALT"]);
  assert.deepEqual(parsed[0].alternatives.map((part) => part.manufacturerPart), ["MPN-MAIN", "MPN-ALT"]);
  assert.deepEqual(parsed[0].alternatives.map((part) => part.manufacturerName), ["Texas Instruments", "Nexperia"]);
});

test("keeps repeated item numbers separate across VB-D, 60, VB-T, and PCB branches", () => {
  const header = ["項次", "料號", "數量", "插件位置", "製造商料號"];
  const matrix = [
    header,
    ["", "69 - ROOT000001", "1", "", "ROOT-MPN"],
    ["001", "|---VBBOARD00002D", "1", "", "VB-D-MPN"],
    ["00U", "| |---D-MAIN-000001", "1", "U1", "D-MAIN-MPN"],
    ["", "| | |---D-ALT-0000001", "1", "", "D-ALT-MPN"],
    ["014", "|---60-BOARD000001", "1", "", "60-MPN"],
    ["00U", "| |---M-MAIN-000001", "1", "U1", "M-MAIN-MPN"],
    ["", "| | |---M-ALT-0000001", "1", "", "M-ALT-MPN"],
    ["002", "|---VB - BOARD00001T", "1", "", "VB-T-MPN"],
    ["00U", "| |---T-MAIN-000001", "1", "U1", "T-MAIN-MPN"],
    ["", "| | |---T-ALT-0000001", "1", "", "T-ALT-MPN"],
    ["00U", "| |---081234567890", "1", "U1", "PCB-MPN"],
    ["", "| | |---S-08123456789", "1", "", "PCB-ALT-MPN"],
  ];

  const parsed = parseCompanyBomMatrix(matrix);
  const repeated = parsed.filter((item) => item.ref === "00U");
  assert.equal(parsed.length, 8);
  assert.equal(repeated.length, 4);
  assert.deepEqual(repeated.map((item) => item.structureKind), ["vb-d", "board60", "vb-t", "pcb"]);
  assert.deepEqual(repeated.map((item) => item.alternatives.length), [2, 2, 2, 2]);
  assert.match(repeated[0].structurePath.join(" > "), /69-ROOT000001 > VBBOARD00002D/);
  assert.match(repeated[1].structurePath.join(" > "), /69-ROOT000001 > 60-BOARD000001/);
  assert.equal(parsed.find((item) => item.structureKind === "root69")?.part, canonicalPartNumber("69 - ROOT000001"));
  assert.equal(parsed.find((item) => item.structureKind === "root69")?.ref, "69-架構-1");
  assert.equal(parsed.find((item) => item.structureKind === "pcb")?.part, "081234567890");

  const audit = analyzeCompanyBomMatrix(matrix, 0).audit;
  assert.equal(audit.positionCount, 4);
  assert.ok(!audit.issues.some((issue) => issue.code === "duplicate-position"));
  assert.ok(!audit.issues.some((issue) => issue.code === "part-collision"));
});

test("still warns when a placement repeats inside the same structure branch", () => {
  const matrix = [
    ["項次", "料號", "數量", "插件位置"],
    ["000", "69-ROOT000001", "1", ""],
    ["001", "VB-BOARD00001T", "1", ""],
    ["00U", "T-MAIN-000001", "1", "U1"],
    ["00V", "T-MAIN-000002", "1", "U1"],
  ];
  const result = analyzeCompanyBomMatrix(matrix, 0);
  assert.ok(result.audit.issues.some((issue) => issue.code === "duplicate-position"));
});

test("uses parent structure rows only for grouping and never reports them as differences", () => {
  const header = ["項次", "料號", "數量", "插件位置"];
  const bom = (rootPart, tPart, dPart) => parseCompanyBomMatrix([
    header,
    ["000", rootPart, "1", ""],
    ["001", "VB-BOARD00001T", "1", ""],
    ["00U", tPart, "1", "U1"],
    ["002", "VB-BOARD00002D", "1", ""],
    ["00U", dPart, "1", "U1"],
  ]);
  const before = bom("69-ROOTOLD00001", "PART-T-SAME", "PART-D-OLD");
  const after = bom("69-ROOTNEW00001", "PART-T-SAME", "PART-D-NEW");
  const diffs = compareBom(before, after);

  const sameT = diffs.find((diff) => diff.before?.structureKind === "vb-t" && diff.before?.ref === "00U");
  assert.equal(sameT?.kind, "same");
  assert.ok(!diffs.some((diff) => diff.before?.structureKind === "vb-t" && diff.after?.structureKind === "vb-d"));
  assert.ok(!diffs.some((diff) => diff.before?.isStructure || diff.after?.isStructure));
  assert.ok(!diffs.some((diff) => diff.before?.structureKind === "root69" || diff.after?.structureKind === "root69"));
  const changed = diffs.filter((diff) => diff.categories.length > 0);
  assert.equal(changed.length, 1);
  assert.equal(changed[0].primaryType, "partReplaced");
});

test("detects the same part moving between SMT and DIP structures", async () => {
  const header = ["項次", "料號", "數量", "插件位置", "製造商料號"];
  const bom = (branch) => parseCompanyBomMatrix([
    header,
    ["000", "69G14LM12A01", "1", "", "ROOT-MPN"],
    ["001", branch, "1", "", "BRANCH-MPN"],
    ["00A", "0500-04WB0ZY", "1", "U20", "FLASH-MPN"],
  ]);
  const before = bom("VBG14LM12A01T");
  const after = bom("VBG14LM12A01D");
  const beforePart = before.find((item) => item.ref === "00A");
  const afterPart = after.find((item) => item.ref === "00A");
  assert.equal(bomProcessKind(beforePart), "SMT");
  assert.equal(bomProcessKind(afterPart), "DIP");

  const moved = compareBom(before, after).find((diff) => diff.before?.ref === "00A" && diff.after?.ref === "00A");
  assert.ok(moved);
  assert.equal(moved.kind, "changed");
  assert.equal(moved.primaryType, "positionChanged");
  assert.deepEqual(moved.processChange, { before: "SMT", after: "DIP" });
  assert.deepEqual(moved.structureChange, { before: "vb-t", after: "vb-d" });
  assert.deepEqual(moved.fields, ["架構變更", "製程別放置異常"]);
  assert.deepEqual(bomDiffDisplayFields(moved), ["架構變更（VB-T → VB-D）", "製程別放置異常（SMT → DIP）"]);
  assert.deepEqual(moved.newParts, []);
  assert.deepEqual(moved.deletedParts, []);
  assert.equal(moved.needsReview, true);
  assert.match(moved.matchReason, /同一料號跨製程架構移動/);

  const { exportCategories } = await import("../app/export-report.ts");
  assert.deepEqual(exportCategories(moved), ["製程別放置異常"]);
});

test("reports a 60 to VB-T structure change but ignores VB-T to VB-T", async () => {
  const header = ["項次", "料號", "數量", "插件位置", "製造商料號"];
  const bom = (branch) => parseCompanyBomMatrix([
    header,
    ["000", "69G14LM12A01", "1", "", "ROOT-MPN"],
    ["001", branch, "1", "", "BRANCH-MPN"],
    ["00A", "0500-04WB0ZY", "1", "U20", "FLASH-MPN"],
  ]);

  const moved = compareBom(bom("60-BOARD000001"), bom("VBG14LM12A01T"))
    .find((diff) => diff.before?.ref === "00A" && diff.after?.ref === "00A");
  assert.ok(moved);
  assert.equal(moved.kind, "changed");
  assert.equal(moved.primaryType, "positionChanged");
  assert.deepEqual(moved.structureChange, { before: "board60", after: "vb-t" });
  assert.equal(moved.processChange, undefined);
  assert.deepEqual(moved.fields, ["架構變更"]);
  assert.deepEqual(bomDiffDisplayFields(moved), ["架構變更（60 → VB-T）"]);
  assert.deepEqual(moved.newParts, []);
  assert.deepEqual(moved.deletedParts, []);
  assert.equal(moved.needsReview, true);
  assert.match(moved.matchReason, /同一料號跨架構移動（60 → VB-T）/);

  const unchanged = compareBom(bom("VBG14LM12A01T"), bom("VB-G14LM12A01T"))
    .find((diff) => diff.before?.ref === "00A" && diff.after?.ref === "00A");
  assert.ok(unchanged);
  assert.equal(unchanged.kind, "same");
  assert.equal(unchanged.structureChange, undefined);

  const { buildBomReport, exportCategories } = await import("../app/export-report.ts");
  assert.deepEqual(exportCategories(moved), ["製程別放置異常"]);
  const workbook = buildBomReport([moved], "before.xlsx", "after.xlsx");
  assert.equal(workbook.getWorksheet("差異摘要").getColumn(1).values.includes("架構／製程變更（1）"), true);
  const english = buildBomReport([moved], "before.xlsx", "after.xlsx", {}, "en");
  assert.equal(english.getWorksheet("Difference Summary").getColumn(1).values.includes("Structure / Process Changes（1）"), true);
});

test("does not pair different parts merely because they moved between SMT and DIP", () => {
  const header = ["項次", "料號", "數量", "插件位置"];
  const bom = (branch, part) => parseCompanyBomMatrix([
    header,
    ["000", "69G14LM12A01", "1", ""],
    ["001", branch, "1", ""],
    ["00A", part, "1", "U20"],
  ]);
  const diffs = compareBom(
    bom("VBG14LM12A01T", "0500-04WB0ZY"),
    bom("VBG14LM12A01D", "0603-020R0VD"),
  );
  assert.ok(!diffs.some((diff) => diff.processChange));
  assert.ok(diffs.some((diff) => diff.primaryType === "componentRemoved" && diff.before?.ref === "00A"));
  assert.ok(diffs.some((diff) => diff.primaryType === "componentAdded" && diff.after?.ref === "00A"));
});

test("allows optional manufacturer columns to be absent", () => {
  const matrix = [
    ["插件位置", "料號", "項次", "數量"],
    ["U1", "PREFIX-123456789012", "001", "1.0/1"],
  ];
  const parsed = parseCompanyBomMatrix(matrix);

  assert.equal(findCompanyHeader(matrix), 0);
  assert.equal(parsed.length, 1);
  assert.equal(parsed[0].part, "123456789012");
  assert.equal(parsed[0].manufacturerPart, "");
  assert.equal(parsed[0].manufacturerName, "");
});

test("reports quantity mismatches, duplicate positions, and 12-character collisions", () => {
  const matrix = [
    ["項次", "料號", "數量", "插件位置"],
    ["001", "PREFIX-A-123456789012", "2", "U1"],
    ["", "PREFIX-B-123456789012", "2", ""],
    ["002", "UNIQUE-PART", "1", "U1"],
  ];
  const result = analyzeCompanyBomMatrix(matrix, 0);
  const codes = result.audit.issues.map((issue) => issue.code);
  assert.ok(codes.includes("quantity-mismatch"));
  assert.ok(codes.includes("duplicate-position"));
  assert.ok(codes.includes("part-collision"));
  assert.equal(result.audit.mappingLabels.part, "料號");
  assert.deepEqual(result.items[0].sourceRows, [2, 3]);
});

test("accepts the same canonical part in separate location groups for different customer TPN classifications", () => {
  const matrix = [
    ["項次", "主件料號", "組成用量", "插件位置", "製造廠商料號", "對應客戶料號"],
    ["0PA", "| |---1A30-04F40ZY", "8.0/1", "C34,C36,C56,C57,C134,C135,C138,C173", "CL10B105KO8VPNC", "TPN-A"],
    ["0PK", "| |---1A30-04U90ZY", "1.0/1", "C10", "EMK107B7105KAHT", "TPN-B"],
    ["", "| | |S-1A30-04F40ZY", "1.0/1", "", "CL10B105KO8VPNC", "TPN-B"],
  ];

  const result = analyzeCompanyBomMatrix(matrix, 0);
  assert.equal(result.items.length, 2);
  assert.deepEqual(result.items.map((item) => item.positions), [
    ["C34", "C36", "C56", "C57", "C134", "C135", "C138", "C173"],
    ["C10"],
  ]);
  assert.equal(result.audit.issues.some((issue) => issue.code === "part-collision"), false);
});

test("ignores main/substitute ordering but reports added parts and placements", () => {
  const header = Array(18).fill("");
  header[0] = "項次"; header[1] = "料號"; header[4] = "數量"; header[5] = "插件位置"; header[15] = "製造商料號";
  const row = (item, part, qty, positions, mpn, customerPart = "") => {
    const values = Array(18).fill("");
    values[0] = item; values[1] = part; values[4] = qty; values[5] = positions; values[15] = mpn; values[17] = customerPart;
    return values;
  };
  const before = parseCompanyBomMatrix([header, row("00K", "MAIN", "2.0/1", "U1,U2", "M1", "CUSTOMER-OLD"), row("", "ALT", "2.0/1", "", "M2")]);
  const reordered = parseCompanyBomMatrix([header, row("DIFFERENT-A", "ALT", "2.0/1", "U1,U2", "M2", "CUSTOMER-NEW"), row("", "MAIN", "2.0/1", "", "M1")]);
  assert.equal(compareBom(before, reordered)[0].kind, "same");
  assert.equal(compareBom(before, reordered)[0].primaryType, "same");

  const changed = parseCompanyBomMatrix([header, row("ANOTHER-A", "ALT", "3.0/1", "U1,U2,U3", "M2"), row("", "MAIN", "3.0/1", "", "M1"), row("", "ALT-2", "3.0/1", "", "M3")]);
  const diff = compareBom(before, changed)[0];
  assert.equal(diff.kind, "changed");
  assert.equal(diff.primaryType, "substituteAdded");
  assert.deepEqual(diff.addedParts.map((part) => part.part), ["ALT-2"]);
  assert.deepEqual(diff.addedPositions, ["U3"]);
  assert.match(diff.fields.join(" "), /新增替料/);
  assert.match(diff.fields.join(" "), /新增插件位置/);
  assert.doesNotMatch(diff.fields.join(" "), /數量差異/);
  assert.deepEqual(diff.categories, ["added", "changed"]);
  assert.equal(diff.after?.qty, 3);

  const entirelyAdded = compareBom([], changed)[0];
  assert.equal(entirelyAdded.primaryType, "componentAdded");
  assert.match(entirelyAdded.fields.join(" "), /新增料號/);
  assert.doesNotMatch(entirelyAdded.fields.join(" "), /新增替料/);

  const reverse = compareBom(changed, before)[0];
  assert.deepEqual(reverse.removedParts.map((part) => part.part), ["ALT-2"]);
  assert.deepEqual(reverse.removedPositions, ["U3"]);
  assert.equal(reverse.primaryType, "substituteRemoved");
  assert.match(reverse.fields.join(" "), /刪除替料/);
  assert.match(reverse.fields.join(" "), /移除插件位置/);
});

test("classifies substitute-only additions and deletions as added or removed", () => {
  const base = [{
    ref: "A",
    part: "MAIN",
    manufacturerPart: "M1",
    value: "",
    description: "",
    qty: 1,
    positions: ["U1"],
    alternatives: [{ part: "MAIN", manufacturerPart: "M1", description: "", spec: "" }],
  }];
  const withSubstitute = [{
    ...base[0],
    alternatives: [...base[0].alternatives, { part: "ALT", manufacturerPart: "M2", description: "", spec: "" }],
  }];

  const added = compareBom(base, withSubstitute)[0];
  assert.equal(added.kind, "added");
  assert.equal(added.primaryType, "substituteAdded");
  assert.deepEqual(added.categories, ["added"]);
  assert.deepEqual(added.fields, ["新增替料"]);
  assert.deepEqual(added.newParts.map((part) => part.part), ["ALT"]);

  const removed = compareBom(withSubstitute, base)[0];
  assert.equal(removed.kind, "removed");
  assert.equal(removed.primaryType, "substituteRemoved");
  assert.deepEqual(removed.categories, ["removed"]);
  assert.deepEqual(removed.fields, ["刪除替料"]);
  assert.deepEqual(removed.deletedParts.map((part) => part.part), ["ALT"]);
});

test("shows a same-position replacement in added, removed, and changed categories", () => {
  const item = (part, position) => ({
    ref: position,
    part,
    manufacturerPart: `${part}-MPN`,
    value: "",
    description: "",
    qty: 1,
    positions: [position],
    alternatives: [{ part, manufacturerPart: `${part}-MPN`, description: "", spec: "" }],
  });
  const diff = compareBom([item("PART-A", "U20")], [item("PART-B", "U20")])[0];

  assert.equal(diff.kind, "changed");
  assert.equal(diff.primaryType, "partReplaced");
  assert.deepEqual(diff.categories, ["added", "removed", "changed"]);
  assert.deepEqual(diff.replacementPositions, ["U20"]);
  assert.equal(diff.matchConfidence, "high");
  assert.equal(diff.needsReview, false);
  assert.deepEqual(diff.newParts.map((part) => part.part), ["PART-B"]);
  assert.deepEqual(diff.deletedParts.map((part) => part.part), ["PART-A"]);
  assert.match(diff.fields.join(" "), /更換料號/);
  assert.match(diff.fields.join(" "), /新增料號/);
  assert.match(diff.fields.join(" "), /刪除料號/);
});

test("marks ambiguous pairings for manual review", () => {
  const item = (ref) => ({ ref, part: "SAME", manufacturerPart: "M", value: "", description: "", qty: 1, positions: ["U1"], alternatives: [{ part: "SAME", manufacturerPart: "M", description: "", spec: "" }] });
  const diffs = compareBom([item("A"), item("B")], [item("C"), item("D")]);
  assert.ok(diffs.every((diff) => diff.needsReview));
  assert.ok(diffs.every((diff) => diff.matchConfidence === "low"));
});

test("classifies a same-quantity placement relocation as changed only", () => {
  const item = (positions) => ({
    ref: "018",
    part: "0603-02960YD",
    manufacturerPart: "SN74LVC1G32QDCKRQ1",
    value: "",
    description: "",
    qty: positions.length,
    positions,
    alternatives: [{ part: "0603-02960YD", manufacturerPart: "SN74LVC1G32QDCKRQ1", description: "", spec: "" }],
  });
  const diff = compareBom(
    [item(["U29", "U32", "U49", "U52"])],
    [item(["U29", "U32", "U49", "U53"])],
  )[0];

  assert.equal(diff.kind, "changed");
  assert.equal(diff.primaryType, "positionChanged");
  assert.deepEqual(diff.categories, ["changed"]);
  assert.deepEqual(diff.addedPositions, ["U53"]);
  assert.deepEqual(diff.removedPositions, ["U52"]);
  assert.deepEqual(diff.fields, ["新增插件位置", "移除插件位置"]);
});

test("sorts the all view into addition, removal, and change groups", () => {
  const item = (part, positions = ["U1"], qty = positions.length, mpn = `${part}-MPN`, alternatives = null) => ({
    ref: positions[0] ?? part,
    part,
    manufacturerPart: mpn,
    value: "",
    description: "",
    qty,
    positions,
    alternatives: alternatives ?? [{ part, manufacturerPart: mpn, description: "", spec: "" }],
  });
  const base = item("MAIN");
  const withSubstitute = item("MAIN", ["U1"], 1, "MAIN-MPN", [
    ...base.alternatives,
    { part: "ALT", manufacturerPart: "ALT-MPN", description: "", spec: "" },
  ]);
  const diffs = [
    compareBom([], [item("NEW")])[0],
    compareBom([base], [withSubstitute])[0],
    compareBom([item("OLD")], [])[0],
    compareBom([withSubstitute], [base])[0],
    compareBom([item("MOVE", ["U2", "U3"])], [item("MOVE", ["U2", "U4"])])[0],
    compareBom([item("A", ["U5"])], [item("B", ["U5"])])[0],
    compareBom([item("QTY", [], 1)], [item("QTY", [], 2)])[0],
  ];

  assert.deepEqual(sortBomDiffsForAll(diffs).map((diff) => diff.primaryType), [
    "componentAdded",
    "substituteAdded",
    "componentRemoved",
    "substituteRemoved",
    "positionChanged",
    "positionChanged",
    "partReplaced",
  ]);
});

test("does not call a moved part globally new or globally deleted", async () => {
  const item = (ref, part, position) => ({
    ref,
    part,
    manufacturerPart: `${part}-MPN`,
    value: "",
    description: "",
    qty: 1,
    positions: [position],
    alternatives: [{ part, manufacturerPart: `${part}-MPN`, description: "", spec: "" }],
  });
  const before = [item("A", "PART-A", "U20"), item("B", "PART-B", "U30")];
  const after = [item("A", "PART-B", "U20"), item("B", "PART-A", "U30")];
  const diffs = compareBom(before, after);

  assert.equal(diffs.length, 2);
  assert.deepEqual(diffs.flatMap((diff) => diff.newParts), []);
  assert.deepEqual(diffs.flatMap((diff) => diff.deletedParts), []);
  assert.ok(diffs.every((diff) => diff.categories.includes("changed")));
  assert.ok(diffs.every((diff) => diff.primaryType === "partReplaced"));
  const { exportCategories } = await import("../app/export-report.ts");
  assert.ok(diffs.every((diff) => exportCategories(diff).length === 0));
});

test("ignores P- and K-column-only changes", () => {
  const header = Array(18).fill("");
  header[0] = "項次"; header[1] = "料號"; header[4] = "數量"; header[5] = "插件位置"; header[10] = "製造商名稱"; header[15] = "製造商料號";
  const row = (item, part, mpn, manufacturerName) => {
    const values = Array(18).fill("");
    values[0] = item; values[1] = part; values[4] = "1.0/1"; values[5] = "U20"; values[10] = manufacturerName; values[15] = mpn;
    return values;
  };
  const before = parseCompanyBomMatrix([header, row("OLD-A", "PREFIX-123456789012", "MPN-OLD", "OLD MAKER")]);
  const after = parseCompanyBomMatrix([header, row("NEW-A", "OTHER-123456789012", "MPN-NEW", "NEW MAKER")]);
  const diff = compareBom(before, after)[0];
  assert.equal(diff.kind, "same");
  assert.equal(diff.primaryType, "same");
  assert.deepEqual(diff.categories, []);
  assert.equal(diff.addedParts.length, 0);
  assert.equal(diff.removedParts.length, 0);
  assert.deepEqual(diff.fields, []);
});

test("locks the local app to same-origin resources and disables sensitive permissions", async () => {
  const response = await render();
  const csp = response.headers.get("content-security-policy") ?? "";
  assert.match(csp, /connect-src 'self'/);
  assert.match(csp, /worker-src 'self' blob:/);
  assert.match(csp, /frame-ancestors 'none'/);
  assert.equal(response.headers.get("referrer-policy"), "no-referrer");
  assert.match(response.headers.get("permissions-policy") ?? "", /camera=\(\)/);

  const page = await readFile(new URL("../app/page.tsx", import.meta.url), "utf8");
  assert.doesNotMatch(page, /\bfetch\s*\(|XMLHttpRequest|sendBeacon|WebSocket/);
  assert.match(page, /PdfSchematicViewer/);
  assert.match(page, /在線路圖定位/);
});

test("falls back to a local PDF compatibility mode with actionable errors", async () => {
  const [viewer, css] = await Promise.all([
    readFile(new URL("../app/pdf-schematic-viewer.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/globals.css", import.meta.url), "utf8"),
  ]);
  assert.match(viewer, /pdfjs-dist\/legacy\/build\/pdf\.mjs/);
  assert.match(viewer, /legacy\/build\/pdf\.worker\.min\.mjs\?url/);
  assert.match(viewer, /PDF_LOAD_TIMEOUT_MS/);
  assert.match(viewer, /瀏覽器版本較舊，已自動切換 PDF 相容模式/);
  assert.match(viewer, /PDF Worker 被瀏覽器或公司資安政策阻擋/);
  assert.match(viewer, /PDF 已加密或需要密碼/);
  assert.match(viewer, /PDF 檔案損壞或格式不完整/);
  assert.match(viewer, /PDF 無法載入/);
  assert.match(css, /\.pdf-status\.compatibility/);
  assert.match(css, /\.pdf-error-state/);
});

test("supports collapsing the desktop sidebar without changing the mobile menu", async () => {
  const [page, css] = await Promise.all([
    readFile(new URL("../app/page.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/globals.css", import.meta.url), "utf8"),
  ]);
  assert.match(page, /sidebarCollapsed/);
  assert.match(page, /收折側邊選單/);
  assert.match(page, /展開側邊選單/);
  assert.match(page, /aria-expanded/);
  assert.match(css, /\.sidebar\.collapsed/);
  assert.match(css, /\.workspace\.sidebar-collapsed/);
  assert.match(css, /@media \(min-width: 761px\)/);
  assert.match(css, /@media \(max-width: 760px\)/);
});

test("includes offline launchers and GitHub-built Windows packages", async () => {
  const [macLauncher, windowsLauncher, portableBuilder, iconVerifier, offlineServer, installer, workflow, packageJson] = await Promise.all([
    readFile(new URL("../scripts/start-offline.command", import.meta.url), "utf8"),
    readFile(new URL("../scripts/start-offline.bat", import.meta.url), "utf8"),
    readFile(new URL("../scripts/build-windows-portable.ps1", import.meta.url), "utf8"),
    readFile(new URL("../scripts/verify-windows-icon.ps1", import.meta.url), "utf8"),
    readFile(new URL("../scripts/offline-server.mjs", import.meta.url), "utf8"),
    readFile(new URL("../scripts/BOMLens.iss", import.meta.url), "utf8"),
    readFile(new URL("../.github/workflows/build-windows.yml", import.meta.url), "utf8"),
    readFile(new URL("../package.json", import.meta.url), "utf8"),
  ]);
  assert.match(macLauncher, /127\.0\.0\.1:3784/);
  assert.match(windowsLauncher, /127\.0\.0\.1:3784/);
  assert.match(portableBuilder, /BOMLens\.exe/);
  assert.match(portableBuilder, /runtime\\node\.exe/);
  assert.match(portableBuilder, /Framework64\\v4\.0\.30319\\csc\.exe/);
  assert.match(portableBuilder, /\/platform:x64/);
  assert.match(portableBuilder, /BOMLens\.ico/);
  assert.match(portableBuilder, /\/win32icon:/);
  assert.match(portableBuilder, /Icon\.ExtractAssociatedIcon/);
  assert.match(portableBuilder, /verify-windows-icon\.ps1/);
  assert.match(iconVerifier, /Icon\]::ExtractAssociatedIcon/);
  assert.match(iconVerifier, /differenceRatio/);
  assert.match(iconVerifier, /Windows icon 驗證通過/);
  assert.match(portableBuilder, /Compress-Archive/);
  assert.match(portableBuilder, /BOMLens-Windows-x64-Portable\.zip/);
  assert.match(portableBuilder, /offline-server\.mjs/);
  assert.match(offlineServer, /noCompression: true/);
  assert.match(offlineServer, /host: "127\.0\.0\.1"/);
  assert.match(offlineServer, /createServer/);
  assert.match(offlineServer, /dist", "client/);
  assert.match(offlineServer, /text\/css; charset=utf-8/);
  assert.match(installer, /PrivilegesRequired=lowest/);
  assert.match(installer, /\{autodesktop\}/);
  assert.match(installer, /SetupIconFile=.*BOMLens\.ico/);
  assert.match(installer, /IconFilename: "\{app\}\\\{#MyAppExeName\}"/);
  assert.match(installer, /BOMLens-Setup-x64/);
  assert.match(workflow, /runs-on: windows-latest/);
  assert.match(workflow, /actions\/setup-node@v4/);
  assert.match(workflow, /actions\/upload-artifact@v4/);
  assert.match(workflow, /Verify portable runtime/);
  assert.match(workflow, /Stylesheet 內容不完整/);
  assert.match(workflow, /\\\.app-shell/);
  assert.match(workflow, /Verify setup installer/);
  assert.equal((workflow.match(/verify-windows-icon\.ps1/g) ?? []).length, 3);
  assert.match(packageJson, /package:windows/);
  assert.match(windowsLauncher, /Invoke-WebRequest/);
  assert.match(windowsLauncher, /call npm run offline:serve/);
  assert.match(packageJson, /vinext start --hostname 127\.0\.0\.1 --port 3784/);
});

test("detects customer BOM defaults and maps reordered location sets with exact MPNs", () => {
  const header = Array(17).fill("");
  header[1] = "PartNumber";
  header[6] = "MfgPNos";
  header[16] = "ReferenceDesignator";
  assert.deepEqual(detectCustomerColumns(header), { customerPartNumber: 1, manufacturerParts: 6, positions: 16 });
  const rows = [header, ["", "CUST-001", "", "", "", "", "MPN-A|MPN-B", "", "", "", "", "", "", "", "", "", "C2|C1"]];
  const records = parseCustomerBomMatrix(rows, 0, detectCustomerColumns(header));
  const company = [{
    ref: "001", part: "PART-A", manufacturerPart: "MPN-A", manufacturerName: "", value: "", description: "", qty: 2,
    positions: ["C1", "C2"],
    rdCustomerPartNumbers: ["CUST-001"],
    alternatives: [
      { part: "PART-A", manufacturerPart: "mpn-a", description: "", spec: "", rdCustomerPartNumbers: ["CUST-001"] },
      { part: "PART-B", manufacturerPart: "MPN-B", description: "", spec: "", rdCustomerPartNumbers: ["CUST-001"] },
    ],
  }];
  const result = mapCustomerBom(company, records);
  assert.equal(result.counts.matched, 2);
  assert.equal(result.items[0].customerPartNumber, "CUST-001");
  assert.equal(result.items[0].customerMappingStatus, "matched");
});

test("splits virtual customer TPN MPN lists separated by double tildes", () => {
  const header = ["PartNumber", "ReferenceDesignator", "MfgPNos"];
  const records = parseCustomerBomMatrix([
    header,
    [
      "1125327-00-A",
      "C167",
      "HMJ107BB7104KAHT~~MCJCH168BB7104KTPA01~~GCJ188R72A104KA01D",
    ],
  ], 0, detectCustomerColumns(header));
  assert.deepEqual(records[0].manufacturerParts, [
    "HMJ107BB7104KAHT",
    "MCJCH168BB7104KTPA01",
    "GCJ188R72A104KA01D",
  ]);

  const company = [{
    ref: "001",
    part: "1A30-04CT0ZY",
    manufacturerPart: "HMJ107BB7104KAHT",
    manufacturerName: "",
    value: "",
    description: "",
    qty: 1,
    positions: ["C167"],
    alternatives: [
      { part: "1A30-04CT0ZY", manufacturerPart: "HMJ107BB7104KAHT", description: "", spec: "", rdCustomerPartNumbers: ["1125327-00-A"] },
      { part: "1A30-04CT1ZY", manufacturerPart: "MCJCH168BB7104KTPA01", description: "", spec: "", rdCustomerPartNumbers: ["1125327-00-A"] },
      { part: "1A30-04CC0ZY", manufacturerPart: "GCJ188R72A104KA01D", description: "", spec: "", rdCustomerPartNumbers: ["1125327-00-A"] },
    ],
  }];
  const result = mapCustomerBom(company, records);
  assert.equal(result.rows[0].status, "matched");
  assert.equal(result.counts.matched, 3);
  assert.deepEqual(
    result.items[0].alternatives.map((alternative) => alternative.customerPartNumbers),
    [["1125327-00-A"], ["1125327-00-A"], ["1125327-00-A"]],
  );
});

test("matches an exact virtual TPN MPN after removing invisible Excel formatting", () => {
  const company = parseCompanyBomMatrix([
    ["項次", "主件料號", "組成用量", "插件位置", "製造廠商", "製造廠商料號", "對應客戶料號", "廠商/料號及型號"],
    ["000", "69-2G126519P03", 1, "", "", "", "", ""],
    ["00A", "VB-2G126519P03T", 1, "", "", "", "", ""],
    ["001", "0704-070M0VD", 1, "Q32", "TOSHIBA", "", "2303598-00-A,1062138-00-C", "TOSHIBA/SSM3K318R_LXGF,TOSHIBA/SSM3K318R_LXGF"],
  ]);
  const records = parseCustomerBomMatrix([
    ["PartNumber", "ReferenceDesignator", "MfgPNos"],
    ["1062138-00-C", "Q32", "UNUSED-MPN~~SSM3K318R＿LXGF\u200B"],
  ], 0, detectCustomerColumns(["PartNumber", "ReferenceDesignator", "MfgPNos"]));

  const result = mapCustomerBom(company, records);
  const row = result.alternativeRows.find((candidate) => candidate.part === "0704-070M0VD");
  assert.equal(row?.status, "matched");
  assert.deepEqual(row?.customerPartNumbers, ["1062138-00-C"]);
  assert.match(row?.customerAssociationEvidence[0] ?? "", /1062138-00-C → TOSHIBA\/SSM3K318R_LXGF/);
  assert.equal(result.rows[0].status, "mpn-unmatched");
});

test("detects TPN as the customer BOM part number header", () => {
  assert.deepEqual(
    detectCustomerColumns(["Level", "TPN", "MfgPNos", "ReferenceDesignator"]),
    { customerPartNumber: 1, manufacturerParts: 2, positions: 3 },
  );
});

test("requires every company MPN to match exactly including suffixes and symbols", () => {
  const company = [{
    ref: "001", part: "PART-A", manufacturerPart: "ABC-100-A", manufacturerName: "", value: "", description: "", qty: 1,
    positions: ["U1"],
    rdCustomerPartNumbers: ["CUST"],
    alternatives: [
      { part: "PART-A", manufacturerPart: "ABC-100-A", description: "", spec: "", rdCustomerPartNumbers: ["CUST"] },
      { part: "PART-B", manufacturerPart: "XYZ/200", description: "", spec: "", rdCustomerPartNumbers: ["CUST"] },
    ],
  }];
  const record = (mpns) => [{ customerPartNumber: "CUST", manufacturerParts: mpns, positions: ["U1"], sourceRow: 2 }];
  assert.equal(mapCustomerBom(company, record(["abc-100-a", "XYZ/200"])).items[0].customerMappingStatus, "matched");
  assert.equal(mapCustomerBom(company, record(["ABC-100", "XYZ/200"])).items[0].customerMappingStatus, "mpn-unmatched");
  assert.equal(mapCustomerBom(company, record(["ABC100-A", "XYZ/200"])).items[0].customerMappingStatus, "mpn-unmatched");
  assert.equal(mapCustomerBom(company, record(["ABC-100-A", "XYZ-200"])).items[0].customerMappingStatus, "mpn-unmatched");
});

test("flags RD maintenance when Location and MPN match but the BOM R column is missing or inconsistent", () => {
  const companyItem = (rdCustomerPartNumbers) => ({
    ref: "001", part: "PART-A", manufacturerPart: "MPN-A", manufacturerName: "", value: "", description: "", qty: 1,
    positions: ["U1"], rdCustomerPartNumbers,
    alternatives: [{ part: "PART-A", manufacturerPart: "MPN-A", description: "", spec: "", rdCustomerPartNumbers }],
  });
  const records = [{ customerPartNumber: "CUST-001", manufacturerParts: ["MPN-A"], positions: ["U1"], sourceRow: 2 }];
  const missing = mapCustomerBom([companyItem([])], records);
  assert.equal(missing.items[0].customerMappingStatus, "rd-maintenance-missing");
  assert.equal(missing.rows[0].status, "rd-maintenance-missing");
  assert.match(missing.rows[0].reason, /請 RD 維護/);
  const mismatch = mapCustomerBom([companyItem(["CUST-999"])], records);
  assert.equal(mismatch.items[0].customerMappingStatus, "rd-maintenance-mismatch");
  const matched = mapCustomerBom([companyItem(["CUST-999", "CUST-001"])], records);
  assert.equal(matched.items[0].customerMappingStatus, "matched");
});

test("parses ordered BOM R and S associations and validates duplicate TPN candidates by exact MPN", () => {
  const header = ["項次", "主件料號", "組成用量", "插件位置", "製造廠商", "製造廠商料號", "對應客戶料號", "廠商/料號及型態"];
  assert.equal(detectCompanyColumns(header).customerPartAssociations, 7);
  const company = parseCompanyBomMatrix([
    header,
    ["001", "0704-06HL0ZY", 1, "Q33,Q13", "TOSHIBA", "SSM3K361R_LXHF", "2156321-00-A,1982053-00-A", "TOSHIBA/SSM3K361R_LXHF,TOSHIBA/OTHER-MPN"],
    ["002", "0704-OTHER01", 1, "Q99", "TOSHIBA", "SSM3K361R_LXHF", "2156321-00-A", "TOSHIBA/SSM3K361R_LXHF"],
  ]);
  assert.deepEqual(company[0].alternatives[0].rdCustomerPartAssociations, [
    { customerPartNumber: "2156321-00-A", manufacturerPart: "SSM3K361R_LXHF", manufacturerName: "TOSHIBA", sourceRow: 2 },
    { customerPartNumber: "1982053-00-A", manufacturerPart: "OTHER-MPN", manufacturerName: "TOSHIBA", sourceRow: 2 },
  ]);
  const result = mapCustomerBom(company, [
    { customerPartNumber: "2156321-00-A", manufacturerParts: ["SSM3K361R_LXHF"], positions: ["Q13", "Q33"], sourceRow: 2 },
  ]);
  assert.equal(result.alternativeRows[0].status, "matched");
  assert.match(result.alternativeRows[0].customerAssociationEvidence[0], /2156321-00-A → TOSHIBA\/SSM3K361R_LXHF/);
});

test("flags S-column MPN mismatches and malformed R/S ordering without fuzzy matching", () => {
  const header = ["項次", "主件料號", "組成用量", "插件位置", "製造廠商料號", "對應客戶料號", "廠商／料號及型態"];
  const mismatch = parseCompanyBomMatrix([
    header,
    ["001", "PART-A", 1, "U1", "ABC-100-A", "TPN-1", "MAKER/ABC100-A"],
  ]);
  const mismatchResult = mapCustomerBom(mismatch, [
    { customerPartNumber: "TPN-1", manufacturerParts: ["ABC-100-A"], positions: ["U1"], sourceRow: 2 },
  ]);
  assert.equal(mismatchResult.alternativeRows[0].status, "tpn-association-mismatch");
  assert.match(mismatchResult.alternativeRows[0].customerAssociationEvidence[0], /MAKER\/ABC100-A/);

  const malformedMatrix = [
    header,
    ["001", "PART-A", 1, "U1", "ABC-100-A", "TPN-1,TPN-2", "MAKER/ABC-100-A"],
  ];
  const malformed = parseCompanyBomMatrix(malformedMatrix);
  const malformedResult = mapCustomerBom(malformed, [
    { customerPartNumber: "TPN-1", manufacturerParts: ["ABC-100-A"], positions: ["U1"], sourceRow: 2 },
  ]);
  assert.equal(malformedResult.alternativeRows[0].status, "tpn-association-invalid");
  assert.match(malformedResult.alternativeRows[0].reason, /R／S 欄順序資料不完整/);
  assert.match(malformedResult.alternativeRows[0].customerAssociationEvidence[0], /Excel 第 2 列/);
  const audit = analyzeCompanyBomMatrix(malformedMatrix, 0).audit;
  assert.ok(audit.issues.some((issue) => issue.code === "customer-association-invalid"));
});

test("accepts an exact duplicate TPN association even when another R/S candidate is malformed", () => {
  const header = ["項次", "主件料號", "組成用量", "插件位置", "製造廠商料號", "對應客戶料號", "廠商料號及型態"];
  const company = parseCompanyBomMatrix([
    header,
    ["001", "PART-A", 1, "U1", "MPN-A", "TPN-1,TPN-2", "MAKER/MPN-A"],
    ["002", "PART-B", 1, "U2", "MPN-A", "TPN-1", "OTHER/MPN-A"],
  ]);
  const result = mapCustomerBom(company, [
    { customerPartNumber: "TPN-1", manufacturerParts: ["MPN-A"], positions: ["U1"], sourceRow: 2 },
  ]);
  assert.equal(result.alternativeRows[0].status, "matched");
  assert.match(result.alternativeRows[0].customerAssociationEvidence[0], /Excel 第 3 列/);
});

test("maps every main and substitute part independently and keeps multiple customer TPNs", () => {
  const company = [{
    ref: "001", part: "PART-A", manufacturerPart: "MPN-A", manufacturerName: "", value: "", description: "", qty: 1,
    positions: ["U1"],
    alternatives: [
      { part: "PART-A", manufacturerPart: "MPN-A", description: "", spec: "", rdCustomerPartNumbers: ["TPN-1", "TPN-2"] },
      { part: "PART-B", manufacturerPart: "MPN-B", description: "", spec: "", rdCustomerPartNumbers: [] },
    ],
  }];
  const records = [
    { customerPartNumber: "TPN-1", manufacturerParts: ["MPN-A", "MPN-B"], positions: ["U1"], sourceRow: 2 },
    { customerPartNumber: "TPN-2", manufacturerParts: ["MPN-A"], positions: ["U1"], sourceRow: 3 },
  ];
  const result = mapCustomerBom(company, records);
  assert.deepEqual(result.items[0].alternatives[0].customerPartNumbers, ["TPN-1", "TPN-2"]);
  assert.equal(result.items[0].alternatives[0].customerMappingStatus, "matched");
  assert.deepEqual(result.items[0].alternatives[1].customerPartNumbers, ["TPN-1"]);
  assert.equal(result.items[0].alternatives[1].customerMappingStatus, "rd-maintenance-missing");
  assert.equal(result.alternativeRows[1].role, "替料");
  assert.match(result.alternativeRows[1].reason, /RD 維護/);
});

test("uses Location to select the correct TPN when the same MPN appears more than once", () => {
  const item = (ref, position, tpn) => ({
    ref, part: "SAME-PART", manufacturerPart: "SAME-MPN", manufacturerName: "", value: "", description: "", qty: 1,
    positions: [position],
    alternatives: [{ part: "SAME-PART", manufacturerPart: "SAME-MPN", description: "", spec: "", rdCustomerPartNumbers: [tpn] }],
  });
  const result = mapCustomerBom(
    [item("001", "U1", "TPN-U1"), item("002", "U2", "TPN-U2")],
    [
      { customerPartNumber: "TPN-U2", manufacturerParts: ["SAME-MPN"], positions: ["U2"], sourceRow: 2 },
      { customerPartNumber: "TPN-U1", manufacturerParts: ["SAME-MPN"], positions: ["U1"], sourceRow: 3 },
    ],
  );
  assert.deepEqual(result.items[0].alternatives[0].customerPartNumbers, ["TPN-U1"]);
  assert.deepEqual(result.items[1].alternatives[0].customerPartNumbers, ["TPN-U2"]);
  assert.equal(result.counts.matched, 2);
});

test("reports BOM R column changes as TPN differences in exports", async () => {
  const item = (rdCustomerPartNumbers) => ({
    ref: "001", part: "PART-A", manufacturerPart: "MPN-A", manufacturerName: "", value: "", description: "", qty: 1,
    positions: ["U1"], rdCustomerPartNumbers,
    alternatives: [{ part: "PART-A", manufacturerPart: "MPN-A", description: "", spec: "", rdCustomerPartNumbers }],
  });
  const diff = compareBom([item(["CUST-OLD"])], [item(["CUST-NEW"])])[0];
  assert.equal(diff.kind, "changed");
  assert.deepEqual(diff.fields, ["客戶料號差異"]);
  const { exportCategories } = await import("../app/export-report.ts");
  assert.deepEqual(exportCategories(diff), ["TPN 差異"]);
});

test("parses pick and place columns and counts unique SMT/DIP Top/Bottom placements", () => {
  const header = ["Designator", "Layer"];
  assert.deepEqual(detectPlacementColumns(header), { designator: 0, layer: 1 });
  const records = parsePlacementMatrix([
    header,
    ["U1", "TopLayer"],
    ["U1", "TopLayer"],
    ["U2", "BottomLayer"],
    ["J1", "Top"],
    ["FID1", "BottomLayer"],
    ["BAD1", "Unknown"],
  ], 0, detectPlacementColumns(header));
  const item = (part, positions, structureKind) => ({
    ref: part, part, manufacturerPart: `${part}-MPN`, manufacturerName: "", value: "", description: "", qty: positions.length, positions, structureKind,
    alternatives: [{ part, manufacturerPart: `${part}-MPN`, description: "", spec: "" }],
  });
  const summary = calculateMva([
    item("SMT-PART", ["U1", "U2"], "vb-t"),
    item("DIP-PART", ["J1"], "vb-d"),
  ], records);
  assert.equal(summary.smtTop, 1);
  assert.equal(summary.smtBottom, 1);
  assert.equal(summary.dipTop, 1);
  assert.equal(summary.dipBottom, 0);
  assert.equal(summary.included.length, 3);
  assert.equal(summary.excluded.length, 2);
  assert.ok(summary.excluded.some((row) => row.designator === "FID1"));
  assert.ok(summary.excluded.some((row) => row.designator === "BAD1"));
});

test("exports customer mapping statuses and MVA details to Excel and HTML", async () => {
  const company = [{
    ref: "001", part: "PART-A", manufacturerPart: "MPN-A", manufacturerName: "MAKER", value: "", description: "", qty: 1,
    positions: ["U1"], structureKind: "vb-t",
    rdCustomerPartNumbers: ["CUST-001"],
    alternatives: [{ part: "PART-A", manufacturerPart: "MPN-A", manufacturerName: "MAKER", description: "", spec: "" }],
  }];
  const customer = mapCustomerBom(company, [{ customerPartNumber: "CUST-001", manufacturerParts: ["MPN-A"], positions: ["U1"], sourceRow: 2 }]);
  const mva = calculateMva(company, [{ designator: "U1", side: "Top", rawLayer: "TopLayer", sourceRow: 14 }]);
  const diffs = compareBom([], customer.items);
  const { buildBomReport } = await import("../app/export-report.ts");
  const { buildBomHtmlReport } = await import("../app/export-html-report.ts");
  const context = { customerAfter: customer, mvaAfter: mva };
  const workbook = buildBomReport(diffs, "before.xlsx", "after.xlsx", context);
  assert.ok(workbook.getWorksheet("客戶 BOM TPN 對應"));
  assert.ok(workbook.getWorksheet("MVA 明細"));
  assert.equal(workbook.getWorksheet("客戶 BOM TPN 對應").getCell("G3").value, "CUST-001");
  assert.equal(workbook.getWorksheet("客戶 BOM TPN 對應").getCell("I2").value, "BOM S欄關聯");
  assert.equal(workbook.getWorksheet("MVA 明細").getCell("B5").value, 1);
  const dataHeaders = workbook.getWorksheet("差異資料").getRow(5).values;
  assert.ok(dataHeaders.includes("新版 BOM R欄 TPN"));
  assert.ok(dataHeaders.includes("新版客戶 BOM TPN 驗證"));
  const englishWorkbook = buildBomReport(diffs, "before.xlsx", "after.xlsx", context, "en");
  assert.ok(englishWorkbook.getWorksheet("Customer TPN Mapping"));
  assert.ok(englishWorkbook.getWorksheet("MVA Details"));
  assert.equal(englishWorkbook.getWorksheet("Customer TPN Mapping").getCell("I2").value, "BOM S-column Association");
  assert.equal(englishWorkbook.getWorksheet("MVA Details").getCell("A3").value, "Version");
  const html = buildBomHtmlReport(diffs, "before.xlsx", "after.xlsx", context);
  assert.match(html, /客戶 BOM TPN 對應/);
  assert.match(html, /BOM S欄關聯/);
  assert.match(html, /CUST-001/);
  assert.match(html, /MVA 製程顆數/);
  assert.match(html, /SMT Top/);
});

test("shows a TPN for every summary part and uses mapped customer TPNs after customer BOM import", async () => {
  const company = [{
    ref: "001",
    part: "MAIN",
    manufacturerPart: "MAIN-MPN",
    manufacturerName: "MAKER-A",
    value: "",
    description: "",
    qty: 1,
    positions: ["U1"],
    alternatives: [
      { part: "MAIN", manufacturerPart: "MAIN-MPN", manufacturerName: "MAKER-A", description: "", spec: "", rdCustomerPartNumbers: ["RD-MAIN"] },
      { part: "ALT", manufacturerPart: "ALT-MPN", manufacturerName: "MAKER-B", description: "", spec: "", rdCustomerPartNumbers: ["RD-ALT"] },
    ],
  }];
  const customer = mapCustomerBom(company, [
    { customerPartNumber: "CUSTOMER-MAIN", manufacturerParts: ["MAIN-MPN"], positions: ["U1"], sourceRow: 2 },
    { customerPartNumber: "CUSTOMER-ALT", manufacturerParts: ["ALT-MPN"], positions: ["U1"], sourceRow: 3 },
  ]);
  const { buildBomReport } = await import("../app/export-report.ts");

  const mappedWorkbook = buildBomReport(compareBom([], customer.items), "before.xlsx", "after.xlsx", { customerAfter: customer });
  const mappedSummary = mappedWorkbook.getWorksheet("差異摘要");
  assert.equal(mappedSummary.getCell("N10").value, "MAIN");
  assert.equal(mappedSummary.getCell("Q10").value, "CUSTOMER-MAIN");
  assert.equal(mappedSummary.getCell("N11").value, "ALT");
  assert.equal(mappedSummary.getCell("Q11").value, "CUSTOMER-ALT");

  const originalWorkbook = buildBomReport(compareBom([], company), "before.xlsx", "after.xlsx");
  const originalSummary = originalWorkbook.getWorksheet("差異摘要");
  assert.equal(originalSummary.getCell("Q10").value, "RD-MAIN");
  assert.equal(originalSummary.getCell("Q11").value, "RD-ALT");
});
