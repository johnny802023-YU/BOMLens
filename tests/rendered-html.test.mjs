import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { analyzeCompanyBomMatrix, canonicalPartNumber, compareBom, detectCompanyColumns, findCompanyHeader, parseCompanyBomMatrix, parseQuantity, sortBomDiffsForAll } from "../app/bom-logic.ts";
import { buildPageReferenceIndex, centeredPdfHitScroll, centeredRenderedHitScroll, findReferenceHits, lookupReferenceHits, normalizeReference } from "../app/pdf-search.ts";
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
  assert.match(html, /依標題名稱自動定位欄位/);
  assert.match(html, /主要異動/);
  assert.match(html, /差異項目/);
  assert.match(html, /新版加入的主料或替料/);
  assert.match(html, /插件位置、數量或同位置換料/);
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
  assert.doesNotMatch(html, /codex-preview|Your site is taking shape/);
});

test("ships real BOM parsing, comparison, and export behavior", async () => {
  const [page, layout, styles, packageJson] = await Promise.all([
    readFile(new URL("../app/page.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/layout.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/globals.css", import.meta.url), "utf8"),
    readFile(new URL("../package.json", import.meta.url), "utf8"),
  ]);
  assert.match(page, /XLSX\.read/);
  assert.match(page, /analyzeCompanyBomMatrix/);
  assert.match(page, /book\.SheetNames\.map/);
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
  assert.match(page, /exportBomHtmlReport\(visible/);
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
  assert.equal(summarySheet.getCell("A7").value, "異動");
  assert.equal(summarySheet.getCell("G7").value, "舊版料號資訊");
  assert.equal(summarySheet.getCell("K7").value, "新版料號資訊");
  assert.equal(summarySheet.getCell("A8").value, "異動類型");
  assert.equal(summarySheet.getCell("G8").value, "舊版料號");
  assert.equal(summarySheet.getCell("H8").value, "舊版製造商料號");
  assert.equal(summarySheet.getCell("I8").value, "舊版製造商");
  assert.equal(summarySheet.getCell("K8").value, "新版料號");
  assert.equal(summarySheet.getCell("L8").value, "新版製造商料號");
  assert.equal(summarySheet.getCell("M8").value, "新版製造商");
  assert.equal(dataSheet.views[0].ySplit, 4);
  assert.equal(dataSheet.views[0].showGridLines, false);
  assert.deepEqual(dataSheet.model.tables.map((table) => table.name), ["BomDiffData"]);
  assert.deepEqual(exportCategories(replacementDiffs[0]), ["新版完全新料", "新版完全移除"]);
  assert.deepEqual(exportCategories(diffs[1]), ["數量差異"]);
  const buffer = await workbook.xlsx.writeBuffer();
  const parsed = XLSX.read(buffer, { type: "buffer" });

  assert.deepEqual(parsed.SheetNames, ["差異摘要", "差異資料"]);
  assert.equal(parsed.Sheets["差異摘要"].A1.v, "BOM 差異比較報告");
  assert.equal(parsed.Sheets["差異摘要"].A4.v, "新版完全新料\n1");
  assert.equal(parsed.Sheets["差異摘要"].A9.v, "新版完全新料（1）");
  assert.equal(parsed.Sheets["差異摘要"].A10.v, "新增料號");
  assert.equal(parsed.Sheets["差異摘要"].E10.v, "U20 換料");
  assert.equal(parsed.Sheets["差異摘要"].G10.v, "OLD-PART");
  assert.equal(parsed.Sheets["差異摘要"].H10.v, "OLD-PART-MPN");
  assert.equal(parsed.Sheets["差異摘要"].I10.v, "OLD-PART-MAKER");
  assert.equal(parsed.Sheets["差異摘要"].K10.v, "NEW-PART");
  assert.equal(parsed.Sheets["差異摘要"].L10.v, "NEW-PART-MPN");
  assert.equal(parsed.Sheets["差異摘要"].M10.v, "NEW-PART-MAKER");
  assert.equal(parsed.Sheets["差異摘要"].N10.v, "1 → 1");
  assert.equal(parsed.Sheets["差異摘要"].A13.v, "新版完全移除（1）");
  assert.equal(parsed.Sheets["差異摘要"].A17.v, "數量差異（1）");
  assert.equal(parsed.Sheets["差異摘要"].A19.v, "僅插件位置差異（0）");
  assert.equal(parsed.Sheets["差異資料"].A4.v, "分類");
  assert.equal(parsed.Sheets["差異資料"].B4.v, "主要異動");
  assert.equal(parsed.Sheets["差異資料"].F4.v, "舊版主件料號");
  assert.equal(parsed.Sheets["差異資料"].I4.v, "新版主件料號");
  assert.equal(parsed.Sheets["差異資料"].N4.v, "數量變化");
  assert.equal(parsed.Sheets["差異資料"].A5.v, "新版完全新料、新版完全移除");
  assert.equal(parsed.Sheets["差異資料"].E5.v, "U20 換料");
  assert.equal(parsed.Sheets["差異資料"].N5.v, "Same");
  assert.equal(parsed.Sheets["差異資料"].N6.v, "Increase");
  assert.ok(buffer.byteLength > 5_000);
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
    ["新版完全新料（1）", "新增替代（1）", "新版完全移除（1）", "刪除替代（1）", "數量差異（1）", "僅插件位置差異（1）"],
  );
  const substituteSectionRow = summarySheet.getColumn(1).values.findIndex((value) => value === "新增替代（1）");
  assert.equal(summarySheet.getCell(`G${substituteSectionRow + 1}`).value, "MAIN");
  assert.equal(summarySheet.getCell(`K${substituteSectionRow + 1}`).value, "MAIN");
  assert.equal(summarySheet.getCell(`K${substituteSectionRow + 2}`).value, "ALT");
  assert.equal(summarySheet.getCell(`L${substituteSectionRow + 2}`).value, "ALT-MPN");
  assert.deepEqual(
    [5, 6, 7, 8, 9, 10].map((row) => workbook.getWorksheet("差異資料").getCell(`A${row}`).value),
    ["新版完全新料", "新增替代", "新版完全移除", "刪除替代", "數量差異", "僅插件位置差異"],
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
  assert.equal(canonicalPartNumber("PREFIX-123456789012"), "123456789012");
  const parsed = parseCompanyBomMatrix(matrix);
  assert.equal(parsed.length, 1);
  assert.equal(parsed[0].ref, "00K");
  assert.equal(parsed[0].qty, 2);
  assert.deepEqual(parsed[0].positions, ["U20", "U31"]);
  assert.deepEqual(parsed[0].alternatives.map((part) => part.part), ["123456789012", "INTERNAL-ALT"]);
  assert.deepEqual(parsed[0].alternatives.map((part) => part.manufacturerName), ["Texas Instruments", "Nexperia"]);
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

test("shows explicit manual review and enlarged schematic controls", async () => {
  const page = await readFile(new URL("../app/page.tsx", import.meta.url), "utf8");
  const styles = await readFile(new URL("../app/globals.css", import.meta.url), "utf8");
  assert.match(page, /人工待審核/);
  assert.match(page, /reviewDiffs\.length/);
  assert.match(page, /setImpactFilter\(impactFilter === "review" \? "all" : "review"\)/);
  assert.match(page, /放大顯示/);
  assert.match(page, /退出放大/);
  assert.match(styles, /\.schematic-panel\.expanded/);
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

test("compares parent structure rows while pairing children only within the same branch", () => {
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
  assert.ok(diffs.some((diff) => diff.primaryType === "componentRemoved" && diff.before?.structureKind === "root69"));
  assert.ok(diffs.some((diff) => diff.primaryType === "componentAdded" && diff.after?.structureKind === "root69"));
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
    ["002", "PREFIX-B-123456789012", "1", "U1"],
  ];
  const result = analyzeCompanyBomMatrix(matrix, 0);
  const codes = result.audit.issues.map((issue) => issue.code);
  assert.ok(codes.includes("quantity-mismatch"));
  assert.ok(codes.includes("duplicate-position"));
  assert.ok(codes.includes("part-collision"));
  assert.equal(result.audit.mappingLabels.part, "料號");
  assert.deepEqual(result.items[0].sourceRows, [2]);
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
  const [macLauncher, windowsLauncher, portableBuilder, offlineServer, installer, workflow, packageJson] = await Promise.all([
    readFile(new URL("../scripts/start-offline.command", import.meta.url), "utf8"),
    readFile(new URL("../scripts/start-offline.bat", import.meta.url), "utf8"),
    readFile(new URL("../scripts/build-windows-portable.ps1", import.meta.url), "utf8"),
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
  assert.match(installer, /BOMLens-Setup-x64/);
  assert.match(workflow, /runs-on: windows-latest/);
  assert.match(workflow, /actions\/setup-node@v4/);
  assert.match(workflow, /actions\/upload-artifact@v4/);
  assert.match(workflow, /Verify portable runtime/);
  assert.match(workflow, /Stylesheet 內容不完整/);
  assert.match(workflow, /\\\.app-shell/);
  assert.match(workflow, /Verify setup installer/);
  assert.match(packageJson, /package:windows/);
  assert.match(windowsLauncher, /Invoke-WebRequest/);
  assert.match(windowsLauncher, /call npm run offline:serve/);
  assert.match(packageJson, /vinext start --hostname 127\.0\.0\.1 --port 3784/);
});
