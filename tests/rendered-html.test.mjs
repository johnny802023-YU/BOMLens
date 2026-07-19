import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { analyzeCompanyBomMatrix, canonicalPartNumber, compareBom, findCompanyHeader, parseCompanyBomMatrix, parseQuantity, sortBomDiffsForAll } from "../app/bom-logic.ts";
import { buildPageReferenceIndex, centeredPdfHitScroll, findReferenceHits, lookupReferenceHits, normalizeReference } from "../app/pdf-search.ts";
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
  assert.match(html, /料號新增／刪除/);
  assert.match(html, /插件位置差異/);
  assert.match(html, /依標題名稱自動定位欄位/);
  assert.match(html, /主要異動/);
  assert.match(html, /影響標籤/);
  assert.match(html, /同位置換料/);
  assert.match(html, /位置變更/);
  assert.match(html, /新版完全新料/);
  assert.match(html, /新版完全移除/);
  assert.match(html, /新增元件/);
  assert.match(html, /新增替料/);
  assert.match(html, /刪除替料/);
  assert.match(html, /移除元件/);
  assert.match(html, /新增插件位置/);
  assert.match(html, /移除插件位置/);
  assert.match(html, /更換料號/);
  assert.match(html, /新版完全新料/);
  assert.doesNotMatch(html, /需確認下單/);
  assert.match(html, /新版完全移除/);
  assert.doesNotMatch(html, /主替料變更|主體料變更/);
  assert.doesNotMatch(html, /來源 A/);
  assert.match(html, /線路圖比對/);
  assert.match(html, /離線隱私模式/);
  assert.match(html, /不會上傳、同步或儲存/);
  assert.match(html, /PCB_Main_v1\.3\.xlsx/);
  assert.doesNotMatch(html, /codex-preview|Your site is taking shape/);
});

test("ships real BOM parsing, comparison, and export behavior", async () => {
  const [page, layout, packageJson] = await Promise.all([
    readFile(new URL("../app/page.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/layout.tsx", import.meta.url), "utf8"),
    readFile(new URL("../package.json", import.meta.url), "utf8"),
  ]);
  assert.match(page, /XLSX\.read/);
  assert.match(page, /analyzeCompanyBomMatrix/);
  assert.match(page, /book\.SheetNames\.map/);
  assert.match(page, /待人工確認/);
  assert.match(page, /exportBomReport/);
  assert.match(page, /getImageData/);
  assert.match(page, /application\/pdf/);
  assert.match(layout, /lang="zh-Hant"/);
  assert.match(packageJson, /"xlsx"/);
  assert.match(packageJson, /"exceljs"/);
  assert.match(packageJson, /"offline:serve"/);
  assert.doesNotMatch(packageJson, /react-loading-skeleton/);
});

test("builds a formatted five-sheet Excel report with audit and review tabs", async () => {
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
  const diffs = compareBom([item("OLD-PART", "U20")], [item("NEW-PART", "U20")]);
  const { buildBomReport } = await import("../app/export-report.ts");
  const workbook = buildBomReport(diffs, "before.xlsx", "after.xlsx");
  const detailSheet = workbook.getWorksheet("差異明細");
  assert.equal(detailSheet.views[0].state, "frozen");
  assert.equal(detailSheet.views[0].ySplit, 4);
  assert.deepEqual(detailSheet.autoFilter, { from: "A4", to: "X4" });
  assert.equal(detailSheet.getCell("A4").fill.fgColor.argb, "2F6BCE");
  assert.ok(detailSheet.getColumn(2).width >= 20);
  const buffer = await workbook.xlsx.writeBuffer();
  const parsed = XLSX.read(buffer, { type: "buffer" });

  assert.deepEqual(parsed.SheetNames, ["差異摘要", "差異明細", "料號生命週期", "待人工確認", "匯入稽核"]);
  assert.equal(parsed.Sheets["差異摘要"].A1.v, "BOM 版本差異報告");
  assert.equal(parsed.Sheets["差異明細"].A4.v, "主要異動");
  assert.match(parsed.Sheets["差異明細"].A5.v, /所屬架構：69-ROOT › VB-BOARD-T/);
  assert.equal(parsed.Sheets["差異明細"].I4.v, "前版製造商名稱");
  assert.equal(parsed.Sheets["料號生命週期"].A4.v, "生命週期");
  assert.equal(parsed.Sheets["料號生命週期"].D4.v, "製造商名稱");
  assert.equal(parsed.Sheets["待人工確認"].A1.v, "待人工確認清單");
  assert.equal(parsed.Sheets["匯入稽核"].A1.v, "匯入稽核紀錄");
  assert.ok(buffer.byteLength > 10_000);
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
  assert.match(html, /料號生命週期/);
  assert.match(html, /待人工確認/);
  assert.match(html, /匯入稽核/);
  assert.match(html, /製造商名稱/);
  assert.match(html, /所屬架構/);
  assert.match(html, /69-ROOT › VB-BOARD-T/);
  assert.match(html, /Content-Security-Policy/);
  assert.match(html, /OLD&lt;&amp;/);
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

test("progressively indexes and caches schematic PDFs", async () => {
  const viewer = await readFile(new URL("../app/pdf-schematic-viewer.tsx", import.meta.url), "utf8");
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
  assert.match(viewer, /currentHit\?\.page === pageNumber/);
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
  assert.match(diff.fields.join(" "), /數量差異/);
  assert.deepEqual(diff.categories, ["added", "changed"]);
  assert.equal(diff.after?.qty, 3);

  const entirelyAdded = compareBom([], changed)[0];
  assert.equal(entirelyAdded.primaryType, "componentAdded");
  assert.match(entirelyAdded.fields.join(" "), /新增元件/);
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
  assert.deepEqual(added.fields, ["新增替料", "新增料號"]);
  assert.deepEqual(added.newParts.map((part) => part.part), ["ALT"]);

  const removed = compareBom(withSubstitute, base)[0];
  assert.equal(removed.kind, "removed");
  assert.equal(removed.primaryType, "substituteRemoved");
  assert.deepEqual(removed.categories, ["removed"]);
  assert.deepEqual(removed.fields, ["刪除替料", "刪除料號"]);
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

test("does not call a moved part globally new or globally deleted", () => {
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
