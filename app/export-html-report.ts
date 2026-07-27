import { bomDiffDisplayFields, bomStructureLabel, sortBomDiffsForAll, type BomAlternative, type BomDiff, type DiffPrimaryType } from "./bom-logic.ts";
import type { ReportContext } from "./export-report";

const primaryLabels: Record<DiffPrimaryType, string> = {
  componentAdded: "新增",
  componentRemoved: "刪除",
  substituteAdded: "新增",
  substituteRemoved: "刪除",
  partReplaced: "變更",
  positionChanged: "變更",
  same: "相同",
};

function escapeHtml(value: unknown) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function partLabel(part: BomAlternative) {
  return part.part || part.manufacturerPart || "未提供料號";
}

function joinParts(parts: BomAlternative[]) {
  return parts.map((part) => escapeHtml(partLabel(part))).join("<br>") || "—";
}

function joinValues(values: string[]) {
  return values.map(escapeHtml).join("、") || "—";
}

function typeTone(type: DiffPrimaryType) {
  if (type === "componentAdded" || type === "substituteAdded") return "added";
  if (type === "componentRemoved" || type === "substituteRemoved") return "removed";
  if (type === "partReplaced") return "replacement";
  return "changed";
}

function diffStructureLabel(diff: BomDiff) {
  const before = bomStructureLabel(diff.before);
  const after = bomStructureLabel(diff.after);
  if (before && after && before !== after) return `${before} → ${after}`;
  return after || before;
}

function structureHtml(diff: BomDiff) {
  const structure = diffStructureLabel(diff);
  return structure ? `<small class="structure"><b>所屬架構</b> ${escapeHtml(structure)}</small>` : "";
}

export function buildBomHtmlReport(diffs: BomDiff[], beforeName: string, afterName: string, context: ReportContext = {}) {
  const orderedDiffs = sortBomDiffsForAll(diffs);
  const newParts = new Set(diffs.flatMap((diff) => diff.newParts.map((part) => part.part)));
  const deletedParts = new Set(diffs.flatMap((diff) => diff.deletedParts.map((part) => part.part)));
  const detailRows = orderedDiffs.map((diff) => `
    <tr>
      <td><span class="type ${typeTone(diff.primaryType)}">${escapeHtml(primaryLabels[diff.primaryType])}</span>${structureHtml(diff)}</td>
      <td>${joinValues(bomDiffDisplayFields(diff))}</td>
      <td><span class="plus">${diff.addedParts.length ? `＋ ${joinParts(diff.addedParts)}` : "—"}</span><span class="minus">${diff.removedParts.length ? `－ ${joinParts(diff.removedParts)}` : ""}</span></td>
      <td><span class="plus">${diff.addedPositions.length ? `＋ ${joinValues(diff.addedPositions)}` : "—"}</span><span class="minus">${diff.removedPositions.length ? `－ ${joinValues(diff.removedPositions)}` : ""}</span>${diff.replacementPositions.length ? `<span class="change">${joinValues(diff.replacementPositions)} 換料</span>` : ""}</td>
      <td><strong>${diff.before ? joinParts(diff.before.alternatives) : "—"}</strong><small><b>製造廠商料號</b> ${diff.before ? joinValues(diff.before.alternatives.map((part) => part.manufacturerPart).filter(Boolean)) : "—"}</small><small><b>製造廠商</b> ${diff.before ? joinValues(diff.before.alternatives.map((part) => part.manufacturerName ?? "").filter(Boolean)) : "—"}</small></td>
      <td><strong>${diff.after ? joinParts(diff.after.alternatives) : "—"}</strong><small><b>製造廠商料號</b> ${diff.after ? joinValues(diff.after.alternatives.map((part) => part.manufacturerPart).filter(Boolean)) : "—"}</small><small><b>製造廠商</b> ${diff.after ? joinValues(diff.after.alternatives.map((part) => part.manufacturerName ?? "").filter(Boolean)) : "—"}</small></td>
      <td class="qty">${diff.before?.qty ?? 0} → ${diff.after?.qty ?? 0}</td>
      <td>${diff.needsReview ? `<span class="confidence low">待人工確認</span><small>${escapeHtml(diff.matchReason)}</small>` : "—"}</td>
    </tr>`).join("");

  const lifecycleRows = orderedDiffs.flatMap((diff) => [
    ...diff.newParts.map((part) => `<tr><td><span class="type added">新版完全新料</span></td><td>${escapeHtml(partLabel(part))}</td><td>${escapeHtml(part.manufacturerPart || "—")}</td><td>${escapeHtml(part.manufacturerName || "—")}</td><td>${joinValues(diff.after?.positions ?? [])}</td><td>${diff.after?.qty ?? 0}</td><td>${escapeHtml(primaryLabels[diff.primaryType])}${structureHtml(diff)}</td></tr>`),
    ...diff.deletedParts.map((part) => `<tr><td><span class="type removed">新版完全移除</span></td><td>${escapeHtml(partLabel(part))}</td><td>${escapeHtml(part.manufacturerPart || "—")}</td><td>${escapeHtml(part.manufacturerName || "—")}</td><td>${joinValues(diff.before?.positions ?? [])}</td><td>${diff.before?.qty ?? 0}</td><td>${escapeHtml(primaryLabels[diff.primaryType])}${structureHtml(diff)}</td></tr>`),
  ]).join("");
  const reviewRows = orderedDiffs.filter((diff) => diff.needsReview).map((diff) => `<tr><td>${escapeHtml(primaryLabels[diff.primaryType])}${structureHtml(diff)}</td><td>${diff.before ? joinParts(diff.before.alternatives) : "—"}</td><td>${diff.after ? joinParts(diff.after.alternatives) : "—"}</td><td>${joinValues(diff.after?.positions ?? diff.before?.positions ?? [])}</td><td>${escapeHtml(diff.matchReason)}</td><td>${escapeHtml(diff.before?.sourceRows?.join(", ") || "—")}</td><td>${escapeHtml(diff.after?.sourceRows?.join(", ") || "—")}</td></tr>`).join("");
  const auditRows = ([{ label: "舊版", source: context.before }, { label: "新版", source: context.after }] as const).flatMap(({ label, source }) => {
    if (!source) return [];
    return source.audit.issues.map((issue) => `<tr><td>${label}</td><td>${escapeHtml(source.fileName)}</td><td>${issue.severity === "error" ? "錯誤" : "警告"}</td><td>${escapeHtml(issue.message)}</td><td>${escapeHtml(issue.rows?.join(", ") || "—")}</td></tr>`);
  }).join("");
  const lifecycleSection = lifecycleRows ? `<section><h2>新版新料／移除清單</h2><p class="note">只列出需要採購注意的完全新料與新版已完全移除料號。</p><div class="table-wrap"><table><thead><tr><th>狀態</th><th>主件料號</th><th>製造廠商料號</th><th>製造廠商</th><th>插件位置</th><th>數量</th><th>主要異動</th></tr></thead><tbody>${lifecycleRows}</tbody></table></div></section>` : "";
  const reviewSection = reviewRows ? `<section><h2>待人工確認</h2><p class="note">只在配對不明確時顯示。</p><div class="table-wrap"><table><thead><tr><th>主要異動</th><th>舊版料號</th><th>新版料號</th><th>插件位置</th><th>配對依據</th><th>舊版原始列</th><th>新版原始列</th></tr></thead><tbody>${reviewRows}</tbody></table></div></section>` : "";
  const auditSection = auditRows ? `<section><h2>匯入警告</h2><div class="table-wrap"><table><thead><tr><th>版本</th><th>檔案</th><th>類型</th><th>內容</th><th>原始列</th></tr></thead><tbody>${auditRows}</tbody></table></div></section>` : "";

  return `<!doctype html>
<html lang="zh-Hant">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; script-src 'unsafe-inline'; img-src data:">
  <title>BOM 版本差異報告</title>
  <style>
    :root{--navy:#1f3a5f;--blue:#2f6bce;--green:#16835d;--red:#b14949;--amber:#9a651b;--line:#d9e0e8;--muted:#667085}*{box-sizing:border-box}body{margin:0;background:#f3f5f8;color:#253247;font:14px/1.55 "Microsoft JhengHei","Noto Sans TC",Arial,sans-serif}.report{max-width:1500px;margin:28px auto;padding:0 24px}.hero{padding:28px 32px;border-radius:14px;background:linear-gradient(135deg,var(--navy),#315b8d);color:white;box-shadow:0 8px 28px #1f3a5f22}.hero h1{margin:0 0 8px;font-size:27px}.hero p{margin:0;color:#dce9f8}.actions{display:flex;justify-content:flex-end;margin:14px 0}.actions button{padding:9px 16px;border:0;border-radius:7px;background:var(--blue);color:white;font-weight:700;cursor:pointer}.files{display:grid;grid-template-columns:1fr auto 1fr;gap:14px;align-items:center;margin:18px 0;padding:18px 22px;border:1px solid var(--line);border-radius:10px;background:white}.file small{display:block;color:var(--muted)}.file strong{font-size:16px}.files .arrow{color:#98a2b3;font-size:22px}.cards{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin:18px 0}.card{padding:18px 20px;border:1px solid var(--line);border-radius:10px;background:white}.card small{color:var(--muted)}.card strong{display:block;margin-top:5px;font-size:28px}.card.new strong{color:var(--green)}.card.removed strong{color:var(--red)}section{margin-top:20px;padding:20px;border:1px solid var(--line);border-radius:12px;background:white;box-shadow:0 2px 8px #1623370a}h2{margin:0 0 14px;font-size:18px;color:var(--navy)}.table-wrap{overflow:auto}table{width:100%;min-width:980px;border-collapse:separate;border-spacing:0}th{position:sticky;top:0;padding:10px 9px;background:var(--blue);color:white;font-size:12px;text-align:left;white-space:nowrap}td{padding:10px 9px;border-bottom:1px solid #e8ecf1;vertical-align:top;font-size:12px}tbody tr:nth-child(even){background:#fafbfc}td small{display:block;margin-top:4px;color:var(--muted)}td span{display:block}.type{display:inline-block!important;padding:4px 7px;border-radius:5px;font-weight:800;white-space:nowrap}.type.added,.plus{color:var(--green);background:#e8f7f0}.type.removed,.minus{color:var(--red);background:#fff0f0}.type.changed{color:var(--blue);background:#eaf1fd}.type.replacement,.change{color:var(--amber);background:#fff3dc}.plus,.minus,.change{margin-bottom:4px;padding:3px 5px;border-radius:4px}.structure{max-width:190px;color:#526176;font-weight:600;overflow-wrap:anywhere}.qty{font-weight:700;white-space:nowrap}.confidence{font-weight:800}.confidence.low{color:var(--red)}.note{color:var(--muted);font-size:12px}.empty{padding:20px;color:var(--muted);text-align:center}@media(max-width:800px){.report{padding:0 12px}.cards{grid-template-columns:1fr}.files{grid-template-columns:1fr}.files .arrow{transform:rotate(90deg);text-align:center}}@media print{body{background:white}.report{max-width:none;margin:0;padding:0}.actions{display:none}.hero,section,.files,.card{box-shadow:none;break-inside:avoid}section{border-radius:0}.table-wrap{overflow:visible}table{min-width:0;font-size:9px}th,td{padding:5px 4px}}
  </style>
</head>
<body>
  <main class="report">
    <header class="hero"><h1>BOM 版本差異報告</h1><p>完全離線產生，不包含任何外部連線或資源。</p></header>
    <div class="actions"><button type="button" onclick="window.print()">列印／另存 PDF</button></div>
    <div class="files"><div class="file"><small>舊版 BOM</small><strong>${escapeHtml(beforeName)}</strong></div><div class="arrow">→</div><div class="file"><small>新版 BOM</small><strong>${escapeHtml(afterName)}</strong></div></div>
    <div class="cards"><div class="card"><small>差異群組</small><strong>${diffs.length}</strong></div><div class="card new"><small>新版完全新料</small><strong>${newParts.size}</strong></div><div class="card removed"><small>新版完全移除</small><strong>${deletedParts.size}</strong></div></div>
    <section><h2>差異明細</h2><div class="table-wrap"><table><thead><tr><th>主要異動</th><th>差異項目</th><th>料號異動</th><th>插件位置差異</th><th>舊版主件料號／製造廠商資訊</th><th>新版主件料號／製造廠商資訊</th><th>數量</th><th>人工確認</th></tr></thead><tbody>${detailRows || `<tr><td colspan="8" class="empty">沒有差異</td></tr>`}</tbody></table></div></section>
    ${lifecycleSection}
    ${reviewSection}
    ${auditSection}
  </main>
</body>
</html>`;
}

export function exportBomHtmlReport(diffs: BomDiff[], beforeName: string, afterName: string, context: ReportContext = {}) {
  const html = buildBomHtmlReport(diffs, beforeName, afterName, context);
  const blob = new Blob([html], { type: "text/html;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `BOM_diff_report_${new Date().toISOString().slice(0, 10).replaceAll("-", "")}.html`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}
