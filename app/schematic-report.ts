import { PDFDocument } from "pdf-lib";
import type { BomDiff } from "./bom-logic";
import { getPdfReportSource, type PdfReportSource } from "./pdf-schematic-viewer";
import { lookupReferenceHits, PDF_REFERENCE_INDEX_SCALE, type PdfTextBox } from "./pdf-search";
import { createSchematicReportPlan, type SchematicReportPlanEntry } from "./schematic-report-logic";

export type SchematicReportProgress = {
  current: number;
  total: number;
  message: string;
};

type ExportOptions = {
  beforeFile: File;
  afterFile: File;
  beforeName: string;
  afterName: string;
  diffs: BomDiff[];
  onProgress?: (progress: SchematicReportProgress) => void;
};

type LocatedReference = {
  hit?: PdfTextBox;
  hitCount: number;
  page?: number;
  status: "found" | "multiple" | "not-found" | "no-text";
  image?: HTMLCanvasElement;
};

const PAGE_WIDTH = 1240;
const PAGE_HEIGHT = 1754;
const PDF_PAGE_WIDTH = 595.28;
const PDF_PAGE_HEIGHT = 841.89;
const FONT_FAMILY = '"Microsoft JhengHei","Noto Sans TC",Arial,sans-serif';

function canvas(width = PAGE_WIDTH, height = PAGE_HEIGHT) {
  const element = document.createElement("canvas");
  element.width = width;
  element.height = height;
  return element;
}

function roundedRect(context: CanvasRenderingContext2D, x: number, y: number, width: number, height: number, radius: number) {
  context.beginPath();
  context.roundRect(x, y, width, height, radius);
}

function drawText(context: CanvasRenderingContext2D, text: string, x: number, y: number, options: {
  size?: number;
  color?: string;
  weight?: number;
  maxWidth?: number;
  lineHeight?: number;
} = {}) {
  const size = options.size ?? 28;
  const lineHeight = options.lineHeight ?? size * 1.4;
  context.fillStyle = options.color ?? "#253247";
  context.font = `${options.weight ?? 400} ${size}px ${FONT_FAMILY}`;
  context.textBaseline = "top";
  const maxWidth = options.maxWidth ?? PAGE_WIDTH - x - 70;
  const paragraphs = text.split("\n");
  let currentY = y;
  paragraphs.forEach((paragraph) => {
    const characters = Array.from(paragraph);
    let line = "";
    characters.forEach((character) => {
      const next = line + character;
      if (line && context.measureText(next).width > maxWidth) {
        context.fillText(line, x, currentY);
        currentY += lineHeight;
        line = character;
      } else line = next;
    });
    context.fillText(line || " ", x, currentY);
    currentY += lineHeight;
  });
  return currentY;
}

function basePage(title: string, subtitle: string) {
  const page = canvas();
  const context = page.getContext("2d")!;
  context.fillStyle = "#F3F5F8";
  context.fillRect(0, 0, PAGE_WIDTH, PAGE_HEIGHT);
  context.fillStyle = "#1F3A5F";
  context.fillRect(0, 0, PAGE_WIDTH, 180);
  drawText(context, title, 64, 45, { size: 45, color: "#FFFFFF", weight: 700 });
  drawText(context, subtitle, 66, 112, { size: 21, color: "#DCE9F8" });
  return { page, context };
}

function drawFooter(context: CanvasRenderingContext2D, pageNumber: number, totalPages: number) {
  context.strokeStyle = "#D9E0E8";
  context.lineWidth = 2;
  context.beginPath();
  context.moveTo(64, PAGE_HEIGHT - 78);
  context.lineTo(PAGE_WIDTH - 64, PAGE_HEIGHT - 78);
  context.stroke();
  drawText(context, "BOMLens 本機離線產生", 66, PAGE_HEIGHT - 58, { size: 17, color: "#667085" });
  drawText(context, `${pageNumber} / ${totalPages}`, PAGE_WIDTH - 145, PAGE_HEIGHT - 58, { size: 17, color: "#667085" });
}

function statusLabel(location: LocatedReference) {
  if (location.status === "found") return `找到・第 ${location.page} 頁`;
  if (location.status === "multiple") return `找到 ${location.hitCount} 筆・顯示第 1 筆（第 ${location.page} 頁）`;
  if (location.status === "no-text") return "PDF 無可搜尋文字";
  return "找不到插件位置";
}

async function renderReference(source: PdfReportSource, reference: string): Promise<LocatedReference> {
  if (source.status === "no-text") return { hitCount: 0, status: "no-text" };
  const hits = lookupReferenceHits(source.referenceIndex, reference);
  const hit = hits[0];
  if (!hit) return { hitCount: 0, status: "not-found" };
  const page = await source.document.getPage(hit.page);
  const renderScale = 1.65;
  const viewport = page.getViewport({ scale: renderScale });
  const full = canvas(Math.ceil(viewport.width), Math.ceil(viewport.height));
  const fullContext = full.getContext("2d")!;
  fullContext.fillStyle = "#FFFFFF";
  fullContext.fillRect(0, 0, full.width, full.height);
  await page.render({ canvas: full, canvasContext: fullContext, viewport }).promise;

  const output = canvas(540, 560);
  const context = output.getContext("2d")!;
  context.fillStyle = "#FFFFFF";
  context.fillRect(0, 0, output.width, output.height);
  const hitScale = renderScale / PDF_REFERENCE_INDEX_SCALE;
  const hitCenterX = (hit.x + hit.width / 2) * hitScale;
  const hitCenterY = (hit.y + hit.height / 2) * hitScale;
  const cropWidth = Math.min(full.width, 720);
  const cropHeight = Math.min(full.height, 720);
  const sourceX = Math.min(Math.max(0, hitCenterX - cropWidth / 2), Math.max(0, full.width - cropWidth));
  const sourceY = Math.min(Math.max(0, hitCenterY - cropHeight / 2), Math.max(0, full.height - cropHeight));
  context.drawImage(full, sourceX, sourceY, cropWidth, cropHeight, 0, 0, output.width, output.height);
  const factorX = output.width / cropWidth;
  const factorY = output.height / cropHeight;
  context.strokeStyle = "#E13535";
  context.lineWidth = 6;
  context.strokeRect(
    (hit.x * hitScale - sourceX - 7) * factorX,
    (hit.y * hitScale - sourceY - 7) * factorY,
    Math.max(24, (hit.width * hitScale + 14) * factorX),
    Math.max(24, (hit.height * hitScale + 14) * factorY),
  );
  full.width = 1;
  full.height = 1;
  return { hit, hitCount: hits.length, page: hit.page, status: hits.length > 1 ? "multiple" : "found", image: output };
}

function drawLocationCard(context: CanvasRenderingContext2D, x: number, y: number, width: number, title: string, location: LocatedReference) {
  roundedRect(context, x, y, width, 710, 18);
  context.fillStyle = "#FFFFFF";
  context.fill();
  context.strokeStyle = "#D9E0E8";
  context.lineWidth = 2;
  context.stroke();
  drawText(context, title, x + 24, y + 24, { size: 27, color: "#1F3A5F", weight: 700, maxWidth: width - 48 });
  const warning = location.status !== "found";
  drawText(context, statusLabel(location), x + 24, y + 69, { size: 20, color: warning ? "#9A651B" : "#16835D", weight: 700, maxWidth: width - 48 });
  if (location.image) {
    context.drawImage(location.image, x + 20, y + 118, width - 40, 520);
  } else {
    context.fillStyle = warning ? "#FFF3DC" : "#F4F6F8";
    context.fillRect(x + 20, y + 118, width - 40, 520);
    drawText(context, location.status === "no-text" ? "此 PDF 沒有可搜尋文字層\n無法依插件位置自動截圖" : "此版本未定位到該插件位置", x + 60, y + 320, { size: 24, color: warning ? "#9A651B" : "#667085", weight: 700, maxWidth: width - 120 });
  }
}

function coverPage(options: ExportOptions, entries: SchematicReportPlanEntry[], beforeSource: PdfReportSource, afterSource: PdfReportSource, totalPages: number) {
  const { page, context } = basePage("線路圖差異比較報告", "依 BOM 異動插件位置產生・完全本機離線");
  drawText(context, "比較來源", 65, 230, { size: 25, color: "#1F3A5F", weight: 700 });
  roundedRect(context, 64, 278, PAGE_WIDTH - 128, 190, 18);
  context.fillStyle = "#FFFFFF";
  context.fill();
  context.strokeStyle = "#D9E0E8";
  context.stroke();
  drawText(context, `舊版：${options.beforeName}`, 94, 312, { size: 23, weight: 700 });
  drawText(context, `新版：${options.afterName}`, 94, 363, { size: 23, weight: 700 });
  drawText(context, `頁數：${beforeSource.totalPages} → ${afterSource.totalPages}`, 94, 414, { size: 20, color: "#667085" });

  const cards = [
    ["報告插件位置", String(entries.length), "#2F6BCE", "#EAF1FD"],
    ["舊版可搜尋", beforeSource.status === "ready" ? "是" : "否", beforeSource.status === "ready" ? "#16835D" : "#9A651B", beforeSource.status === "ready" ? "#E8F7F0" : "#FFF3DC"],
    ["新版可搜尋", afterSource.status === "ready" ? "是" : "否", afterSource.status === "ready" ? "#16835D" : "#9A651B", afterSource.status === "ready" ? "#E8F7F0" : "#FFF3DC"],
  ];
  cards.forEach(([label, value, color, background], index) => {
    const x = 64 + index * 374;
    roundedRect(context, x, 515, 342, 175, 18);
    context.fillStyle = background;
    context.fill();
    drawText(context, label, x + 26, 545, { size: 21, color, weight: 700, maxWidth: 290 });
    drawText(context, value, x + 26, 594, { size: 48, color, weight: 700, maxWidth: 290 });
  });

  drawText(context, "報告內容", 65, 760, { size: 25, color: "#1F3A5F", weight: 700 });
  const preview = entries.slice(0, 14);
  preview.forEach((entry, index) => {
    const y = 815 + index * 52;
    context.fillStyle = index % 2 ? "#FAFBFC" : "#FFFFFF";
    context.fillRect(64, y, PAGE_WIDTH - 128, 45);
    drawText(context, entry.reference, 88, y + 8, { size: 20, color: "#253247", weight: 700, maxWidth: 130 });
    drawText(context, entry.changeLabels.join("、"), 235, y + 8, { size: 19, color: "#526176", maxWidth: 900 });
  });
  if (entries.length > preview.length) drawText(context, `另有 ${entries.length - preview.length} 筆，請見後續明細頁。`, 70, 815 + preview.length * 52 + 12, { size: 19, color: "#667085" });
  drawFooter(context, 1, totalPages);
  return page;
}

function detailPage(entry: SchematicReportPlanEntry, before: LocatedReference, after: LocatedReference, pageNumber: number, totalPages: number) {
  const { page, context } = basePage(entry.reference, entry.changeLabels.join("・"));
  roundedRect(context, 64, 220, PAGE_WIDTH - 128, 130, 16);
  context.fillStyle = "#FFFFFF";
  context.fill();
  context.strokeStyle = "#D9E0E8";
  context.stroke();
  drawText(context, entry.detail, 92, 252, { size: 22, color: "#526176", maxWidth: PAGE_WIDTH - 184, lineHeight: 32 });
  drawLocationCard(context, 64, 395, 540, "舊版線路圖", before);
  drawLocationCard(context, 636, 395, 540, "新版線路圖", after);

  const reliability = before.status === "found" && after.status === "found"
    ? "定位可靠度：高 - 新舊版本皆為單一文字命中"
    : before.status === "multiple" || after.status === "multiple"
      ? "定位可靠度：需確認 - 至少一個版本有多筆文字命中"
      : "定位可靠度：低 - 至少一個版本無法自動定位";
  roundedRect(context, 64, 1140, PAGE_WIDTH - 128, 135, 16);
  context.fillStyle = reliability.includes("高") ? "#E8F7F0" : reliability.includes("需確認") ? "#FFF3DC" : "#FFF0F0";
  context.fill();
  drawText(context, reliability, 92, 1174, { size: 23, color: reliability.includes("高") ? "#16835D" : reliability.includes("需確認") ? "#9A651B" : "#B14949", weight: 700 });
  drawText(context, "紅框為文字索引定位點；截圖顯示該位置附近範圍，供 RD 快速審核。", 92, 1221, { size: 19, color: "#667085" });
  drawFooter(context, pageNumber, totalPages);
  return page;
}

async function addCanvasPage(pdf: PDFDocument, reportPage: HTMLCanvasElement) {
  const jpg = reportPage.toDataURL("image/jpeg", 0.9);
  const image = await pdf.embedJpg(jpg);
  const page = pdf.addPage([PDF_PAGE_WIDTH, PDF_PAGE_HEIGHT]);
  page.drawImage(image, { x: 0, y: 0, width: PDF_PAGE_WIDTH, height: PDF_PAGE_HEIGHT });
  reportPage.width = 1;
  reportPage.height = 1;
}

function savePdf(bytes: Uint8Array) {
  const blob = new Blob([bytes], { type: "application/pdf" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `Schematic_diff_report_${new Date().toISOString().slice(0, 10).replaceAll("-", "")}.pdf`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

export async function exportSchematicPdfReport(options: ExportOptions) {
  const entries = createSchematicReportPlan(options.diffs);
  if (!entries.length) throw new Error("目前 BOM 差異沒有可用的插件位置");
  options.onProgress?.({ current: 0, total: entries.length, message: "建立舊版線路圖索引…" });
  const beforeSource = await getPdfReportSource(options.beforeFile, (processed, total) => {
    options.onProgress?.({ current: 0, total: entries.length, message: `舊版索引 ${processed}／${total || "…"} 頁` });
  });
  options.onProgress?.({ current: 0, total: entries.length, message: "建立新版線路圖索引…" });
  const afterSource = await getPdfReportSource(options.afterFile, (processed, total) => {
    options.onProgress?.({ current: 0, total: entries.length, message: `新版索引 ${processed}／${total || "…"} 頁` });
  });
  const totalPages = entries.length + 1;
  const pdf = await PDFDocument.create();
  pdf.setTitle("BOMLens 線路圖差異比較報告");
  pdf.setSubject(`${options.beforeName} → ${options.afterName}`);
  pdf.setCreator("BOMLens Offline");
  await addCanvasPage(pdf, coverPage(options, entries, beforeSource, afterSource, totalPages));
  for (let index = 0; index < entries.length; index += 1) {
    const entry = entries[index];
    options.onProgress?.({ current: index + 1, total: entries.length, message: `擷取 ${entry.reference} 附近線路圖…` });
    const before = await renderReference(beforeSource, entry.reference);
    const after = await renderReference(afterSource, entry.reference);
    await addCanvasPage(pdf, detailPage(entry, before, after, index + 2, totalPages));
    await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
  }
  options.onProgress?.({ current: entries.length, total: entries.length, message: "完成 PDF 報告…" });
  const bytes = await pdf.save({ useObjectStreams: true });
  savePdf(bytes);
  return { entries: entries.length, pages: totalPages, bytes };
}
