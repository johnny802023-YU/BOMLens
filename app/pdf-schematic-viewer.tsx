"use client";
/* eslint-disable react-hooks/set-state-in-effect -- file and search changes intentionally reset the local PDF viewer state */

import { ChevronLeft, ChevronRight, FileSearch } from "lucide-react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import legacyPdfWorkerUrl from "pdfjs-dist/legacy/build/pdf.worker.min.mjs?url";
import type { PDFDocumentProxy, PDFPageProxy } from "pdfjs-dist";
import { buildPageReferenceIndex, centeredPdfHitScroll, centeredRenderedHitScroll, lookupReferenceHits, mergeReferenceIndex, normalizeReference, PDF_REFERENCE_INDEX_SCALE, type PdfReferenceIndex, type PdfTextBox } from "./pdf-search";

type IndexStatus = "indexing" | "ready" | "no-text" | "error";
type PdfLoadMode = "standard" | "compatibility";
type PdfFailureKind = "browser-unsupported" | "worker-blocked" | "encrypted" | "invalid" | "unknown";
type PdfFailure = { kind: PdfFailureKind; message: string };
type PdfJsModule = typeof import("pdfjs-dist");
export type PdfViewerSide = "before" | "after";
export type PdfScrollSync = { source: PdfViewerSide; x: number; y: number; revision: number };
type IndexEntry = {
  document: PDFDocumentProxy | null;
  referenceIndex: PdfReferenceIndex;
  textItemCount: number;
  processedPages: number;
  totalPages: number;
  batchSize: number;
  status: IndexStatus;
  mode: PdfLoadMode;
  compatibilityReason: PdfFailure | null;
  failure: PdfFailure | null;
  started: boolean;
  cancelled: boolean;
  lastUsed: number;
  listeners: Set<() => void>;
};

const MAX_CACHED_PDFS = 2;
const PDF_LOAD_TIMEOUT_MS = 12_000;
const pdfIndexCache = new Map<File, IndexEntry>();

function browserCompatibilityFailure(): PdfFailure | null {
  const promiseWithResolvers = (Promise as PromiseConstructor & { withResolvers?: unknown }).withResolvers;
  const urlParse = (URL as typeof URL & { parse?: unknown }).parse;
  const abortSignalAny = (AbortSignal as typeof AbortSignal & { any?: unknown }).any;
  if (typeof promiseWithResolvers !== "function" || typeof urlParse !== "function" || typeof abortSignalAny !== "function") {
    return { kind: "browser-unsupported", message: "瀏覽器版本較舊，已自動切換 PDF 相容模式" };
  }
  return null;
}

function errorText(error: unknown) {
  if (error instanceof Error) return `${error.name}: ${error.message}`;
  return String(error ?? "未知錯誤");
}

export function classifyPdfFailure(error: unknown): PdfFailure {
  const text = errorText(error);
  const normalized = text.toLowerCase();
  if (/passwordexception|password|密碼|encrypted/.test(normalized)) {
    return { kind: "encrypted", message: "PDF 已加密或需要密碼，請先解除密碼後再匯入" };
  }
  if (/invalidpdfexception|invalid pdf|formaterror|xref|bad (?:page|catalog)|corrupt|damaged|pdf file is empty|unexpected eof|end of file/.test(normalized)) {
    return { kind: "invalid", message: "PDF 檔案損壞或格式不完整，請重新輸出 PDF 後再匯入" };
  }
  if (/worker|timeout|timed out|逾時|module script|failed to fetch dynamically imported module/.test(normalized)) {
    return { kind: "worker-blocked", message: "PDF Worker 被瀏覽器或公司資安政策阻擋" };
  }
  if (/withresolvers|url\.parse|abortsignal\.any|syntaxerror|not supported/.test(normalized)) {
    return { kind: "browser-unsupported", message: "瀏覽器版本過舊，不支援目前的 PDF 解析功能" };
  }
  const detail = text.replace(/^\w+Error:\s*/i, "").slice(0, 140);
  return { kind: "unknown", message: `PDF 讀取失敗${detail ? `：${detail}` : ""}` };
}

function terminalFailureMessage(failure: PdfFailure) {
  if (failure.kind === "worker-blocked") return `${failure.message}，相容模式仍無法載入；請聯絡 IT 檢查 localhost 與模組載入政策`;
  if (failure.kind === "browser-unsupported") return `${failure.message}；請更新 Microsoft Edge 或 Chrome`;
  return failure.message;
}

async function withTimeout<T>(promise: Promise<T>, timeoutMs: number, onTimeout: () => void): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const timeout = new Promise<never>((_, reject) => {
    timer = setTimeout(() => {
      onTimeout();
      reject(new Error(`PDF Worker 啟動逾時（${Math.round(timeoutMs / 1000)} 秒）`));
    }, timeoutMs);
  });
  try {
    return await Promise.race([promise, timeout]);
  } finally {
    if (timer) clearTimeout(timer);
  }
}

async function loadPdfDocument(file: File, mode: PdfLoadMode) {
  const pdfjs = (mode === "compatibility" ? await import("pdfjs-dist/legacy/build/pdf.mjs") : await import("pdfjs-dist")) as PdfJsModule;
  pdfjs.GlobalWorkerOptions.workerSrc = mode === "compatibility" ? legacyPdfWorkerUrl : pdfWorkerUrl;
  const scope = globalThis as typeof globalThis & { pdfjsWorker?: unknown };
  const previousWorkerModule = scope.pdfjsWorker;
  let loadingTask: ReturnType<PdfJsModule["getDocument"]>;
  try {
    if (mode === "compatibility") {
      const workerModule = await import(/* @vite-ignore */ legacyPdfWorkerUrl) as { WorkerMessageHandler: unknown };
      scope.pdfjsWorker = { WorkerMessageHandler: workerModule.WorkerMessageHandler };
    }
    const data = await file.arrayBuffer();
    loadingTask = pdfjs.getDocument({ data, isEvalSupported: mode === "standard", useWasm: mode === "standard" });
  } finally {
    if (mode === "compatibility") {
      if (previousWorkerModule === undefined) delete scope.pdfjsWorker;
      else scope.pdfjsWorker = previousWorkerModule;
    }
  }
  try {
    const document = await withTimeout(loadingTask.promise, PDF_LOAD_TIMEOUT_MS, () => { void loadingTask.destroy(); });
    return { document, pdfjs };
  } catch (error) {
    await loadingTask.destroy().catch(() => undefined);
    throw error;
  }
}

function disposeEntry(file: File, entry: IndexEntry) {
  entry.cancelled = true;
  void entry.document?.destroy();
  pdfIndexCache.delete(file);
}

function trimPdfCache(protectedFile?: File) {
  while (pdfIndexCache.size > MAX_CACHED_PDFS) {
    const candidate = [...pdfIndexCache.entries()]
      .filter(([file, entry]) => file !== protectedFile && entry.listeners.size === 0)
      .sort((a, b) => a[1].lastUsed - b[1].lastUsed)[0];
    if (!candidate) return;
    disposeEntry(candidate[0], candidate[1]);
  }
}

export function clearInactivePdfCache() {
  [...pdfIndexCache.entries()].forEach(([file, entry]) => { if (entry.listeners.size === 0) disposeEntry(file, entry); });
}

export type PdfReportSource = {
  document: PDFDocumentProxy;
  referenceIndex: PdfReferenceIndex;
  totalPages: number;
  status: Exclude<IndexStatus, "indexing">;
};

function getIndexEntry(file: File) {
  const existing = pdfIndexCache.get(file);
  if (existing) { existing.lastUsed = Date.now(); return existing; }
  const entry: IndexEntry = { document: null, referenceIndex: new Map(), textItemCount: 0, processedPages: 0, totalPages: 0, batchSize: 2, status: "indexing", mode: "standard", compatibilityReason: null, failure: null, started: false, cancelled: false, lastUsed: Date.now(), listeners: new Set() };
  pdfIndexCache.set(file, entry);
  trimPdfCache(file);
  return entry;
}

function notify(entry: IndexEntry) {
  entry.listeners.forEach((listener) => listener());
}

async function startProgressiveIndex(file: File, entry: IndexEntry) {
  if (entry.started) return;
  entry.started = true;
  const unsupportedBrowser = browserCompatibilityFailure();
  const modes: PdfLoadMode[] = unsupportedBrowser ? ["compatibility"] : ["standard", "compatibility"];
  let lastFailure: PdfFailure = unsupportedBrowser ?? { kind: "unknown", message: "PDF 讀取失敗" };
  for (const mode of modes) {
    entry.mode = mode;
    entry.compatibilityReason = mode === "compatibility" ? lastFailure : null;
    entry.failure = null;
    entry.document = null;
    entry.referenceIndex = new Map();
    entry.textItemCount = 0;
    entry.processedPages = 0;
    entry.totalPages = 0;
    notify(entry);
    try {
      const loaded = await loadPdfDocument(file, mode);
      const pdfjs = loaded.pdfjs;
      entry.document = loaded.document;
      if (entry.cancelled) { void entry.document.destroy(); return; }
      entry.totalPages = entry.document.numPages;
      entry.batchSize = mode === "compatibility" ? 1 : Math.min(4, Math.max(2, Math.floor((navigator.hardwareConcurrency || 4) / 2)));
      notify(entry);
      for (let firstPage = 1; firstPage <= entry.document.numPages; firstPage += entry.batchSize) {
        if (entry.cancelled) return;
        const pageNumbers = Array.from({ length: Math.min(entry.batchSize, entry.document.numPages - firstPage + 1) }, (_, index) => firstPage + index);
        const results = await Promise.all(pageNumbers.map(async (pageNumber) => {
          const page = await entry.document!.getPage(pageNumber);
          const viewport = page.getViewport({ scale: PDF_REFERENCE_INDEX_SCALE });
          const content = await page.getTextContent();
          const boxes: PdfTextBox[] = [];
          content.items.forEach((rawItem) => {
            if (!("str" in rawItem) || !rawItem.str.trim()) return;
            const transform = pdfjs.Util.transform(viewport.transform, rawItem.transform);
            const radians = Math.atan2(transform[1], transform[0]);
            const rotation = (Math.round(radians * 180 / Math.PI / 90) * 90 + 360) % 360;
            const extent = Math.max(12, rawItem.width * viewport.scale);
            const fontHeight = Math.max(10, Math.hypot(transform[2], transform[3]));
            const direction = { x: Math.cos(radians), y: Math.sin(radians) };
            const perpendicular = { x: -direction.y, y: direction.x };
            const origin = { x: transform[4], y: transform[5] };
            const corners = [origin, { x: origin.x + direction.x * extent, y: origin.y + direction.y * extent }, { x: origin.x - perpendicular.x * fontHeight, y: origin.y - perpendicular.y * fontHeight }, { x: origin.x + direction.x * extent - perpendicular.x * fontHeight, y: origin.y + direction.y * extent - perpendicular.y * fontHeight }];
            const xs = corners.map((point) => point.x); const ys = corners.map((point) => point.y);
            const left = Math.min(...xs); const top = Math.min(...ys); const right = Math.max(...xs); const bottom = Math.max(...ys);
            boxes.push({ page: pageNumber, text: rawItem.str.toUpperCase(), x: left, y: top, width: Math.max(10, right - left), height: Math.max(10, bottom - top), rotation });
          });
          return { pageNumber, boxes, index: buildPageReferenceIndex(boxes) };
        }));
        if (entry.cancelled) return;
        results.sort((a, b) => a.pageNumber - b.pageNumber).forEach((result) => {
          entry.textItemCount += result.boxes.length;
          mergeReferenceIndex(entry.referenceIndex, result.index);
        });
        entry.processedPages += results.length;
        notify(entry);
        await new Promise<void>((resolve) => window.setTimeout(resolve, 0));
      }
      entry.status = entry.textItemCount ? "ready" : "no-text";
      notify(entry);
      return;
    } catch (error) {
      if (entry.cancelled) return;
      lastFailure = classifyPdfFailure(error);
      await entry.document?.destroy().catch(() => undefined);
      entry.document = null;
      if (lastFailure.kind === "encrypted" || lastFailure.kind === "invalid" || mode === "compatibility") break;
    }
  }
  if (!entry.cancelled) {
    entry.failure = lastFailure;
    entry.status = "error";
    notify(entry);
  }
}

export async function getPdfReportSource(file: File, onProgress?: (processedPages: number, totalPages: number) => void): Promise<PdfReportSource> {
  const entry = getIndexEntry(file);
  onProgress?.(entry.processedPages, entry.totalPages);
  void startProgressiveIndex(file, entry);
  if (entry.status === "indexing") {
    await new Promise<void>((resolve) => {
      const update = () => {
        onProgress?.(entry.processedPages, entry.totalPages);
        if (entry.status !== "indexing") {
          entry.listeners.delete(update);
          resolve();
        }
      };
      entry.listeners.add(update);
      update();
    });
  }
  if (!entry.document || entry.status === "error" || entry.status === "indexing") throw new Error(entry.failure ? terminalFailureMessage(entry.failure) : "PDF 讀取或索引建立失敗");
  const completedStatus = entry.status;
  entry.lastUsed = Date.now();
  return {
    document: entry.document,
    referenceIndex: new Map(entry.referenceIndex),
    totalPages: entry.totalPages,
    status: completedStatus,
  };
}

export function PdfSchematicViewer({ file, target, sideLabel, side, scale, syncEnabled, syncState, onScaleChange, onSyncScroll }: { file: File; target: string; sideLabel: string; side: PdfViewerSide; scale: number; syncEnabled: boolean; syncState: PdfScrollSync; onScaleChange: (side: PdfViewerSide, scale: number) => void; onSyncScroll: (side: PdfViewerSide, x: number, y: number) => void }) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const pageRef = useRef<HTMLDivElement>(null);
  const highlightRef = useRef<HTMLSpanElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const renderTask = useRef<{ cancel(): void } | null>(null);
  const applyingSyncedScroll = useRef(false);
  const scrollFrame = useRef<number | null>(null);
  const locateFrame = useRef<number | null>(null);
  const lastScrollRatio = useRef({ x: 0, y: 0 });
  const locatedTarget = useRef("");
  const [document, setDocument] = useState<PDFDocumentProxy | null>(null);
  const [referenceIndex, setReferenceIndex] = useState<PdfReferenceIndex>(new Map());
  const [pageNumber, setPageNumber] = useState(1);
  const [pageSize, setPageSize] = useState({ width: 1, height: 1 });
  const [status, setStatus] = useState<IndexStatus>("indexing");
  const [processedPages, setProcessedPages] = useState(0);
  const [totalPages, setTotalPages] = useState(0);
  const [batchSize, setBatchSize] = useState(2);
  const [mode, setMode] = useState<PdfLoadMode>("standard");
  const [compatibilityReason, setCompatibilityReason] = useState<PdfFailure | null>(null);
  const [failure, setFailure] = useState<PdfFailure | null>(null);
  const [fromCache, setFromCache] = useState(false);
  const [renderError, setRenderError] = useState<PdfFailure | null>(null);
  const [hitIndex, setHitIndex] = useState(0);
  const [centerRevision, setCenterRevision] = useState(0);

  useEffect(() => {
    const entry = getIndexEntry(file);
    const reused = entry.started && (entry.status === "ready" || entry.status === "no-text");
    const update = () => {
      setDocument(entry.document);
      setReferenceIndex(new Map(entry.referenceIndex));
      setProcessedPages(entry.processedPages);
      setTotalPages(entry.totalPages);
      setBatchSize(entry.batchSize);
      setStatus(entry.status);
      setMode(entry.mode);
      setCompatibilityReason(entry.compatibilityReason);
      setFailure(entry.failure);
    };
    entry.listeners.add(update);
    setFromCache(reused);
    setRenderError(null);
    setPageNumber(1);
    update();
    void startProgressiveIndex(file, entry);
    return () => {
      entry.listeners.delete(update);
      entry.lastUsed = Date.now();
      trimPdfCache();
      renderTask.current?.cancel();
      if (scrollFrame.current != null) window.cancelAnimationFrame(scrollFrame.current);
      if (locateFrame.current != null) window.cancelAnimationFrame(locateFrame.current);
    };
  }, [file]);

  const normalizedTarget = normalizeReference(target);
  const hits = useMemo(() => lookupReferenceHits(referenceIndex, normalizedTarget), [referenceIndex, normalizedTarget]);
  const currentHit = hits[hitIndex];

  const centerCurrentHit = useCallback(() => {
    const element = scrollRef.current;
    const pageElement = pageRef.current;
    if (!element || !pageElement || currentHit?.page !== pageNumber) return;
    const marker = highlightRef.current;
    const viewportRect = element.getBoundingClientRect();
    const markerRect = marker?.getBoundingClientRect();
    const centered = marker ? centeredRenderedHitScroll(
      { left: element.scrollLeft, top: element.scrollTop },
      { left: viewportRect.left, top: viewportRect.top, width: element.clientWidth, height: element.clientHeight },
      { left: markerRect!.left, top: markerRect!.top, width: markerRect!.width, height: markerRect!.height },
      { left: Math.max(0, element.scrollWidth - element.clientWidth), top: Math.max(0, element.scrollHeight - element.clientHeight) },
    ) : centeredPdfHitScroll(
      currentHit, scale, element.clientWidth, element.clientHeight, element.scrollWidth, element.scrollHeight,
      pageElement.offsetLeft, pageElement.offsetTop,
    );
    applyingSyncedScroll.current = true;
    element.scrollLeft = centered.left;
    element.scrollTop = centered.top;
    lastScrollRatio.current = {
      x: centered.left / Math.max(1, element.scrollWidth - element.clientWidth),
      y: centered.top / Math.max(1, element.scrollHeight - element.clientHeight),
    };
    window.requestAnimationFrame(() => { applyingSyncedScroll.current = false; });
  }, [currentHit, pageNumber, scale]);

  const scheduleCenterCurrentHit = useCallback(() => {
    if (locateFrame.current != null) window.cancelAnimationFrame(locateFrame.current);
    locateFrame.current = window.requestAnimationFrame(() => {
      locateFrame.current = window.requestAnimationFrame(() => {
        locateFrame.current = null;
        centerCurrentHit();
      });
    });
  }, [centerCurrentHit]);

  useEffect(() => {
    locatedTarget.current = "";
    setHitIndex(0);
  }, [normalizedTarget]);

  useEffect(() => {
    if (normalizedTarget && hits[0] && locatedTarget.current !== normalizedTarget) {
      locatedTarget.current = normalizedTarget;
      setPageNumber(hits[0].page);
    }
  }, [hits, normalizedTarget]);

  useEffect(() => {
    if (!document) return;
    let page: PDFPageProxy | null = null;
    document.getPage(pageNumber).then((loadedPage) => {
      page = loadedPage;
      const viewport = page.getViewport({ scale });
      const canvas = canvasRef.current;
      const context = canvas?.getContext("2d");
      if (!canvas || !context) return;
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      canvas.style.width = `${viewport.width}px`;
      canvas.style.height = `${viewport.height}px`;
      const pageElement = pageRef.current;
      if (pageElement) {
        pageElement.style.width = `${viewport.width}px`;
        pageElement.style.height = `${viewport.height}px`;
      }
      setPageSize({ width: viewport.width, height: viewport.height });
      if (currentHit?.page === pageNumber) {
        scheduleCenterCurrentHit();
      } else window.requestAnimationFrame(() => {
        const element = scrollRef.current;
        if (!element) return;
        element.scrollLeft = lastScrollRatio.current.x * Math.max(0, element.scrollWidth - element.clientWidth);
        element.scrollTop = lastScrollRatio.current.y * Math.max(0, element.scrollHeight - element.clientHeight);
      });
      renderTask.current?.cancel();
      const task = page.render({ canvas, canvasContext: context, viewport });
      renderTask.current = task;
      task.promise.then(() => {
        if (currentHit?.page === pageNumber) scheduleCenterCurrentHit();
      }).catch((error) => { if (error?.name !== "RenderingCancelledException") setRenderError(classifyPdfFailure(error)); });
    }).catch((error) => setRenderError(classifyPdfFailure(error)));
    return () => renderTask.current?.cancel();
  }, [currentHit, document, pageNumber, scale, scheduleCenterCurrentHit]);

  useEffect(() => {
    const element = scrollRef.current;
    if (!syncEnabled || !element || syncState.source === side) return;
    applyingSyncedScroll.current = true;
    element.scrollLeft = syncState.x * Math.max(0, element.scrollWidth - element.clientWidth);
    element.scrollTop = syncState.y * Math.max(0, element.scrollHeight - element.clientHeight);
    lastScrollRatio.current = { x: syncState.x, y: syncState.y };
    window.requestAnimationFrame(() => { applyingSyncedScroll.current = false; });
  }, [side, syncEnabled, syncState]);

  const handleScroll = () => {
    if (!syncEnabled || applyingSyncedScroll.current) return;
    if (scrollFrame.current != null) window.cancelAnimationFrame(scrollFrame.current);
    scrollFrame.current = window.requestAnimationFrame(() => {
      const element = scrollRef.current;
      if (!element) return;
      const x = element.scrollLeft / Math.max(1, element.scrollWidth - element.clientWidth);
      const y = element.scrollTop / Math.max(1, element.scrollHeight - element.clientHeight);
      lastScrollRatio.current = { x, y };
      onSyncScroll(side, x, y);
    });
  };

  const selectHit = (nextIndex: number) => {
    const normalized = (nextIndex + hits.length) % hits.length;
    setHitIndex(normalized);
    setPageNumber(hits[normalized].page);
  };

  const recenterHit = () => {
    if (!currentHit) return;
    setCenterRevision((revision) => revision + 1);
    if (currentHit.page !== pageNumber) setPageNumber(currentHit.page);
    else scheduleCenterCurrentHit();
  };

  const visibleFailure = renderError ?? (status === "error" ? failure : null);
  const statusMessage = visibleFailure
    ? terminalFailureMessage(visibleFailure)
    : status === "indexing"
      ? mode === "compatibility"
        ? `PDF 相容模式・已處理 ${processedPages}／${totalPages || "…"} 頁・逐頁解析`
        : `已處理 ${processedPages}／${totalPages || "…"} 頁・每批平行 ${batchSize} 頁`
      : status === "no-text"
        ? mode === "compatibility" ? `PDF 相容模式完成・共 ${totalPages} 頁，但沒有可搜尋文字` : "PDF 沒有可搜尋文字"
        : fromCache
          ? `已使用索引快取・共 ${totalPages} 頁`
          : mode === "compatibility" ? `PDF 相容模式完成・共 ${totalPages} 頁` : `索引表完成・共 ${totalPages} 頁`;

  return <div className="pdf-viewer">
    <div className={`pdf-status${visibleFailure ? " error" : mode === "compatibility" ? " compatibility" : ""}`}>
      <span>{statusMessage}</span>
      {mode === "compatibility" && compatibilityReason && !visibleFailure && <small>{compatibilityReason.message}</small>}
      {status === "indexing" && totalPages > 0 && <span className="pdf-progress"><i style={{ width: `${Math.round(processedPages / totalPages * 100)}%` }} /></span>}
      {normalizedTarget && !visibleFailure && <strong className={hits.length ? "found" : status === "ready" || status === "no-text" ? "not-found" : "searching"}><FileSearch size={13} /> {hits.length ? `${sideLabel}找到 ${hits.length} 筆 ${normalizedTarget}${status === "indexing" ? "・背景索引中" : ""}` : status === "indexing" ? `優先搜尋 ${normalizedTarget}…` : `${sideLabel}找不到 ${normalizedTarget}`}</strong>}
    </div>
    <div className="pdf-page-controls">
      <button disabled={!!visibleFailure || pageNumber <= 1} onClick={() => setPageNumber((page) => page - 1)}><ChevronLeft size={15} /></button><span>{visibleFailure ? "PDF 無法載入" : `第 ${pageNumber}／${document?.numPages ?? 0} 頁`}</span><button disabled={!!visibleFailure || !document || pageNumber >= document.numPages} onClick={() => setPageNumber((page) => page + 1)}><ChevronRight size={15} /></button>
      <div className="zoom-controls"><button disabled={scale <= 0.6} onClick={() => onScaleChange(side, Math.max(0.6, scale - 0.15))}>−</button><span>{Math.round(scale * 100)}%</span><button disabled={scale >= 2.4} onClick={() => onScaleChange(side, Math.min(2.4, scale + 0.15))}>＋</button></div>
      {currentHit && <button className="recenter-hit" onClick={recenterHit}>重新置中 {normalizedTarget}</button>}
      {hits.length > 1 && <div className="hit-controls"><button onClick={() => selectHit(hitIndex - 1)}>上一筆</button><span>{hitIndex + 1}／{hits.length}</span><button onClick={() => selectHit(hitIndex + 1)}>下一筆</button></div>}
    </div>
    <div className="pdf-page-scroll" ref={scrollRef} onScroll={handleScroll}>
      {visibleFailure ? <div className="pdf-error-state"><strong>{visibleFailure.message}</strong><span>{visibleFailure.kind === "encrypted" ? "請使用 Excel／PDF 工具另存成未加密 PDF。" : visibleFailure.kind === "invalid" ? "請確認檔案可以在 Edge 開啟，或由線路圖工具重新輸出。" : visibleFailure.kind === "browser-unsupported" ? "建議更新 Microsoft Edge 或 Chrome 後重新啟動 BOMLens。" : visibleFailure.kind === "worker-blocked" ? "請將 BOMLens 與 127.0.0.1 加入公司端點防護允許清單。" : "請改用另一份 PDF 測試；若仍失敗，請將此訊息提供給維護人員。"}</span></div> : <div className="pdf-page" ref={pageRef} style={{ width: pageSize.width, height: pageSize.height }}>
          <canvas ref={canvasRef} />
          {currentHit?.page === pageNumber && <span ref={highlightRef} key={`${normalizedTarget}-${hitIndex}-${centerRevision}`} className="pdf-highlight" style={{ left: Math.max(0, (currentHit.x - 5) * scale / PDF_REFERENCE_INDEX_SCALE), top: Math.max(0, (currentHit.y - 4) * scale / PDF_REFERENCE_INDEX_SCALE), width: (currentHit.width + 10) * scale / PDF_REFERENCE_INDEX_SCALE, height: (currentHit.height + 8) * scale / PDF_REFERENCE_INDEX_SCALE }} />}
        </div>}
    </div>
  </div>;
}
