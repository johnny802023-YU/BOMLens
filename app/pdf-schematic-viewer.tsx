"use client";
/* eslint-disable react-hooks/set-state-in-effect -- file and search changes intentionally reset the local PDF viewer state */

import { ChevronLeft, ChevronRight, FileSearch } from "lucide-react";
import { useEffect, useMemo, useRef, useState } from "react";
import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import type { PDFDocumentProxy, PDFPageProxy } from "pdfjs-dist";
import { buildPageReferenceIndex, centeredPdfHitScroll, lookupReferenceHits, mergeReferenceIndex, normalizeReference, PDF_REFERENCE_INDEX_SCALE, type PdfReferenceIndex, type PdfTextBox } from "./pdf-search";

type IndexStatus = "indexing" | "ready" | "no-text" | "error";
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
  started: boolean;
  cancelled: boolean;
  lastUsed: number;
  listeners: Set<() => void>;
};

const MAX_CACHED_PDFS = 2;
const pdfIndexCache = new Map<File, IndexEntry>();

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

function getIndexEntry(file: File) {
  const existing = pdfIndexCache.get(file);
  if (existing) { existing.lastUsed = Date.now(); return existing; }
  const entry: IndexEntry = { document: null, referenceIndex: new Map(), textItemCount: 0, processedPages: 0, totalPages: 0, batchSize: 2, status: "indexing", started: false, cancelled: false, lastUsed: Date.now(), listeners: new Set() };
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
  try {
    const data = await file.arrayBuffer();
    const pdfjs = await import("pdfjs-dist");
    pdfjs.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;
    entry.document = await pdfjs.getDocument({ data }).promise;
    if (entry.cancelled) { void entry.document.destroy(); return; }
    entry.totalPages = entry.document.numPages;
    entry.batchSize = Math.min(4, Math.max(2, Math.floor((navigator.hardwareConcurrency || 4) / 2)));
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
  } catch {
    if (!entry.cancelled) { entry.status = "error"; notify(entry); }
  }
}

export function PdfSchematicViewer({ file, target, sideLabel, side, scale, syncEnabled, syncState, onScaleChange, onSyncScroll }: { file: File; target: string; sideLabel: string; side: PdfViewerSide; scale: number; syncEnabled: boolean; syncState: PdfScrollSync; onScaleChange: (side: PdfViewerSide, scale: number) => void; onSyncScroll: (side: PdfViewerSide, x: number, y: number) => void }) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const pageRef = useRef<HTMLDivElement>(null);
  const scrollRef = useRef<HTMLDivElement>(null);
  const renderTask = useRef<{ cancel(): void } | null>(null);
  const applyingSyncedScroll = useRef(false);
  const scrollFrame = useRef<number | null>(null);
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
  const [fromCache, setFromCache] = useState(false);
  const [renderError, setRenderError] = useState(false);
  const [hitIndex, setHitIndex] = useState(0);

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
    };
    entry.listeners.add(update);
    setFromCache(reused);
    setRenderError(false);
    setPageNumber(1);
    update();
    void startProgressiveIndex(file, entry);
    return () => {
      entry.listeners.delete(update);
      entry.lastUsed = Date.now();
      trimPdfCache();
      renderTask.current?.cancel();
      if (scrollFrame.current != null) window.cancelAnimationFrame(scrollFrame.current);
    };
  }, [file]);

  const normalizedTarget = normalizeReference(target);
  const hits = useMemo(() => lookupReferenceHits(referenceIndex, normalizedTarget), [referenceIndex, normalizedTarget]);
  const currentHit = hits[hitIndex];

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
      setPageSize({ width: viewport.width, height: viewport.height });
      window.requestAnimationFrame(() => {
        const element = scrollRef.current;
        if (!element) return;
        if (currentHit?.page === pageNumber) {
          const pageElement = pageRef.current;
          const centered = centeredPdfHitScroll(
            currentHit,
            scale,
            element.clientWidth,
            element.clientHeight,
            element.scrollWidth,
            element.scrollHeight,
            pageElement?.offsetLeft ?? 0,
            pageElement?.offsetTop ?? 0,
          );
          applyingSyncedScroll.current = true;
          element.scrollLeft = centered.left;
          element.scrollTop = centered.top;
          lastScrollRatio.current = {
            x: centered.left / Math.max(1, element.scrollWidth - element.clientWidth),
            y: centered.top / Math.max(1, element.scrollHeight - element.clientHeight),
          };
          window.requestAnimationFrame(() => { applyingSyncedScroll.current = false; });
        } else {
          element.scrollLeft = lastScrollRatio.current.x * Math.max(0, element.scrollWidth - element.clientWidth);
          element.scrollTop = lastScrollRatio.current.y * Math.max(0, element.scrollHeight - element.clientHeight);
        }
      });
      renderTask.current?.cancel();
      const task = page.render({ canvas, canvasContext: context, viewport });
      renderTask.current = task;
      task.promise.catch((error) => { if (error?.name !== "RenderingCancelledException") setRenderError(true); });
    });
    return () => renderTask.current?.cancel();
  }, [currentHit, document, pageNumber, scale]);

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

  return <div className="pdf-viewer">
    <div className="pdf-status">
      <span>{renderError || status === "error" ? "PDF 讀取失敗" : status === "indexing" ? `已處理 ${processedPages}／${totalPages || "…"} 頁・每批平行 ${batchSize} 頁` : status === "no-text" ? "PDF 沒有可搜尋文字" : fromCache ? `已使用索引快取・共 ${totalPages} 頁` : `索引表完成・共 ${totalPages} 頁`}</span>
      {status === "indexing" && totalPages > 0 && <span className="pdf-progress"><i style={{ width: `${Math.round(processedPages / totalPages * 100)}%` }} /></span>}
      {normalizedTarget && status !== "error" && <strong className={hits.length ? "found" : status === "ready" || status === "no-text" ? "not-found" : "searching"}><FileSearch size={13} /> {hits.length ? `${sideLabel}找到 ${hits.length} 筆 ${normalizedTarget}${status === "indexing" ? "・背景索引中" : ""}` : status === "indexing" ? `優先搜尋 ${normalizedTarget}…` : `${sideLabel}找不到 ${normalizedTarget}`}</strong>}
    </div>
    <div className="pdf-page-controls">
      <button disabled={pageNumber <= 1} onClick={() => setPageNumber((page) => page - 1)}><ChevronLeft size={15} /></button><span>第 {pageNumber}／{document?.numPages ?? 0} 頁</span><button disabled={!document || pageNumber >= document.numPages} onClick={() => setPageNumber((page) => page + 1)}><ChevronRight size={15} /></button>
      <div className="zoom-controls"><button disabled={scale <= 0.6} onClick={() => onScaleChange(side, Math.max(0.6, scale - 0.15))}>−</button><span>{Math.round(scale * 100)}%</span><button disabled={scale >= 2.4} onClick={() => onScaleChange(side, Math.min(2.4, scale + 0.15))}>＋</button></div>
      {hits.length > 1 && <div className="hit-controls"><button onClick={() => selectHit(hitIndex - 1)}>上一筆</button><span>{hitIndex + 1}／{hits.length}</span><button onClick={() => selectHit(hitIndex + 1)}>下一筆</button></div>}
    </div>
    <div className="pdf-page-scroll" ref={scrollRef} onScroll={handleScroll}>
      <div className="pdf-page" ref={pageRef} style={{ width: pageSize.width, height: pageSize.height }}>
        <canvas ref={canvasRef} />
        {currentHit?.page === pageNumber && <span className="pdf-highlight" style={{ left: Math.max(0, (currentHit.x - 5) * scale / PDF_REFERENCE_INDEX_SCALE), top: Math.max(0, (currentHit.y - 4) * scale / PDF_REFERENCE_INDEX_SCALE), width: (currentHit.width + 10) * scale / PDF_REFERENCE_INDEX_SCALE, height: (currentHit.height + 8) * scale / PDF_REFERENCE_INDEX_SCALE }} />}
      </div>
    </div>
  </div>;
}
