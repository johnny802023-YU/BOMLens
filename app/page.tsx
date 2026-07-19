"use client";
/* eslint-disable @next/next/no-img-element -- previews use local object URLs and canvas output */

import {
  AlertTriangle,
  ArrowRight,
  CircuitBoard,
  Download,
  FileSpreadsheet,
  FileText,
  Filter,
  GitCompareArrows,
  History,
  Image as ImageIcon,
  Menu,
  PanelLeftClose,
  PanelLeftOpen,
  Plus,
  Search,
  ShieldCheck,
  UploadCloud,
  X,
} from "lucide-react";
import { ChangeEvent, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { clearInactivePdfCache, PdfSchematicViewer, type PdfScrollSync, type PdfViewerSide } from "./pdf-schematic-viewer";
import {
  canonicalPartNumber,
  bomStructureLabel,
  compareBom,
  analyzeCompanyBomMatrix,
  companyColumnLabels,
  detectCompanyColumns,
  findCompanyHeader,
  isCompleteCompanyMapping,
  requiredCompanyColumns,
  sortBomDiffsForAll,
  type BomAlternative,
  type BomDiff,
  type BomItem,
  type CompanyColumnKey,
  type CompanyColumnMapping,
  type DiffPrimaryType,
  type ImportAudit,
} from "./bom-logic";

type DiffFilter = "all" | Exclude<DiffPrimaryType, "same">;
type ImpactFilter = "all" | "purchase" | "deleted" | "review";

type ImportSheet = { name: string; matrix: unknown[][] };
type PendingImport = {
  side: "before" | "after";
  fileName: string;
  sheets: ImportSheet[];
  sheetIndex: number;
  headerIndex: number;
  mapping: CompanyColumnMapping;
};
type ImportRecord = { fileName: string; sheetName: string; importedAt: string; audit: ImportAudit };

const primaryFilterOptions: Array<{ key: DiffFilter; label: string }> = [
  { key: "all", label: "全部" },
  { key: "componentAdded", label: "新增元件" },
  { key: "substituteAdded", label: "新增替料" },
  { key: "componentRemoved", label: "移除元件" },
  { key: "substituteRemoved", label: "刪除替料" },
  { key: "partReplaced", label: "同位置換料" },
  { key: "positionChanged", label: "位置變更" },
];

function bomItem(ref: string, alternatives: Array<[string, string, string?]>, positions: string[], description = ""): BomItem {
  const parts = alternatives.map(([part, manufacturerPart, manufacturerName = ""]) => ({ part: canonicalPartNumber(part), manufacturerPart, manufacturerName, description, spec: "" }));
  return {
    ref,
    part: parts[0]?.part ?? "",
    manufacturerPart: parts[0]?.manufacturerPart ?? "",
    manufacturerName: parts[0]?.manufacturerName ?? "",
    value: "",
    description,
    qty: positions.length,
    positions,
    alternatives: parts,
  };
}

const demoBefore: BomItem[] = [
  bomItem("00A", [["I-0500-05P9OVD", "S25FL127SABNFM703"], ["I-S-0603-0207V", "SN74LVC1G14QDCKRQ1"]], ["U20"], "SPI FLASH / LOGIC"),
  bomItem("00K", [["I-0603-0207V", "SN74LVC1G14QDCKRQ1"]], ["U31"], "LOGIC"),
  bomItem("014", [["I-0603-02RD0VD", "74LVC1G32GW-Q100H"]], ["U33", "U38", "U39"], "LOGIC"),
  bomItem("018", [["I-0603-02960YD", "SN74LVC1G32QDCKRQ1"]], ["U29", "U32", "U49", "U52"], "LOGIC"),
  bomItem("03C", [["I-0618-00N30VD", "NCV301LSN28T1G"], ["I-S-0618-ALT01", "NCV301LSN28T1G-ALT"]], ["U17", "U37"], "VOLT DETEC."),
  bomItem("03D", [["OLD-PART-A001", "OLD-MPN-A"]], ["U45"], "SAME POSITION REPLACEMENT"),
  bomItem("040", [["I-0628-02200VD", "BD900N1WEFJ-CE2"]], ["U11"], "LDO REG."),
];

const demoAfter: BomItem[] = [
  // 00A 的主料／替料刻意交換順序，應判定為相同。
  bomItem("00A", [["I-S-0603-0207V", "SN74LVC1G14QDCKRQ1"], ["I-0500-05P9OVD", "S25FL127SABNFM703"]], ["U20"], "SPI FLASH / LOGIC"),
  bomItem("00K", [["I-0603-0207V", "SN74LVC1G14QDCKRQ1"], ["I-S-0603-0207V", "SN74LVC1G14QDCKRQ1"]], ["U31"], "LOGIC"),
  bomItem("014", [["I-0603-02RD0VD", "74LVC1G32GW-Q100H"]], ["U33", "U38", "U39", "U41"], "LOGIC"),
  bomItem("018", [["I-0603-02960YD", "SN74LVC1G32QDCKRQ1"]], ["U29", "U32", "U49", "U53"], "LOGIC"),
  bomItem("03C", [["I-0618-00N30VD", "NCV301LSN28T1G"]], ["U17", "U37"], "VOLT DETEC."),
  bomItem("03D", [["NEW-PART-B001", "NEW-MPN-B"]], ["U45"], "SAME POSITION REPLACEMENT"),
  bomItem("05A", [["I-0628-036A02Y", "LDH40PURQY"]], ["U13", "U102A", "U102B", "U102C"], "LDO REG."),
];

function displayPart(item?: BomItem) {
  if (!item) return "—";
  return item.alternatives.map((alternative) => alternative.part || "未提供料號").join(" / ") || "未提供料號";
}

function alternativeLabel(alternative: BomAlternative) {
  return alternative.part || alternative.manufacturerPart || "未提供料號";
}

function primaryTypeLabel(type: DiffPrimaryType) {
  return primaryFilterOptions.find((option) => option.key === type)?.label ?? "相同";
}

function primaryTypeTone(type: DiffPrimaryType) {
  if (type === "componentAdded" || type === "substituteAdded") return "added";
  if (type === "componentRemoved" || type === "substituteRemoved") return "removed";
  if (type === "partReplaced") return "replacement";
  return "changed";
}

function diffStructureLabel(item: BomDiff) {
  const before = bomStructureLabel(item.before);
  const after = bomStructureLabel(item.after);
  if (before && after && before !== after) return `${before} → ${after}`;
  return after || before;
}

export default function Home() {
  const [tab, setTab] = useState<"bom" | "schematic">("bom");
  const [before, setBefore] = useState(demoBefore);
  const [after, setAfter] = useState(demoAfter);
  const [beforeName, setBeforeName] = useState("PCB_Main_v1.3.xlsx");
  const [afterName, setAfterName] = useState("PCB_Main_v1.4.xlsx");
  const [query, setQuery] = useState("");
  const [filter, setFilter] = useState<DiffFilter>("all");
  const [impactFilter, setImpactFilter] = useState<ImpactFilter>("all");
  const [sheetBefore, setSheetBefore] = useState<string | null>(null);
  const [sheetAfter, setSheetAfter] = useState<string | null>(null);
  const [sheetBeforeFile, setSheetBeforeFile] = useState<File | null>(null);
  const [sheetAfterFile, setSheetAfterFile] = useState<File | null>(null);
  const [sheetBeforeName, setSheetBeforeName] = useState("");
  const [sheetAfterName, setSheetAfterName] = useState("");
  const [diffImage, setDiffImage] = useState<string | null>(null);
  const [imageBusy, setImageBusy] = useState(false);
  const [mobileNav, setMobileNav] = useState(false);
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [pendingImport, setPendingImport] = useState<PendingImport | null>(null);
  const [beforeAudit, setBeforeAudit] = useState<ImportRecord | null>(null);
  const [afterAudit, setAfterAudit] = useState<ImportRecord | null>(null);
  const [sessionOpen, setSessionOpen] = useState(false);
  const [sessionRecords, setSessionRecords] = useState<ImportRecord[]>([]);
  const [schematicTarget, setSchematicTarget] = useState("");
  const beforeInput = useRef<HTMLInputElement>(null);
  const afterInput = useRef<HTMLInputElement>(null);

  const diffs = useMemo(() => compareBom(before, after), [before, after]);
  const changedDiffs = diffs.filter((item) => item.categories.length > 0);
  const summary = useMemo(
    () => ({
      added: new Set(diffs.flatMap((d) => d.newParts.map((part) => part.part))).size,
      removed: new Set(diffs.flatMap((d) => d.deletedParts.map((part) => part.part))).size,
      changed: diffs.filter((d) => d.categories.includes("changed")).length,
      same: diffs.filter((d) => d.categories.length === 0).length,
    }),
    [diffs],
  );
  const matchingDiffs = changedDiffs.filter((item) => {
    const manufacturerNames = [...(item.before?.alternatives ?? []), ...(item.after?.alternatives ?? [])].map((part) => part.manufacturerName ?? "").join(" ");
    const text = `${item.ref} ${diffStructureLabel(item)} ${displayPart(item.before)} ${displayPart(item.after)} ${item.before?.manufacturerPart ?? ""} ${item.after?.manufacturerPart ?? ""} ${manufacturerNames} ${item.fields.join(" ")} ${item.addedPositions.join(" ")} ${item.removedPositions.join(" ")}`.toLowerCase();
    const matchesFilter = filter === "all" || item.primaryType === filter;
    const matchesImpact = impactFilter === "all"
      || (impactFilter === "purchase" && item.newParts.length > 0)
      || (impactFilter === "deleted" && item.deletedParts.length > 0)
      || (impactFilter === "review" && item.needsReview);
    return matchesFilter && matchesImpact && text.includes(query.toLowerCase());
  });
  const visible = filter === "all" ? sortBomDiffsForAll(matchingDiffs) : matchingDiffs;

  async function loadBom(file: File, side: "before" | "after") {
    const data = await file.arrayBuffer();
    const book = XLSX.read(data, { type: "array" });
    const sheets = book.SheetNames.map((name) => ({ name, matrix: XLSX.utils.sheet_to_json<unknown[]>(book.Sheets[name], { header: 1, defval: "", raw: false }) }));
    const candidateIndex = sheets.findIndex((sheet) => findCompanyHeader(sheet.matrix) >= 0);
    const sheetIndex = candidateIndex >= 0 ? candidateIndex : 0;
    const matrix = sheets[sheetIndex]?.matrix ?? [];
    const detectedHeader = findCompanyHeader(matrix);
    const headerIndex = detectedHeader >= 0 ? detectedHeader : Math.max(0, matrix.findIndex((row) => row.filter((cell) => String(cell ?? "").trim()).length >= 2));
    setPendingImport({ side, fileName: file.name, sheets, sheetIndex, headerIndex, mapping: detectCompanyColumns(matrix[headerIndex] ?? []) });
  }

  function confirmImport() {
    if (!pendingImport) return;
    const sheet = pendingImport.sheets[pendingImport.sheetIndex];
    const result = analyzeCompanyBomMatrix(sheet.matrix, pendingImport.headerIndex, pendingImport.mapping);
    if (!isCompleteCompanyMapping(pendingImport.mapping) || result.audit.issues.some((issue) => issue.severity === "error")) return;
    const items = result.items.map((item) => ({ ...item, sourceSheet: sheet.name }));
    const record = { fileName: pendingImport.fileName, sheetName: sheet.name, importedAt: new Date().toLocaleString("zh-TW", { hour12: false }), audit: result.audit };
    if (pendingImport.side === "before") {
      setBefore(items); setBeforeName(pendingImport.fileName); setBeforeAudit(record);
    } else {
      setAfter(items); setAfterName(pendingImport.fileName); setAfterAudit(record);
    }
    setSessionRecords((records) => [record, ...records].slice(0, 20));
    setPendingImport(null);
  }

  function resetComparison() {
    setBefore([]); setAfter([]); setBeforeName(""); setAfterName(""); setBeforeAudit(null); setAfterAudit(null);
    setQuery(""); setFilter("all"); setImpactFilter("all"); setTab("bom");
    window.setTimeout(() => beforeInput.current?.click(), 0);
  }

  function handleBomFile(event: ChangeEvent<HTMLInputElement>, side: "before" | "after") {
    const file = event.target.files?.[0];
    if (file) loadBom(file, side).catch(() => window.alert("檔案讀取失敗，請改用 XLSX、XLS 或 CSV 格式。"));
  }

  function handleSheetFile(event: ChangeEvent<HTMLInputElement>, side: "before" | "after") {
    const file = event.target.files?.[0];
    if (!file) return;
    const url = URL.createObjectURL(file);
    setDiffImage(null);
    const fileIsPdf = file.name.toLowerCase().endsWith(".pdf");
    const otherReady = side === "before" ? Boolean(sheetAfter) : Boolean(sheetBefore);
    const otherIsPdf = (side === "before" ? sheetAfterName : sheetBeforeName).toLowerCase().endsWith(".pdf");
    setImageBusy(otherReady && !fileIsPdf && !otherIsPdf);
    if (side === "before") {
      if (sheetBefore) URL.revokeObjectURL(sheetBefore);
      setSheetBefore(url);
      setSheetBeforeFile(file);
      setSheetBeforeName(file.name);
    } else {
      if (sheetAfter) URL.revokeObjectURL(sheetAfter);
      setSheetAfter(url);
      setSheetAfterFile(file);
      setSheetAfterName(file.name);
    }
  }

  useEffect(() => {
    if (!sheetBefore || !sheetAfter) {
      return;
    }
    const leftPdf = sheetBeforeName.toLowerCase().endsWith(".pdf");
    const rightPdf = sheetAfterName.toLowerCase().endsWith(".pdf");
    if (leftPdf || rightPdf) {
      return;
    }
    let cancelled = false;
    const load = (src: string) => new Promise<HTMLImageElement>((resolve, reject) => {
      const image = new Image();
      image.onload = () => resolve(image);
      image.onerror = reject;
      image.src = src;
    });
    Promise.all([load(sheetBefore), load(sheetAfter)]).then(([a, b]) => {
      if (cancelled) return;
      const maxWidth = 1400;
      const width = Math.min(maxWidth, Math.max(a.naturalWidth, b.naturalWidth));
      const ratio = Math.max(a.naturalWidth, b.naturalWidth) / width;
      const height = Math.round(Math.max(a.naturalHeight, b.naturalHeight) / ratio);
      const c1 = document.createElement("canvas");
      const c2 = document.createElement("canvas");
      const output = document.createElement("canvas");
      [c1, c2, output].forEach((canvas) => { canvas.width = width; canvas.height = height; });
      const x1 = c1.getContext("2d", { willReadFrequently: true })!;
      const x2 = c2.getContext("2d", { willReadFrequently: true })!;
      const xo = output.getContext("2d")!;
      x1.fillStyle = x2.fillStyle = "white";
      x1.fillRect(0, 0, width, height); x2.fillRect(0, 0, width, height);
      x1.drawImage(a, 0, 0, width, height); x2.drawImage(b, 0, 0, width, height);
      const p1 = x1.getImageData(0, 0, width, height);
      const p2 = x2.getImageData(0, 0, width, height);
      const result = xo.createImageData(width, height);
      for (let i = 0; i < p1.data.length; i += 4) {
        const delta = Math.abs(p1.data[i] - p2.data[i]) + Math.abs(p1.data[i + 1] - p2.data[i + 1]) + Math.abs(p1.data[i + 2] - p2.data[i + 2]);
        if (delta > 70) {
          result.data[i] = 225; result.data[i + 1] = 53; result.data[i + 2] = 53; result.data[i + 3] = Math.min(255, 90 + delta / 2);
        } else {
          const gray = Math.round((p2.data[i] + p2.data[i + 1] + p2.data[i + 2]) / 3);
          result.data[i] = result.data[i + 1] = result.data[i + 2] = Math.min(248, 210 + gray * 0.16);
          result.data[i + 3] = 255;
        }
      }
      xo.putImageData(result, 0, 0);
      setDiffImage(output.toDataURL("image/png"));
    }).catch(() => setDiffImage(null)).finally(() => setImageBusy(false));
    return () => { cancelled = true; };
  }, [sheetBefore, sheetAfter, sheetBeforeName, sheetAfterName]);

  async function exportCsv() {
    try {
      const { exportBomReport } = await import("./export-report");
      await exportBomReport(changedDiffs, beforeName, afterName, { before: beforeAudit, after: afterAudit });
    } catch {
      window.alert("報表產生失敗，請重新整理後再試一次。");
    }
  }

  async function exportHtml() {
    try {
      const { exportBomHtmlReport } = await import("./export-html-report");
      exportBomHtmlReport(changedDiffs, beforeName, afterName, { before: beforeAudit, after: afterAudit });
    } catch {
      window.alert("HTML 報告產生失敗，請重新整理後再試一次。");
    }
  }

  return (
    <main className="app-shell">
      <aside className={`sidebar ${sidebarCollapsed ? "collapsed" : ""} ${mobileNav ? "open" : ""}`}>
        <button className="sidebar-collapse" onClick={() => setSidebarCollapsed((collapsed) => !collapsed)} aria-label={sidebarCollapsed ? "展開側邊選單" : "收折側邊選單"} aria-expanded={!sidebarCollapsed} title={sidebarCollapsed ? "展開側邊選單" : "收折側邊選單"}>{sidebarCollapsed ? <PanelLeftOpen size={15} /> : <PanelLeftClose size={15} />}</button>
        <div className="brand"><span className="brand-mark"><CircuitBoard size={21} /></span><span className="brand-text">BOM<span>Lens</span></span></div>
        <button className="new-compare" title="新增比對" onClick={resetComparison}><Plus size={18} /><span className="sidebar-text">新增比對</span></button>
        <nav>
          <p className="nav-label">工作區</p>
          <button className="nav-item active" title="目前比對"><GitCompareArrows size={18} /><span className="sidebar-text">目前比對</span><span className="nav-count">1</span></button>
          <button className="nav-item" title="此次工作階段" onClick={() => setSessionOpen(true)}><History size={18} /><span className="sidebar-text">此次工作階段</span><span className="nav-count">{sessionRecords.length}</span></button>
          <p className="nav-label recent-label">資料保護</p>
          <div className="privacy-nav"><ShieldCheck size={17} /><div><strong>完全本機處理</strong><small>不登入・不上傳・不留存</small></div></div>
        </nav>
        <div className="sidebar-footer"><div className="avatar"><ShieldCheck size={16} /></div><div className="sidebar-footer-text"><strong>離線工作階段</strong><small>127.0.0.1 本機限定</small></div></div>
      </aside>
      {mobileNav && <button className="nav-scrim" aria-label="關閉選單" onClick={() => setMobileNav(false)} />}

      <section className={`workspace ${sidebarCollapsed ? "sidebar-collapsed" : ""}`}>
        <header className="topbar">
          <button className="mobile-menu" onClick={() => setMobileNav(true)} aria-label="開啟選單"><Menu /></button>
          <div><div className="breadcrumb">本機工具 <span>/</span> BOM 版本比對</div><h1>版本比對 <span className="version-pill">{beforeName || "前版"} → {afterName || "後版"}</span></h1></div>
          <div className="top-actions"><span className="saved offline"><ShieldCheck size={14} /> 本機離線</span><button className="secondary" onClick={exportHtml}><FileText size={17} /> 匯出 HTML</button><button className="primary" onClick={exportCsv}><Download size={17} /> 匯出差異</button></div>
        </header>

        <div className="content">
          <section className="privacy-banner" role="status"><span><ShieldCheck size={19} /></span><div><strong>離線隱私模式</strong><p>檔案只在這台電腦的瀏覽器記憶體內分析，不會上傳、同步或儲存；關閉頁面後即清除。</p></div><span className="local-only">LOCAL ONLY</span></section>
          <section className="upload-bar">
            <div className="upload-title"><UploadCloud size={20} /><div><strong>比對來源</strong><small>選擇檔案後立即在本機完成分析</small></div></div>
            <div className="file-pair">
              <div className="file-chip"><button className="file-select" onClick={() => beforeInput.current?.click()}><span className="file-icon"><FileSpreadsheet size={18} /></span><span><small>前版 BOM</small><strong>{beforeName || "選擇檔案"}</strong></span></button>{beforeName && <button className="chip-x" aria-label="移除前版 BOM" onClick={() => { setBefore([]); setBeforeName(""); setBeforeAudit(null); }}><X size={15} /></button>}</div>
              <ArrowRight size={18} className="pair-arrow" />
              <div className="file-chip"><button className="file-select" onClick={() => afterInput.current?.click()}><span className="file-icon after"><FileSpreadsheet size={18} /></span><span><small>後版 BOM</small><strong>{afterName || "選擇檔案"}</strong></span></button>{afterName && <button className="chip-x" aria-label="移除後版 BOM" onClick={() => { setAfter([]); setAfterName(""); setAfterAudit(null); }}><X size={15} /></button>}</div>
              <input ref={beforeInput} hidden type="file" accept=".xlsx,.xls,.csv,.tsv" onChange={(e) => handleBomFile(e, "before")} />
              <input ref={afterInput} hidden type="file" accept=".xlsx,.xls,.csv,.tsv" onChange={(e) => handleBomFile(e, "after")} />
            </div>
          </section>
          {(beforeAudit || afterAudit) && <section className="audit-strip" aria-label="匯入檢查摘要">
            {[{ label: "前版", audit: beforeAudit }, { label: "後版", audit: afterAudit }].map(({ label, audit }) => audit && <div key={label}><ShieldCheck size={16} /><span><strong>{label}匯入完成</strong><small>{audit.sheetName}・{audit.audit.groupCount} 組料・{audit.audit.positionCount} 個位置・{audit.audit.issues.filter((issue) => issue.severity === "warning").length} 項警告</small></span></div>)}
          </section>}
          <section className="format-strip" aria-label="BOM 欄位規則">
            <span><b>項次</b> 只切分同架構主替料，不跨版比對</span><span><b>架構</b> 69 → VB-T／VB-D／08 PCB</span><span><b>料號</b> 取最右 12 碼</span><span><b>數量</b> 一般數量</span><span><b>插件位置</b> 優先計數</span><span><b>製造商名稱</b> 僅顯示</span><span><b>製造商料號</b> 僅顯示</span><span className="ignored"><b>客戶料號</b> 暫不比對</span><strong>依標題名稱自動定位欄位</strong>
          </section>

          <section className="summary-grid">
            <article className="summary-card total"><span className="summary-icon"><GitCompareArrows /></span><div><small>差異群組</small><strong>{changedDiffs.length}</strong><p>共配對 {diffs.length} 個料號／位置群組</p></div></article>
            <article className="summary-card added"><span className="summary-icon"><Plus /></span><div><small>新增料號</small><strong>{summary.added}</strong><p>舊版整份 BOM 未出現</p></div></article>
            <article className="summary-card removed"><span className="summary-icon"><X /></span><div><small>刪除料號</small><strong>{summary.removed}</strong><p>新版整份 BOM 已無使用</p></div></article>
            <article className="summary-card changed"><span className="summary-icon"><AlertTriangle /></span><div><small>變更群組</small><strong>{summary.changed}</strong><p>數量、位置或同位置換料</p></div></article>
          </section>

          <section className="panel">
            <div className="tabs"><button className={tab === "bom" ? "active" : ""} onClick={() => setTab("bom")}><FileSpreadsheet size={18} /> BOM 差異 <span>{changedDiffs.length}</span></button><button className={tab === "schematic" ? "active" : ""} onClick={() => setTab("schematic")}><CircuitBoard size={18} /> 線路圖比對</button></div>

            {tab === "bom" ? <>
              <div className="table-tools">
                <label className="search"><Search size={17} /><input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="搜尋料號、製造商或插件位置" /></label>
                <div className="filter-panel">
                  <div className="filters primary-filters"><span className="filter-title"><Filter size={14} />主要異動</span>{primaryFilterOptions.map((option) => <button key={option.key} onClick={() => setFilter(option.key)} className={filter === option.key ? "active" : ""}>{option.label}</button>)}</div>
                  <div className="filters impact-filters"><span className="filter-title">影響條件</span>{([{ key: "all", label: "不限" }, { key: "purchase", label: "新版完全新料" }, { key: "deleted", label: "新版完全移除" }, { key: "review", label: "待人工確認" }] as const).map((option) => <button key={option.key} onClick={() => setImpactFilter(option.key)} className={impactFilter === option.key ? "active" : ""}>{option.label}</button>)}</div>
                </div>
              </div>
              <div className="table-wrap">
                <table>
                  <thead><tr><th>主要異動</th><th>影響標籤</th><th>料號新增／刪除</th><th>插件位置差異</th><th>舊版料號／製造商資訊</th><th></th><th>新版料號／製造商資訊</th><th>數量</th></tr></thead>
                  <tbody>{visible.map((item, index) => <tr key={`${item.before?.structureKey ?? "none"}>${item.after?.structureKey ?? "none"}:${item.ref}:${index}`}>
                    <td><PrimaryTypeBadge item={item} /></td>
                    <td><ChangeFields fields={item.fields} /></td>
                    <td><BPartDifference item={item} /></td>
                    <td><PositionSummary added={item.addedPositions} removed={item.removedPositions} replacement={item.replacementPositions} allPositions={item.after?.positions ?? item.before?.positions ?? []} onLocate={(position) => { setSchematicTarget(position); setTab("schematic"); }} /></td>
                    <td><PartList item={item.before} changedParts={item.removedParts} tone="removed" /></td>
                    <td><ArrowRight size={16} className="row-arrow" /></td>
                    <td><PartList item={item.after} changedParts={item.addedParts} tone="added" /></td>
                    <td><span className={item.before?.qty !== item.after?.qty ? "qty changed-qty" : "qty"}>{item.before?.qty ?? 0} → {item.after?.qty ?? 0}</span></td>
                  </tr>)}</tbody>
                </table>
                {!visible.length && <div className="empty-state">沒有符合條件的差異</div>}
              </div>
              <div className="table-footer"><span>顯示 {visible.length} 筆，共 {diffs.length} 個料號／位置群組；項次名稱不列入差異</span><span className="legend"><i className="green-dot" /> 相同 {summary.same} <i className="amber-dot" /> 待審核 {changedDiffs.length}</span></div>
            </> : <SchematicPanel beforeUrl={sheetBefore} afterUrl={sheetAfter} beforeFile={sheetBeforeFile} afterFile={sheetAfterFile} beforeName={sheetBeforeName} afterName={sheetAfterName} target={schematicTarget} onTargetChange={setSchematicTarget} diffImage={diffImage} busy={imageBusy} onFile={handleSheetFile} />}
          </section>
        </div>
      </section>
      {pendingImport && <ImportReviewDialog pending={pendingImport} onChange={setPendingImport} onCancel={() => setPendingImport(null)} onConfirm={confirmImport} />}
      {sessionOpen && <SessionDialog records={sessionRecords} onClose={() => setSessionOpen(false)} />}
    </main>
  );
}

function ImportReviewDialog({ pending, onChange, onCancel, onConfirm }: { pending: PendingImport; onChange: (value: PendingImport) => void; onCancel: () => void; onConfirm: () => void }) {
  const sheet = pending.sheets[pending.sheetIndex];
  const header = sheet?.matrix[pending.headerIndex] ?? [];
  const analysis = useMemo(() => analyzeCompanyBomMatrix(sheet?.matrix ?? [], pending.headerIndex, pending.mapping), [sheet, pending.headerIndex, pending.mapping]);
  const errors = analysis.audit.issues.filter((issue) => issue.severity === "error");
  const warnings = analysis.audit.issues.filter((issue) => issue.severity === "warning");
  const columnKeys = Object.keys(companyColumnLabels) as CompanyColumnKey[];
  const changeSheet = (sheetIndex: number) => {
    const matrix = pending.sheets[sheetIndex].matrix;
    const found = findCompanyHeader(matrix);
    const headerIndex = found >= 0 ? found : Math.max(0, matrix.findIndex((row) => row.filter((cell) => String(cell ?? "").trim()).length >= 2));
    onChange({ ...pending, sheetIndex, headerIndex, mapping: detectCompanyColumns(matrix[headerIndex] ?? []) });
  };
  const changeHeader = (headerIndex: number) => onChange({ ...pending, headerIndex, mapping: detectCompanyColumns(sheet.matrix[headerIndex] ?? []) });

  return <div className="modal-backdrop" role="presentation">
    <section className="import-dialog" role="dialog" aria-modal="true" aria-labelledby="import-title">
      <header><div><small>匯入前資料檢查</small><h2 id="import-title">{pending.fileName}</h2></div><button aria-label="關閉" onClick={onCancel}><X size={19} /></button></header>
      <div className="import-controls">
        <label>工作表<select value={pending.sheetIndex} onChange={(event) => changeSheet(Number(event.target.value))}>{pending.sheets.map((item, index) => <option key={item.name} value={index}>{item.name}</option>)}</select></label>
        <label>標題列<select value={pending.headerIndex} onChange={(event) => changeHeader(Number(event.target.value))}>{sheet.matrix.slice(0, 30).map((row, index) => <option key={index} value={index}>第 {index + 1} 列｜{row.filter((cell) => String(cell ?? "").trim()).slice(0, 4).join("、") || "空白"}</option>)}</select></label>
      </div>
      <div className="mapping-grid">
        {columnKeys.map((key) => <label key={key}><span>{companyColumnLabels[key]}{requiredCompanyColumns.includes(key) && <b>*</b>}</span><select value={pending.mapping[key] ?? ""} onChange={(event) => onChange({ ...pending, mapping: { ...pending.mapping, [key]: event.target.value === "" ? undefined : Number(event.target.value) } })}><option value="">未指定</option>{header.map((cell, index) => <option key={index} value={index}>{index + 1}. {String(cell || `未命名欄位 ${index + 1}`)}</option>)}</select></label>)}
      </div>
      <div className="audit-cards"><div><small>資料列</small><strong>{analysis.audit.sourceRows}</strong></div><div><small>料號／位置群組</small><strong>{analysis.audit.groupCount}</strong></div><div><small>插件位置</small><strong>{analysis.audit.positionCount}</strong></div><div className={warnings.length ? "warning" : "ok"}><small>警告</small><strong>{warnings.length}</strong></div></div>
      <div className="issue-list">
        {!analysis.audit.issues.length && <p className="issue-ok"><ShieldCheck size={17} /> 資料檢查通過，可以開始比對。</p>}
        {analysis.audit.issues.map((issue, index) => <p className={`issue ${issue.severity}`} key={`${issue.code}-${index}`}><AlertTriangle size={16} /><span>{issue.message}{issue.rows?.length ? <small>Excel 列：{[...new Set(issue.rows)].join("、")}</small> : null}</span></p>)}
      </div>
      <footer><button className="secondary" onClick={onCancel}>取消</button><button className="primary" disabled={errors.length > 0 || !isCompleteCompanyMapping(pending.mapping)} onClick={onConfirm}>確認匯入</button></footer>
    </section>
  </div>;
}

function SessionDialog({ records, onClose }: { records: ImportRecord[]; onClose: () => void }) {
  return <div className="modal-backdrop" role="presentation"><section className="session-dialog" role="dialog" aria-modal="true" aria-labelledby="session-title"><header><div><small>只保留在目前頁面記憶體</small><h2 id="session-title">此次工作階段</h2></div><button aria-label="關閉" onClick={onClose}><X size={19} /></button></header>{records.length ? <div className="session-list">{records.map((record, index) => <article key={`${record.importedAt}-${index}`}><FileSpreadsheet size={18} /><div><strong>{record.fileName}</strong><small>{record.sheetName}・{record.importedAt}・{record.audit.groupCount} 組料・{record.audit.issues.length} 項提示</small></div></article>)}</div> : <div className="empty-state">目前還沒有匯入紀錄</div>}<footer><button className="primary" onClick={onClose}>完成</button></footer></section></div>;
}

function PrimaryTypeBadge({ item }: { item: BomDiff }) {
  const label = item.matchConfidence === "high" ? "高可信" : item.matchConfidence === "medium" ? "中可信" : "低可信";
  const structure = diffStructureLabel(item);
  return <div className="primary-stack"><span className={`primary-badge ${primaryTypeTone(item.primaryType)}`}>{primaryTypeLabel(item.primaryType)}</span>{structure && <small className="structure-context">所屬架構<br />{structure}</small>}<small className={`confidence ${item.matchConfidence}`}>{item.needsReview ? "待確認 · " : ""}{label}</small><small className="match-reason">{item.matchReason}</small></div>;
}

function ChangeFields({ fields }: { fields: string[] }) {
  return <div className="change-list">{fields.map((field) => {
    const tone = field === "新增元件" || field === "新增替料" || field === "新增料號" ? "added"
      : field === "移除元件" || field === "刪除替料" || field === "刪除料號" ? "removed"
        : field === "新增插件位置" ? "position"
          : field === "移除插件位置" ? "removed-position"
            : field === "數量差異" ? "quantity"
              : field === "更換料號" ? "replacement"
                : "manufacturer";
    return <span className={`change-pill ${tone}`} key={field}>{field}</span>;
  })}</div>;
}

function BPartDifference({ item }: { item: BomDiff }) {
  if (!item.addedParts.length && !item.removedParts.length) {
    return <div className="b-diff-none"><strong>料號無增減</strong></div>;
  }
  const newKeys = new Set(item.newParts.map((part) => canonicalPartNumber(part.part).toUpperCase()));
  const deletedKeys = new Set(item.deletedParts.map((part) => canonicalPartNumber(part.part).toUpperCase()));
  return <div className="b-diff-list">
    {item.addedParts.map((part) => {
      const globallyNew = newKeys.has(canonicalPartNumber(part.part).toUpperCase());
      const label = globallyNew ? "新增料號 · 新版完全新料" : item.replacementPositions.length ? "此位置改用" : item.before ? "新增替料" : "新增料號";
      return <div className="b-diff added" key={`a-${part.part}`}><span>{label}</span><strong>+ {alternativeLabel(part)}</strong></div>;
    })}
    {item.removedParts.map((part) => {
      const globallyDeleted = deletedKeys.has(canonicalPartNumber(part.part).toUpperCase());
      const label = globallyDeleted ? "刪除料號 · 新版完全移除" : item.replacementPositions.length ? "此位置停用" : item.after ? "刪除替料" : "刪除料號";
      return <div className="b-diff removed" key={`r-${part.part}`}><span>{label}</span><strong>− {alternativeLabel(part)}</strong></div>;
    })}
  </div>;
}

function PartList({ item, changedParts, tone }: { item?: BomItem; changedParts: BomAlternative[]; tone: "added" | "removed" }) {
  if (!item?.alternatives.length) return <span className="no-change">—</span>;
  const changedKeys = new Set(changedParts.map((part) => `${part.part.trim().toUpperCase()}|${part.manufacturerPart.trim().toUpperCase()}`));
  return <div className="part-list">{item.alternatives.map((alternative, index) => {
    const key = `${alternative.part.trim().toUpperCase()}|${alternative.manufacturerPart.trim().toUpperCase()}`;
    const changed = changedKeys.has(key);
    return <div className={`part-entry ${changed ? tone : ""}`} key={`${key}-${index}`}>
      <span className="part-role">{index === 0 ? "主" : "替"}</span>
      <div><strong>{alternative.part || "—"}</strong><small><b>製造商料號</b> {alternative.manufacturerPart || "—"}</small><small><b>製造商名稱</b> {alternative.manufacturerName || "—"}</small></div>
    </div>;
  })}</div>;
}

function PositionSummary({ added, removed, replacement, allPositions, onLocate }: { added: string[]; removed: string[]; replacement: string[]; allPositions: string[]; onLocate: (position: string) => void }) {
  if (!added.length && !removed.length && !replacement.length) return allPositions.length ? <div className="position-list unchanged"><small>位置未變・點擊定位</small>{allPositions.map((position) => <button className="position-chip neutral" title={`在線路圖定位 ${position}`} onClick={() => onLocate(position)} key={`n-${position}`}>{position}</button>)}</div> : <span className="no-change">—</span>;
  return <div className="position-list">
    {replacement.map((position) => <button className="position-chip changed" title={`在線路圖定位 ${position}`} onClick={() => onLocate(position)} key={`c-${position}`}>{position} 換料</button>)}
    {added.map((position) => <button className="position-chip added" title={`在線路圖定位 ${position}`} onClick={() => onLocate(position)} key={`a-${position}`}>+ {position}</button>)}
    {removed.map((position) => <button className="position-chip removed" title={`在線路圖定位 ${position}`} onClick={() => onLocate(position)} key={`r-${position}`}>− {position}</button>)}
  </div>;
}

function SchematicPanel({ beforeUrl, afterUrl, beforeFile, afterFile, beforeName, afterName, target, onTargetChange, diffImage, busy, onFile }: { beforeUrl: string | null; afterUrl: string | null; beforeFile: File | null; afterFile: File | null; beforeName: string; afterName: string; target: string; onTargetChange: (value: string) => void; diffImage: string | null; busy: boolean; onFile: (event: ChangeEvent<HTMLInputElement>, side: "before" | "after") => void }) {
  const [view, setView] = useState<"side" | "diff">("side");
  const [syncEnabled, setSyncEnabled] = useState(true);
  const [scales, setScales] = useState<Record<PdfViewerSide, number>>({ before: 1.15, after: 1.15 });
  const [syncState, setSyncState] = useState<PdfScrollSync>({ source: "before", x: 0, y: 0, revision: 0 });
  const isPdf = (name: string) => name.toLowerCase().endsWith(".pdf");
  const changeScale = (side: PdfViewerSide, scale: number) => setScales((current) => syncEnabled ? { before: scale, after: scale } : { ...current, [side]: scale });
  const toggleSync = () => {
    const next = !syncEnabled;
    if (next) setScales((current) => ({ before: current.before, after: current.before }));
    setSyncEnabled(next);
  };
  const preview = (url: string | null, file: File | null, name: string, sideLabel: string, side: PdfViewerSide) => !url ? <div className="sheet-empty"><ImageIcon size={28} /><strong>選擇線路圖</strong><small>文字定位支援含文字層的 PDF</small></div> : isPdf(name) && file ? <PdfSchematicViewer file={file} target={target} sideLabel={sideLabel} side={side} scale={scales[side]} syncEnabled={syncEnabled} syncState={syncState} onScaleChange={changeScale} onSyncScroll={(source, x, y) => setSyncState((current) => ({ source, x, y, revision: current.revision + 1 }))} /> : <div className="image-preview"><img src={url} alt={name} />{target && <span>圖片格式無法搜尋 {target}；請使用含文字層的 PDF。</span>}</div>;
  return <div className="schematic-panel">
    <div className="schematic-toolbar"><div><strong>線路圖與 BOM 連動</strong><small>支援同步縮放／捲動、旋轉文字與本機索引快取</small></div><label className="schematic-search"><Search size={15} /><input value={target} onChange={(event) => onTargetChange(event.target.value.toUpperCase())} placeholder="輸入插件位置，例如 U45" />{target && <button aria-label="清除搜尋" onClick={() => onTargetChange("")}><X size={14} /></button>}</label><div className="schematic-options"><button className={syncEnabled ? "sync-active" : ""} onClick={toggleSync}>{syncEnabled ? "同步檢視中" : "獨立檢視"}</button><button onClick={clearInactivePdfCache}>清除閒置快取</button><div className="view-toggle"><button className={view === "side" ? "active" : ""} onClick={() => setView("side")}>並排定位</button><button className={view === "diff" ? "active" : ""} onClick={() => setView("diff")} disabled={!diffImage}>像素差異</button></div></div></div>
    {view === "diff" && diffImage ? <div className="diff-canvas"><div className="diff-note"><span /> 紅色區域代表前後版本的像素差異</div><img src={diffImage} alt="線路圖像素差異" /></div> : <div className="sheet-grid">{(["before", "after"] as const).map((side) => { const url = side === "before" ? beforeUrl : afterUrl; const file = side === "before" ? beforeFile : afterFile; const name = side === "before" ? beforeName : afterName; const sideLabel = side === "before" ? "前版" : "新版"; return <div className="sheet-card" key={side}><div className="sheet-head"><span>{sideLabel}線路圖</span><strong>{name || "尚未選擇檔案"}</strong><label className="sheet-upload" title={`選擇${sideLabel}線路圖`}><input type="file" hidden accept="image/png,image/jpeg,image/webp,application/pdf" onChange={(e) => onFile(e, side)} /><UploadCloud size={17} /></label></div><div className="sheet-preview">{preview(url, file, name, sideLabel, side)}</div></div>; })}</div>}
    {busy && <div className="processing"><span /> 正在產生差異圖…</div>}
  </div>;
}
