"use client";

import {
  AlertTriangle,
  ArrowRight,
  Check,
  ChevronDown,
  CircuitBoard,
  Download,
  FileSpreadsheet,
  FileText,
  Filter,
  GitCompareArrows,
  History,
  Image as ImageIcon,
  Menu,
  Plus,
  Search,
  UploadCloud,
  X,
} from "lucide-react";
import { ChangeEvent, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

type BomItem = {
  ref: string;
  part: string;
  value: string;
  description: string;
  qty: number;
};

type DiffKind = "added" | "removed" | "changed" | "same";
type BomDiff = {
  ref: string;
  kind: DiffKind;
  before?: BomItem;
  after?: BomItem;
  fields: string[];
};

const demoBefore: BomItem[] = [
  { ref: "C101", part: "CL10B104KB8NNNC", value: "100nF", description: "MLCC 50V X7R 0603", qty: 1 },
  { ref: "C102", part: "GRM188R71H104KA93", value: "100nF", description: "MLCC 50V X7R 0603", qty: 1 },
  { ref: "R205", part: "RC0603FR-0710KL", value: "10kΩ", description: "Resistor 1% 0603", qty: 1 },
  { ref: "R206", part: "RC0603FR-0710KL", value: "10kΩ", description: "Resistor 1% 0603", qty: 1 },
  { ref: "U3", part: "STM32G071CBT6", value: "MCU", description: "ARM Cortex-M0+ 48MHz", qty: 1 },
  { ref: "Q7", part: "2N7002", value: "N-MOS", description: "60V N-Channel MOSFET", qty: 1 },
  { ref: "J4", part: "TYPE-C-16P", value: "USB-C", description: "USB Type-C receptacle", qty: 1 },
];

const demoAfter: BomItem[] = [
  { ref: "C101", part: "GRM188R71H104KA93", value: "100nF", description: "MLCC 50V X7R 0603", qty: 1 },
  { ref: "C102", part: "GRM188R71H104KA93", value: "100nF", description: "MLCC 50V X7R 0603", qty: 1 },
  { ref: "C104", part: "GRM188R61A106KE69", value: "10µF", description: "MLCC 10V X5R 0603", qty: 2 },
  { ref: "R205", part: "RC0603FR-0712KL", value: "12kΩ", description: "Resistor 1% 0603", qty: 1 },
  { ref: "R206", part: "RC0603FR-0710KL", value: "10kΩ", description: "Resistor 1% 0603", qty: 1 },
  { ref: "U3", part: "STM32G071CBT6", value: "MCU", description: "ARM Cortex-M0+ 48MHz", qty: 1 },
  { ref: "J4", part: "TYPE-C-16P", value: "USB-C", description: "USB Type-C receptacle", qty: 1 },
];

const aliases = {
  ref: ["ref", "reference", "references", "designator", "refdes", "reference designator", "位號", "位置"],
  part: ["part", "part number", "part no", "pn", "mpn", "manufacturer part number", "料號", "零件料號"],
  value: ["value", "component value", "規格", "數值", "值"],
  description: ["description", "desc", "comment", "item description", "說明", "描述", "品名"],
  qty: ["qty", "quantity", "count", "數量", "用量"],
};

function normalizeKey(value: string) {
  return value.trim().toLowerCase().replace(/[_-]+/g, " ").replace(/\s+/g, " ");
}

function pick(row: Record<string, unknown>, names: string[]) {
  const entry = Object.entries(row).find(([key]) => names.includes(normalizeKey(key)));
  return entry?.[1] == null ? "" : String(entry[1]).trim();
}

function rowsToBom(rows: Record<string, unknown>[]) {
  return rows
    .map((row, index) => {
      const ref = pick(row, aliases.ref) || `ROW-${index + 1}`;
      return {
        ref,
        part: pick(row, aliases.part),
        value: pick(row, aliases.value),
        description: pick(row, aliases.description),
        qty: Number(pick(row, aliases.qty)) || 1,
      };
    })
    .filter((row) => row.part || row.value || row.description || !row.ref.startsWith("ROW-"));
}

function compareBom(before: BomItem[], after: BomItem[]): BomDiff[] {
  const left = new Map(before.map((item) => [item.ref.toUpperCase(), item]));
  const right = new Map(after.map((item) => [item.ref.toUpperCase(), item]));
  const refs = [...new Set([...left.keys(), ...right.keys()])].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  return refs.map((ref) => {
    const a = left.get(ref);
    const b = right.get(ref);
    if (!a) return { ref, kind: "added", after: b, fields: ["新增元件"] };
    if (!b) return { ref, kind: "removed", before: a, fields: ["移除元件"] };
    const fields = [
      a.part !== b.part ? "料號" : "",
      a.value !== b.value ? "規格" : "",
      a.qty !== b.qty ? "數量" : "",
      a.description !== b.description ? "說明" : "",
    ].filter(Boolean);
    return { ref, kind: fields.length ? "changed" : "same", before: a, after: b, fields };
  });
}

function displayPart(item?: BomItem) {
  if (!item) return "—";
  return item.part || item.value || item.description || "未提供料號";
}

export default function Home() {
  const [tab, setTab] = useState<"bom" | "schematic">("bom");
  const [before, setBefore] = useState(demoBefore);
  const [after, setAfter] = useState(demoAfter);
  const [beforeName, setBeforeName] = useState("PCB_Main_v1.3.xlsx");
  const [afterName, setAfterName] = useState("PCB_Main_v1.4.xlsx");
  const [query, setQuery] = useState("");
  const [filter, setFilter] = useState<DiffKind | "all">("all");
  const [sheetBefore, setSheetBefore] = useState<string | null>(null);
  const [sheetAfter, setSheetAfter] = useState<string | null>(null);
  const [sheetBeforeName, setSheetBeforeName] = useState("");
  const [sheetAfterName, setSheetAfterName] = useState("");
  const [diffImage, setDiffImage] = useState<string | null>(null);
  const [imageBusy, setImageBusy] = useState(false);
  const [mobileNav, setMobileNav] = useState(false);
  const beforeInput = useRef<HTMLInputElement>(null);
  const afterInput = useRef<HTMLInputElement>(null);

  const diffs = useMemo(() => compareBom(before, after), [before, after]);
  const changedDiffs = diffs.filter((item) => item.kind !== "same");
  const summary = useMemo(
    () => ({
      added: diffs.filter((d) => d.kind === "added").length,
      removed: diffs.filter((d) => d.kind === "removed").length,
      changed: diffs.filter((d) => d.kind === "changed").length,
      same: diffs.filter((d) => d.kind === "same").length,
    }),
    [diffs],
  );
  const visible = diffs.filter((item) => {
    const text = `${item.ref} ${displayPart(item.before)} ${displayPart(item.after)} ${item.fields.join(" ")}`.toLowerCase();
    return (filter === "all" || item.kind === filter) && text.includes(query.toLowerCase());
  });

  async function loadBom(file: File, side: "before" | "after") {
    const data = await file.arrayBuffer();
    const book = XLSX.read(data, { type: "array" });
    const sheet = book.Sheets[book.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "" });
    const parsed = rowsToBom(rows);
    if (!parsed.length) {
      window.alert("找不到可辨識的 BOM 資料。請確認第一列是欄位名稱，並包含位號、料號或規格。");
      return;
    }
    if (side === "before") {
      setBefore(parsed);
      setBeforeName(file.name);
    } else {
      setAfter(parsed);
      setAfterName(file.name);
    }
  }

  function handleBomFile(event: ChangeEvent<HTMLInputElement>, side: "before" | "after") {
    const file = event.target.files?.[0];
    if (file) loadBom(file, side).catch(() => window.alert("檔案讀取失敗，請改用 XLSX、XLS 或 CSV 格式。"));
  }

  function handleSheetFile(event: ChangeEvent<HTMLInputElement>, side: "before" | "after") {
    const file = event.target.files?.[0];
    if (!file) return;
    const url = URL.createObjectURL(file);
    if (side === "before") {
      if (sheetBefore) URL.revokeObjectURL(sheetBefore);
      setSheetBefore(url);
      setSheetBeforeName(file.name);
    } else {
      if (sheetAfter) URL.revokeObjectURL(sheetAfter);
      setSheetAfter(url);
      setSheetAfterName(file.name);
    }
  }

  useEffect(() => {
    if (!sheetBefore || !sheetAfter) {
      setDiffImage(null);
      return;
    }
    const leftPdf = sheetBeforeName.toLowerCase().endsWith(".pdf");
    const rightPdf = sheetAfterName.toLowerCase().endsWith(".pdf");
    if (leftPdf || rightPdf) {
      setDiffImage(null);
      return;
    }
    let cancelled = false;
    setImageBusy(true);
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

  function exportCsv() {
    const rows = changedDiffs.map((d) => ({
      狀態: d.kind === "added" ? "新增" : d.kind === "removed" ? "移除" : "變更",
      位號: d.ref,
      變更欄位: d.fields.join("、"),
      前版料號: d.before?.part ?? "",
      後版料號: d.after?.part ?? "",
      前版規格: d.before?.value ?? "",
      後版規格: d.after?.value ?? "",
      前版數量: d.before?.qty ?? "",
      後版數量: d.after?.qty ?? "",
    }));
    const sheet = XLSX.utils.json_to_sheet(rows);
    const book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, sheet, "BOM 差異");
    XLSX.writeFile(book, "BOM_diff_report.xlsx");
  }

  return (
    <main className="app-shell">
      <aside className={`sidebar ${mobileNav ? "open" : ""}`}>
        <div className="brand"><span className="brand-mark"><CircuitBoard size={21} /></span><span>BOM<span>Lens</span></span></div>
        <button className="new-compare" onClick={() => { setTab("bom"); beforeInput.current?.click(); }}><Plus size={18} /> 新增比對</button>
        <nav>
          <p className="nav-label">工作區</p>
          <button className="nav-item active"><GitCompareArrows size={18} /> 目前比對 <span>1</span></button>
          <button className="nav-item"><History size={18} /> 歷史紀錄</button>
          <p className="nav-label recent-label">最近開啟</p>
          <button className="recent-item"><span className="status-dot amber" /> PCB Main Board <small>v1.3 → v1.4</small></button>
          <button className="recent-item"><span className="status-dot green" /> Power Module <small>v2.0 → v2.1</small></button>
          <button className="recent-item"><span className="status-dot gray" /> Sensor Board <small>v0.8 → v0.9</small></button>
        </nav>
        <div className="sidebar-footer"><div className="avatar">JW</div><div><strong>Johnny Wang</strong><small>Engineering</small></div><ChevronDown size={16} /></div>
      </aside>
      {mobileNav && <button className="nav-scrim" aria-label="關閉選單" onClick={() => setMobileNav(false)} />}

      <section className="workspace">
        <header className="topbar">
          <button className="mobile-menu" onClick={() => setMobileNav(true)} aria-label="開啟選單"><Menu /></button>
          <div><div className="breadcrumb">專案 <span>/</span> PCB Main Board</div><h1>版本比對 <span className="version-pill">v1.3 → v1.4</span></h1></div>
          <div className="top-actions"><span className="saved"><Check size={14} /> 已儲存</span><button className="secondary" onClick={() => window.print()}><FileText size={17} /> 列印報告</button><button className="primary" onClick={exportCsv}><Download size={17} /> 匯出差異</button></div>
        </header>

        <div className="content">
          <section className="upload-bar">
            <div className="upload-title"><UploadCloud size={20} /><div><strong>比對來源</strong><small>上傳檔案後會立即在本機完成分析</small></div></div>
            <div className="file-pair">
              <button className="file-chip" onClick={() => beforeInput.current?.click()}><span className="file-icon"><FileSpreadsheet size={18} /></span><span><small>前版 BOM</small><strong>{beforeName}</strong></span><X size={15} className="chip-x" /></button>
              <ArrowRight size={18} className="pair-arrow" />
              <button className="file-chip" onClick={() => afterInput.current?.click()}><span className="file-icon after"><FileSpreadsheet size={18} /></span><span><small>後版 BOM</small><strong>{afterName}</strong></span><X size={15} className="chip-x" /></button>
              <input ref={beforeInput} hidden type="file" accept=".xlsx,.xls,.csv,.tsv" onChange={(e) => handleBomFile(e, "before")} />
              <input ref={afterInput} hidden type="file" accept=".xlsx,.xls,.csv,.tsv" onChange={(e) => handleBomFile(e, "after")} />
            </div>
          </section>

          <section className="summary-grid">
            <article className="summary-card total"><span className="summary-icon"><GitCompareArrows /></span><div><small>差異總數</small><strong>{changedDiffs.length}</strong><p>共比對 {diffs.length} 個位號</p></div></article>
            <article className="summary-card added"><span className="summary-icon"><Plus /></span><div><small>新增</small><strong>{summary.added}</strong><p>後版新增元件</p></div></article>
            <article className="summary-card removed"><span className="summary-icon"><X /></span><div><small>移除</small><strong>{summary.removed}</strong><p>後版已移除</p></div></article>
            <article className="summary-card changed"><span className="summary-icon"><AlertTriangle /></span><div><small>內容變更</small><strong>{summary.changed}</strong><p>料號、規格或數量</p></div></article>
          </section>

          <section className="panel">
            <div className="tabs"><button className={tab === "bom" ? "active" : ""} onClick={() => setTab("bom")}><FileSpreadsheet size={18} /> BOM 差異 <span>{changedDiffs.length}</span></button><button className={tab === "schematic" ? "active" : ""} onClick={() => setTab("schematic")}><CircuitBoard size={18} /> 線路圖比對</button></div>

            {tab === "bom" ? <>
              <div className="table-tools"><label className="search"><Search size={17} /><input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="搜尋位號、料號或變更內容" /></label><div className="filters"><Filter size={16} />{(["all", "added", "removed", "changed"] as const).map((key) => <button key={key} onClick={() => setFilter(key)} className={filter === key ? "active" : ""}>{key === "all" ? "全部" : key === "added" ? "新增" : key === "removed" ? "移除" : "變更"}</button>)}</div></div>
              <div className="table-wrap"><table><thead><tr><th>狀態</th><th>位號</th><th>變更欄位</th><th>前版</th><th></th><th>後版</th><th>數量</th></tr></thead><tbody>{visible.map((item) => <tr key={item.ref} className={item.kind === "same" ? "same-row" : ""}><td><span className={`badge ${item.kind}`}>{item.kind === "added" ? "新增" : item.kind === "removed" ? "移除" : item.kind === "changed" ? "變更" : "相同"}</span></td><td><strong className="refdes">{item.ref}</strong></td><td><span className="field-change">{item.fields.join("、") || "—"}</span></td><td><div className="part-cell"><strong>{displayPart(item.before)}</strong><small>{item.before?.value || item.before?.description || "—"}</small></div></td><td><ArrowRight size={16} className="row-arrow" /></td><td><div className="part-cell"><strong>{displayPart(item.after)}</strong><small>{item.after?.value || item.after?.description || "—"}</small></div></td><td><span className={item.before?.qty !== item.after?.qty ? "qty changed-qty" : "qty"}>{item.before?.qty ?? 0} → {item.after?.qty ?? 0}</span></td></tr>)}</tbody></table>{!visible.length && <div className="empty-state">沒有符合條件的差異</div>}</div>
              <div className="table-footer"><span>顯示 {visible.length} 筆，共 {diffs.length} 個位號</span><span className="legend"><i className="green-dot" /> 相同 {summary.same} <i className="amber-dot" /> 待審核 {changedDiffs.length}</span></div>
            </> : <SchematicPanel beforeUrl={sheetBefore} afterUrl={sheetAfter} beforeName={sheetBeforeName} afterName={sheetAfterName} diffImage={diffImage} busy={imageBusy} onFile={handleSheetFile} />}
          </section>
        </div>
      </section>
    </main>
  );
}

function SchematicPanel({ beforeUrl, afterUrl, beforeName, afterName, diffImage, busy, onFile }: { beforeUrl: string | null; afterUrl: string | null; beforeName: string; afterName: string; diffImage: string | null; busy: boolean; onFile: (event: ChangeEvent<HTMLInputElement>, side: "before" | "after") => void }) {
  const [view, setView] = useState<"side" | "diff">("side");
  const isPdf = (name: string) => name.toLowerCase().endsWith(".pdf");
  const preview = (url: string | null, name: string) => !url ? <div className="sheet-empty"><ImageIcon size={28} /><strong>拖曳或選擇線路圖</strong><small>支援 PNG、JPG、WebP、PDF</small></div> : isPdf(name) ? <embed src={url} type="application/pdf" /> : <img src={url} alt={name} />;
  return <div className="schematic-panel">
    <div className="schematic-toolbar"><div><strong>線路圖視覺比對</strong><small>圖片會產生紅色像素差異；PDF 提供並排審閱</small></div><div className="view-toggle"><button className={view === "side" ? "active" : ""} onClick={() => setView("side")}>並排</button><button className={view === "diff" ? "active" : ""} onClick={() => setView("diff")} disabled={!diffImage}>差異圖</button></div></div>
    {view === "diff" && diffImage ? <div className="diff-canvas"><div className="diff-note"><span /> 紅色區域代表前後版本的像素差異</div><img src={diffImage} alt="線路圖像素差異" /></div> : <div className="sheet-grid">{(["before", "after"] as const).map((side) => { const url = side === "before" ? beforeUrl : afterUrl; const name = side === "before" ? beforeName : afterName; return <label className="sheet-card" key={side}><input type="file" hidden accept="image/png,image/jpeg,image/webp,application/pdf" onChange={(e) => onFile(e, side)} /><div className="sheet-head"><span>{side === "before" ? "前版線路圖" : "後版線路圖"}</span><strong>{name || "尚未選擇檔案"}</strong><UploadCloud size={17} /></div><div className="sheet-preview">{preview(url, name)}</div></label>; })}</div>}
    {busy && <div className="processing"><span /> 正在產生差異圖…</div>}
  </div>;
}
