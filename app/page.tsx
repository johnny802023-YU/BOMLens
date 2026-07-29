"use client";
/* eslint-disable @next/next/no-img-element -- previews use local object URLs and canvas output */

import {
  AlertTriangle,
  ArrowRight,
  ChevronDown,
  ChevronRight,
  CircuitBoard,
  Download,
  FileSpreadsheet,
  Filter,
  GitCompareArrows,
  History,
  Image as ImageIcon,
  Layers,
  Link2,
  Menu,
  Maximize2,
  Minimize2,
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
import type { OriginalBomSource } from "./export-report";
import { clearInactivePdfCache, PdfSchematicViewer, type PdfScrollSync, type PdfViewerSide } from "./pdf-schematic-viewer";
import {
  canonicalPartNumber,
  bomStructureLabel,
  bomDiffDisplayFields,
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
import {
  calculateMva,
  customerColumnLabels,
  detectCustomerColumns,
  detectPlacementColumns,
  findCustomerHeader,
  findPlacementHeader,
  isCompleteCustomerMapping,
  isCompletePlacementMapping,
  mapCustomerBom,
  parseCustomerBomMatrix,
  parsePlacementMatrix,
  placementColumnLabels,
  type CustomerBomColumnKey,
  type CustomerBomColumnMapping,
  type CustomerBomRecord,
  type CustomerMappingResult,
  type MvaSummary,
  type PlacementColumnKey,
  type PlacementColumnMapping,
  type PlacementRecord,
} from "./supplemental-logic";

type FieldFilter = "新增料號" | "新增替料" | "新增插件位置" | "刪除料號" | "刪除替料" | "移除插件位置" | "更換料號" | "製程別放置異常";
type ImpactFilter = "all" | "purchase" | "deleted" | "review";
type DiffGroupKey = "added" | "removed" | "changed" | "review";

type ImportSheet = { name: string; matrix: unknown[][] };
type PendingImport = {
  side: "before" | "after";
  fileName: string;
  sourceData: ArrayBuffer;
  sheets: ImportSheet[];
  sheetIndex: number;
  headerIndex: number;
  mapping: CompanyColumnMapping;
};
type PendingCustomerImport = {
  side: "before" | "after";
  fileName: string;
  sheets: ImportSheet[];
  sheetIndex: number;
  headerIndex: number;
  mapping: CustomerBomColumnMapping;
};
type PendingPlacementImport = {
  side: "before" | "after";
  fileName: string;
  sheets: ImportSheet[];
  sheetIndex: number;
  headerIndex: number;
  mapping: PlacementColumnMapping;
};
type ImportRecord = { fileName: string; sheetName: string; importedAt: string; audit: ImportAudit };

const fieldFilterOptions: Array<{ key: FieldFilter; label: string }> = [
  { key: "新增料號", label: "新增料號" },
  { key: "新增替料", label: "新增替料" },
  { key: "新增插件位置", label: "新增插件位置" },
  { key: "刪除料號", label: "刪除料號" },
  { key: "刪除替料", label: "刪除替料" },
  { key: "移除插件位置", label: "移除插件位置" },
  { key: "更換料號", label: "更換料號" },
  { key: "製程別放置異常", label: "SMT／DIP 放置異常" },
];

const primaryTypeLabels: Record<DiffPrimaryType, string> = {
  componentAdded: "新增",
  substituteAdded: "新增",
  componentRemoved: "刪除",
  substituteRemoved: "刪除",
  partReplaced: "變更",
  positionChanged: "變更",
  same: "相同",
};

const diffGroupMeta: Array<{ key: DiffGroupKey; label: string; description: string }> = [
  { key: "added", label: "新增", description: "新版加入的主料或替料" },
  { key: "removed", label: "刪除", description: "新版移除的主料或替料" },
  { key: "changed", label: "變更", description: "插件位置、數量或同位置換料；另偵測 SMT／DIP 製程別異動" },
  { key: "review", label: "待人工確認", description: "配對不明確或料號跨 SMT／DIP 架構，需要人工判斷" },
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
  return primaryTypeLabels[type];
}

function primaryTypeTone(type: DiffPrimaryType) {
  if (type === "componentAdded" || type === "substituteAdded") return "added";
  if (type === "componentRemoved" || type === "substituteRemoved") return "removed";
  if (type === "partReplaced") return "replacement";
  return "changed";
}

function diffGroupKey(item: BomDiff): DiffGroupKey {
  if (item.needsReview) return "review";
  if (item.primaryType === "componentAdded" || item.primaryType === "substituteAdded") return "added";
  if (item.primaryType === "componentRemoved" || item.primaryType === "substituteRemoved") return "removed";
  return "changed";
}

function diffRowKey(item: BomDiff, index: number) {
  return `${item.before?.structureKey ?? "none"}>${item.after?.structureKey ?? "none"}:${item.ref}:${index}`;
}

function diffStructureLabel(item: BomDiff) {
  const before = bomStructureLabel(item.before);
  const after = bomStructureLabel(item.after);
  if (before && after && before !== after) return `${before} → ${after}`;
  return after || before;
}

export default function Home() {
  const [tab, setTab] = useState<"bom" | "customer" | "schematic">("bom");
  const [before, setBefore] = useState(demoBefore);
  const [after, setAfter] = useState(demoAfter);
  const [beforeName, setBeforeName] = useState("PCB_Main_v1.3.xlsx");
  const [afterName, setAfterName] = useState("PCB_Main_v1.4.xlsx");
  const [query, setQuery] = useState("");
  const [selectedFields, setSelectedFields] = useState<FieldFilter[]>([]);
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
  const [pendingCustomerImport, setPendingCustomerImport] = useState<PendingCustomerImport | null>(null);
  const [pendingPlacementImport, setPendingPlacementImport] = useState<PendingPlacementImport | null>(null);
  const [customerBeforeRecords, setCustomerBeforeRecords] = useState<CustomerBomRecord[] | null>(null);
  const [customerAfterRecords, setCustomerAfterRecords] = useState<CustomerBomRecord[] | null>(null);
  const [customerBeforeName, setCustomerBeforeName] = useState("");
  const [customerAfterName, setCustomerAfterName] = useState("");
  const [placementBeforeRecords, setPlacementBeforeRecords] = useState<PlacementRecord[] | null>(null);
  const [placementAfterRecords, setPlacementAfterRecords] = useState<PlacementRecord[] | null>(null);
  const [placementBeforeName, setPlacementBeforeName] = useState("");
  const [placementAfterName, setPlacementAfterName] = useState("");
  const [beforeAudit, setBeforeAudit] = useState<ImportRecord | null>(null);
  const [afterAudit, setAfterAudit] = useState<ImportRecord | null>(null);
  const [originalBefore, setOriginalBefore] = useState<OriginalBomSource | null>(null);
  const [originalAfter, setOriginalAfter] = useState<OriginalBomSource | null>(null);
  const [sessionOpen, setSessionOpen] = useState(false);
  const [sessionRecords, setSessionRecords] = useState<ImportRecord[]>([]);
  const [schematicTarget, setSchematicTarget] = useState("");
  const [selectedDiff, setSelectedDiff] = useState<BomDiff | null>(null);
  const [collapsedGroups, setCollapsedGroups] = useState<DiffGroupKey[]>([]);
  const beforeInput = useRef<HTMLInputElement>(null);
  const afterInput = useRef<HTMLInputElement>(null);
  const customerBeforeInput = useRef<HTMLInputElement>(null);
  const customerAfterInput = useRef<HTMLInputElement>(null);
  const placementBeforeInput = useRef<HTMLInputElement>(null);
  const placementAfterInput = useRef<HTMLInputElement>(null);

  const customerBeforeResult = useMemo(() => customerBeforeRecords ? mapCustomerBom(before, customerBeforeRecords) : null, [before, customerBeforeRecords]);
  const customerAfterResult = useMemo(() => customerAfterRecords ? mapCustomerBom(after, customerAfterRecords) : null, [after, customerAfterRecords]);
  const effectiveBefore = customerBeforeResult?.items ?? before.map((item) => ({ ...item, customerMappingStatus: "not-imported" as const }));
  const effectiveAfter = customerAfterResult?.items ?? after.map((item) => ({ ...item, customerMappingStatus: "not-imported" as const }));
  const mvaBefore = useMemo(() => placementBeforeRecords ? calculateMva(before, placementBeforeRecords) : null, [before, placementBeforeRecords]);
  const mvaAfter = useMemo(() => placementAfterRecords ? calculateMva(after, placementAfterRecords) : null, [after, placementAfterRecords]);
  const diffs = useMemo(() => compareBom(effectiveBefore, effectiveAfter), [effectiveBefore, effectiveAfter]);
  const changedDiffs = diffs.filter((item) => item.categories.length > 0);
  const reviewDiffs = changedDiffs.filter((item) => item.needsReview);
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
    const text = `${item.ref} ${diffStructureLabel(item)} ${displayPart(item.before)} ${displayPart(item.after)} ${item.before?.manufacturerPart ?? ""} ${item.after?.manufacturerPart ?? ""} ${manufacturerNames} ${bomDiffDisplayFields(item).join(" ")} ${item.addedPositions.join(" ")} ${item.removedPositions.join(" ")}`.toLowerCase();
    const matchesFilter = selectedFields.length === 0 || selectedFields.some((field) => item.fields.includes(field));
    const matchesImpact = impactFilter === "all"
      || (impactFilter === "purchase" && item.newParts.length > 0)
      || (impactFilter === "deleted" && item.deletedParts.length > 0)
      || (impactFilter === "review" && item.needsReview);
    return matchesFilter && matchesImpact && text.includes(query.toLowerCase());
  });
  const visible = sortBomDiffsForAll(matchingDiffs);
  const visibleGroups = diffGroupMeta
    .map((group) => ({ ...group, items: visible.filter((item) => diffGroupKey(item) === group.key) }))
    .filter((group) => group.items.length > 0);

  function toggleFieldFilter(field: FieldFilter) {
    setSelectedFields((current) => current.includes(field)
      ? current.filter((selected) => selected !== field)
      : [...current, field]);
  }

  function toggleDiffGroup(group: DiffGroupKey) {
    setCollapsedGroups((current) => current.includes(group)
      ? current.filter((item) => item !== group)
      : [...current, group]);
  }

  async function readWorkbook(file: File) {
    const data = await file.arrayBuffer();
    const book = XLSX.read(data, { type: "array", cellStyles: true, cellNF: true, cellDates: true });
    return {
      data,
      book,
      sheets: book.SheetNames.map((name) => ({ name, matrix: XLSX.utils.sheet_to_json<unknown[]>(book.Sheets[name], { header: 1, defval: "", raw: false }) })),
    };
  }

  async function loadBom(file: File, side: "before" | "after") {
    const { data, book, sheets } = await readWorkbook(file);
    const sourceData = /\.xls[xm]$/i.test(file.name)
      ? data
      : XLSX.write(book, { type: "array", bookType: "xlsx", cellStyles: true });
    const candidateIndex = sheets.findIndex((sheet) => findCompanyHeader(sheet.matrix) >= 0);
    const sheetIndex = candidateIndex >= 0 ? candidateIndex : 0;
    const matrix = sheets[sheetIndex]?.matrix ?? [];
    const detectedHeader = findCompanyHeader(matrix);
    const headerIndex = detectedHeader >= 0 ? detectedHeader : Math.max(0, matrix.findIndex((row) => row.filter((cell) => String(cell ?? "").trim()).length >= 2));
    setPendingImport({ side, fileName: file.name, sourceData, sheets, sheetIndex, headerIndex, mapping: detectCompanyColumns(matrix[headerIndex] ?? []) });
  }

  async function loadCustomerBom(file: File, side: "before" | "after") {
    const { sheets } = await readWorkbook(file);
    const candidateIndex = sheets.findIndex((sheet) => findCustomerHeader(sheet.matrix) >= 0);
    const sheetIndex = candidateIndex >= 0 ? candidateIndex : 0;
    const matrix = sheets[sheetIndex]?.matrix ?? [];
    const detectedHeader = findCustomerHeader(matrix);
    const headerIndex = detectedHeader >= 0 ? detectedHeader : Math.max(0, matrix.findIndex((row) => row.filter((cell) => String(cell ?? "").trim()).length >= 2));
    setPendingCustomerImport({ side, fileName: file.name, sheets, sheetIndex, headerIndex, mapping: detectCustomerColumns(matrix[headerIndex] ?? []) });
  }

  async function loadPlacement(file: File, side: "before" | "after") {
    const { sheets } = await readWorkbook(file);
    const candidateIndex = sheets.findIndex((sheet) => findPlacementHeader(sheet.matrix) >= 0);
    const sheetIndex = candidateIndex >= 0 ? candidateIndex : 0;
    const matrix = sheets[sheetIndex]?.matrix ?? [];
    const detectedHeader = findPlacementHeader(matrix);
    const headerIndex = detectedHeader >= 0 ? detectedHeader : Math.max(0, matrix.findIndex((row) => row.filter((cell) => String(cell ?? "").trim()).length >= 2));
    setPendingPlacementImport({ side, fileName: file.name, sheets, sheetIndex, headerIndex, mapping: detectPlacementColumns(matrix[headerIndex] ?? []) });
  }

  function confirmImport() {
    if (!pendingImport) return;
    const sheet = pendingImport.sheets[pendingImport.sheetIndex];
    const result = analyzeCompanyBomMatrix(sheet.matrix, pendingImport.headerIndex, pendingImport.mapping);
    if (!isCompleteCompanyMapping(pendingImport.mapping) || result.audit.issues.some((issue) => issue.severity === "error")) return;
    const items = result.items.map((item) => ({ ...item, sourceSheet: sheet.name }));
    const record = { fileName: pendingImport.fileName, sheetName: sheet.name, importedAt: new Date().toLocaleString("zh-TW", { hour12: false }), audit: result.audit };
    const original = { fileName: pendingImport.fileName, sheetName: sheet.name, data: pendingImport.sourceData, matrix: sheet.matrix };
    if (pendingImport.side === "before") {
      setBefore(items); setBeforeName(pendingImport.fileName); setBeforeAudit(record); setOriginalBefore(original);
    } else {
      setAfter(items); setAfterName(pendingImport.fileName); setAfterAudit(record); setOriginalAfter(original);
    }
    setSessionRecords((records) => [record, ...records].slice(0, 20));
    setPendingImport(null);
  }

  function confirmCustomerImport() {
    if (!pendingCustomerImport || !isCompleteCustomerMapping(pendingCustomerImport.mapping)) return;
    const sheet = pendingCustomerImport.sheets[pendingCustomerImport.sheetIndex];
    const records = parseCustomerBomMatrix(sheet.matrix, pendingCustomerImport.headerIndex, pendingCustomerImport.mapping);
    if (pendingCustomerImport.side === "before") {
      setCustomerBeforeRecords(records); setCustomerBeforeName(pendingCustomerImport.fileName);
    } else {
      setCustomerAfterRecords(records); setCustomerAfterName(pendingCustomerImport.fileName);
    }
    setPendingCustomerImport(null);
  }

  function confirmPlacementImport() {
    if (!pendingPlacementImport || !isCompletePlacementMapping(pendingPlacementImport.mapping)) return;
    const sheet = pendingPlacementImport.sheets[pendingPlacementImport.sheetIndex];
    const records = parsePlacementMatrix(sheet.matrix, pendingPlacementImport.headerIndex, pendingPlacementImport.mapping);
    if (pendingPlacementImport.side === "before") {
      setPlacementBeforeRecords(records); setPlacementBeforeName(pendingPlacementImport.fileName);
    } else {
      setPlacementAfterRecords(records); setPlacementAfterName(pendingPlacementImport.fileName);
    }
    setPendingPlacementImport(null);
  }

  function resetComparison() {
    setBefore([]); setAfter([]); setBeforeName(""); setAfterName(""); setBeforeAudit(null); setAfterAudit(null); setOriginalBefore(null); setOriginalAfter(null);
    setQuery(""); setSelectedFields([]); setImpactFilter("all"); setSelectedDiff(null); setCollapsedGroups([]); setTab("bom");
    setCustomerBeforeRecords(null); setCustomerAfterRecords(null); setCustomerBeforeName(""); setCustomerAfterName("");
    setPlacementBeforeRecords(null); setPlacementAfterRecords(null); setPlacementBeforeName(""); setPlacementAfterName("");
    window.setTimeout(() => beforeInput.current?.click(), 0);
  }

  function handleBomFile(event: ChangeEvent<HTMLInputElement>, side: "before" | "after") {
    const file = event.target.files?.[0];
    if (file) loadBom(file, side).catch(() => window.alert("檔案讀取失敗，請改用 XLSX、XLS 或 CSV 格式。"));
  }

  function handleSupplementalFile(event: ChangeEvent<HTMLInputElement>, side: "before" | "after", kind: "customer" | "placement") {
    const file = event.target.files?.[0];
    if (!file) return;
    const loader = kind === "customer" ? loadCustomerBom : loadPlacement;
    loader(file, side).catch(() => window.alert("檔案讀取失敗，請改用 XLSX、XLS、CSV 或 TSV 格式。"));
    event.target.value = "";
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
      await exportBomReport(visible, beforeName, afterName, { before: beforeAudit, after: afterAudit, originalBefore, originalAfter, customerBefore: customerBeforeResult, customerAfter: customerAfterResult, mvaBefore, mvaAfter });
    } catch {
      window.alert("報表產生失敗，請重新整理後再試一次。");
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
          <div><div className="breadcrumb">本機工具 <span>/</span> BOM 版本比對</div><h1>版本比對 <span className="version-pill">{beforeName || "舊版"} → {afterName || "新版"}</span></h1></div>
          <div className="top-actions"><span className="saved offline"><ShieldCheck size={14} /> 本機離線</span><button className="primary" onClick={exportCsv} title="依目前搜尋與篩選結果匯出"><Download size={17} /> 匯出差異</button></div>
        </header>

        <div className="content">
          <section className="privacy-banner" role="status"><span><ShieldCheck size={19} /></span><div><strong>離線隱私模式</strong><p>檔案只在這台電腦的瀏覽器記憶體內分析，不會上傳、同步或儲存；關閉頁面後即清除。</p></div><span className="local-only">LOCAL ONLY</span></section>
          <section className="upload-bar">
            <div className="upload-title"><UploadCloud size={20} /><div><strong>比對來源</strong><small>選擇檔案後立即在本機完成分析</small></div></div>
            <div className="file-pair">
              <div className="file-chip"><button className="file-select" onClick={() => beforeInput.current?.click()}><span className="file-icon"><FileSpreadsheet size={18} /></span><span><small>舊版 BOM</small><strong>{beforeName || "選擇檔案"}</strong></span></button>{beforeName && <button className="chip-x" aria-label="移除舊版 BOM" onClick={() => { setBefore([]); setBeforeName(""); setBeforeAudit(null); setOriginalBefore(null); }}><X size={15} /></button>}</div>
              <ArrowRight size={18} className="pair-arrow" />
              <div className="file-chip"><button className="file-select" onClick={() => afterInput.current?.click()}><span className="file-icon after"><FileSpreadsheet size={18} /></span><span><small>新版 BOM</small><strong>{afterName || "選擇檔案"}</strong></span></button>{afterName && <button className="chip-x" aria-label="移除新版 BOM" onClick={() => { setAfter([]); setAfterName(""); setAfterAudit(null); setOriginalAfter(null); }}><X size={15} /></button>}</div>
              <input ref={beforeInput} hidden type="file" accept=".xlsx,.xls,.csv,.tsv" onChange={(e) => handleBomFile(e, "before")} />
              <input ref={afterInput} hidden type="file" accept=".xlsx,.xls,.csv,.tsv" onChange={(e) => handleBomFile(e, "after")} />
            </div>
          </section>
          {(beforeAudit || afterAudit) && <section className="audit-strip" aria-label="匯入檢查摘要">
            {[{ label: "舊版", audit: beforeAudit }, { label: "新版", audit: afterAudit }].map(({ label, audit }) => audit && <div key={label}><ShieldCheck size={16} /><span><strong>{label}匯入完成</strong><small>{audit.sheetName}・{audit.audit.groupCount} 組料・{audit.audit.positionCount} 個位置・{audit.audit.issues.filter((issue) => issue.severity === "warning").length} 項警告</small></span></div>)}
          </section>}
          <section className="format-strip" aria-label="BOM 欄位規則">
            <span><b>項次</b> 只切分同架構主替料，不跨版比對</span><span><b>架構</b> 69 → VB-D／60／VB-T／08 PCB</span><span><b>主件料號</b> 取最右 12 碼</span><span><b>數量</b> 一般數量</span><span><b>插件位置</b> 優先計數</span><span><b>製造廠商</b> 僅顯示</span><span><b>製造廠商料號</b> 完全一致驗證</span><span><b>TPN（R欄）</b> 讀取 BOM R欄並比對；Location＋MPN 命中但 R欄缺漏時請 RD 維護</span><strong>依標題名稱自動定位欄位</strong>
          </section>

          <MvaPanel
            beforeName={placementBeforeName}
            afterName={placementAfterName}
            before={mvaBefore}
            after={mvaAfter}
            onBefore={() => placementBeforeInput.current?.click()}
            onAfter={() => placementAfterInput.current?.click()}
            onClearBefore={() => { setPlacementBeforeRecords(null); setPlacementBeforeName(""); }}
            onClearAfter={() => { setPlacementAfterRecords(null); setPlacementAfterName(""); }}
          />
          <input ref={placementBeforeInput} hidden type="file" accept=".xlsx,.xls,.csv,.tsv" onChange={(event) => handleSupplementalFile(event, "before", "placement")} />
          <input ref={placementAfterInput} hidden type="file" accept=".xlsx,.xls,.csv,.tsv" onChange={(event) => handleSupplementalFile(event, "after", "placement")} />

          <section className="summary-grid">
            <article className="summary-card total"><span className="summary-icon"><GitCompareArrows /></span><div><small>差異群組</small><strong>{changedDiffs.length}</strong><p>共配對 {diffs.length} 個料號／位置群組</p></div></article>
            <article className="summary-card added"><span className="summary-icon"><Plus /></span><div><small>新增料號</small><strong>{summary.added}</strong><p>舊版整份 BOM 未出現</p></div></article>
            <article className="summary-card removed"><span className="summary-icon"><X /></span><div><small>刪除料號</small><strong>{summary.removed}</strong><p>新版整份 BOM 已無使用</p></div></article>
            <article className="summary-card changed"><span className="summary-icon"><AlertTriangle /></span><div><small>變更群組</small><strong>{summary.changed}</strong><p>數量、位置或同位置換料</p></div></article>
          </section>

          <section className="panel">
            <div className="tabs"><button className={tab === "bom" ? "active" : ""} onClick={() => setTab("bom")}><FileSpreadsheet size={18} /> BOM 差異 <span>{changedDiffs.length}</span></button><button className={tab === "customer" ? "active" : ""} onClick={() => setTab("customer")}><Link2 size={18} /> 客戶 BOM TPN 對應 <span>{(customerBeforeResult?.counts.matched ?? 0) + (customerAfterResult?.counts.matched ?? 0)}</span></button><button className={tab === "schematic" ? "active" : ""} onClick={() => setTab("schematic")}><CircuitBoard size={18} /> 線路圖比對</button></div>

            {tab === "bom" ? <>
              <div className="table-tools">
                <label className="search"><Search size={17} /><input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="搜尋主件料號、製造廠商或插件位置" /></label>
                <div className="filter-panel">
                  <div className="filters primary-filters"><span className="filter-title"><Filter size={14} />差異項目（可複選）</span><button onClick={() => setSelectedFields([])} className={selectedFields.length === 0 ? "active" : ""}>全部</button>{fieldFilterOptions.map((option) => <button key={option.key} onClick={() => toggleFieldFilter(option.key)} className={selectedFields.includes(option.key) ? "active" : ""} aria-pressed={selectedFields.includes(option.key)}>{option.label}</button>)}</div>
                  <div className="filters impact-filters"><span className="filter-title">影響條件</span>{([{ key: "all", label: "不限" }, { key: "purchase", label: "新版完全新料" }, { key: "deleted", label: "新版完全移除" }, { key: "review", label: "待人工確認" }] as const).map((option) => <button key={option.key} onClick={() => setImpactFilter(option.key)} className={impactFilter === option.key ? "active" : ""}>{option.label}</button>)}</div>
                </div>
              </div>
              <div className="table-wrap">
                <table className="diff-table">
                  <thead><tr><th>主要異動</th><th>差異項目</th><th>料號異動</th><th>插件位置差異</th><th>舊版主件／製造廠商／TPN</th><th></th><th>新版主件／製造廠商／TPN</th><th>數量</th></tr></thead>
                  {visibleGroups.map((group) => {
                    const collapsed = collapsedGroups.includes(group.key);
                    return <tbody className={`diff-group ${group.key}`} key={group.key}>
                      <tr className="diff-group-heading"><td colSpan={8}><button type="button" onClick={() => toggleDiffGroup(group.key)} aria-expanded={!collapsed}>
                        <span className="group-chevron">{collapsed ? <ChevronRight size={15} /> : <ChevronDown size={15} />}</span>
                        <strong>{group.label}</strong><span className="group-count">{group.items.length}</span><small>{group.description}</small>
                      </button></td></tr>
                      {!collapsed && group.items.map((item, index) => <tr
                        className={`diff-row ${primaryTypeTone(item.primaryType)} ${selectedDiff === item ? "selected" : ""}`}
                        key={diffRowKey(item, index)}
                        tabIndex={0}
                        aria-label={`查看${primaryTypeLabel(item.primaryType)}差異詳細資料`}
                        onClick={() => setSelectedDiff(item)}
                        onKeyDown={(event) => {
                          if (event.key === "Enter" || event.key === " ") {
                            event.preventDefault();
                            setSelectedDiff(item);
                          }
                        }}
                      >
                        <td><PrimaryTypeBadge item={item} /></td>
                        <td><ChangeFields fields={bomDiffDisplayFields(item)} tone={primaryTypeTone(item.primaryType)} /></td>
                        <td><BPartDifference item={item} /></td>
                        <td><PositionSummary added={item.addedPositions} removed={item.removedPositions} replacement={item.replacementPositions} allPositions={item.after?.positions ?? item.before?.positions ?? []} onLocate={(position) => { setSchematicTarget(position); setSelectedDiff(null); setTab("schematic"); }} /></td>
                        <td><PartList item={item.before} changedParts={item.removedParts} tone="removed" /></td>
                        <td><ArrowRight size={16} className="row-arrow" /></td>
                        <td><PartList item={item.after} changedParts={item.addedParts} tone="added" /></td>
                        <td><span className={item.before?.qty !== item.after?.qty ? "qty changed-qty" : "qty"}>{item.before?.qty ?? 0} → {item.after?.qty ?? 0}</span></td>
                      </tr>)}
                    </tbody>;
                  })}
                </table>
                {!visible.length && <div className="empty-state">沒有符合條件的差異</div>}
              </div>
              <div className="table-footer"><span>顯示 {visible.length} 筆，共 {diffs.length} 個料號／位置群組；項次名稱不列入差異</span><span className="legend"><i className="green-dot" /> 相同 {summary.same}<button type="button" className={impactFilter === "review" ? "review-summary active" : "review-summary"} onClick={() => setImpactFilter(impactFilter === "review" ? "all" : "review")} title="低可信或配對不明確、需要人工確認的群組"><i className="amber-dot" /> 人工待審核 {reviewDiffs.length}</button></span></div>
            </> : tab === "customer" ? <CustomerMappingPanel
              beforeName={customerBeforeName}
              afterName={customerAfterName}
              before={customerBeforeResult}
              after={customerAfterResult}
              onBefore={() => customerBeforeInput.current?.click()}
              onAfter={() => customerAfterInput.current?.click()}
              onClearBefore={() => { setCustomerBeforeRecords(null); setCustomerBeforeName(""); }}
              onClearAfter={() => { setCustomerAfterRecords(null); setCustomerAfterName(""); }}
            /> : <SchematicPanel beforeUrl={sheetBefore} afterUrl={sheetAfter} beforeFile={sheetBeforeFile} afterFile={sheetAfterFile} beforeName={sheetBeforeName} afterName={sheetAfterName} target={schematicTarget} onTargetChange={setSchematicTarget} diffImage={diffImage} busy={imageBusy} diffs={changedDiffs} onFile={handleSheetFile} />}
          </section>
          <input ref={customerBeforeInput} hidden type="file" accept=".xlsx,.xls,.csv,.tsv" onChange={(event) => handleSupplementalFile(event, "before", "customer")} />
          <input ref={customerAfterInput} hidden type="file" accept=".xlsx,.xls,.csv,.tsv" onChange={(event) => handleSupplementalFile(event, "after", "customer")} />
        </div>
      </section>
      {pendingImport && <ImportReviewDialog pending={pendingImport} onChange={setPendingImport} onCancel={() => setPendingImport(null)} onConfirm={confirmImport} />}
      {pendingCustomerImport && <CustomerImportDialog pending={pendingCustomerImport} onChange={setPendingCustomerImport} onCancel={() => setPendingCustomerImport(null)} onConfirm={confirmCustomerImport} />}
      {pendingPlacementImport && <PlacementImportDialog pending={pendingPlacementImport} onChange={setPendingPlacementImport} onCancel={() => setPendingPlacementImport(null)} onConfirm={confirmPlacementImport} />}
      {sessionOpen && <SessionDialog records={sessionRecords} onClose={() => setSessionOpen(false)} />}
      {selectedDiff && <DiffDetailDrawer item={selectedDiff} onClose={() => setSelectedDiff(null)} onLocate={(position) => { setSchematicTarget(position); setSelectedDiff(null); setTab("schematic"); }} />}
    </main>
  );
}

function SupplementalFileChip({ label, name, onClick, onClear }: { label: string; name: string; onClick: () => void; onClear: () => void }) {
  return <div className="file-chip supplemental-chip"><button className="file-select" onClick={onClick}><span className="file-icon"><FileSpreadsheet size={17} /></span><span><small>{label}</small><strong>{name || "選擇檔案"}</strong></span></button>{name && <button type="button" className="chip-x" aria-label={`移除${label}`} onClick={onClear}><X size={14} /></button>}</div>;
}

function MvaMetric({ label, before, after }: { label: string; before?: number; after?: number }) {
  const delta = (after ?? 0) - (before ?? 0);
  return <div className="mva-metric"><small>{label}</small><strong>{before ?? "—"} <ArrowRight size={13} /> {after ?? "—"}</strong><span className={delta > 0 ? "increase" : delta < 0 ? "decrease" : ""}>{before == null || after == null ? "待匯入" : delta > 0 ? `+${delta}` : String(delta)}</span></div>;
}

function MvaPanel({ beforeName, afterName, before, after, onBefore, onAfter, onClearBefore, onClearAfter }: {
  beforeName: string; afterName: string; before: MvaSummary | null; after: MvaSummary | null;
  onBefore: () => void; onAfter: () => void; onClearBefore: () => void; onClearAfter: () => void;
}) {
  return <section className="mva-panel" aria-label="MVA 統計">
    <div className="mva-heading"><span><Layers size={18} /></span><div><strong>MVA 製程顆數</strong><small>只計算能對到 BOM 且可判定製程的唯一 Designator</small></div></div>
    <div className="mva-files"><SupplementalFileChip label="舊版 Pick and Place" name={beforeName} onClick={onBefore} onClear={onClearBefore} /><SupplementalFileChip label="新版 Pick and Place" name={afterName} onClick={onAfter} onClear={onClearAfter} /></div>
    <div className="mva-metrics">
      <MvaMetric label="SMT Top" before={before?.smtTop} after={after?.smtTop} />
      <MvaMetric label="SMT Bottom" before={before?.smtBottom} after={after?.smtBottom} />
      <MvaMetric label="DIP Top" before={before?.dipTop} after={after?.dipTop} />
      <MvaMetric label="DIP Bottom" before={before?.dipBottom} after={after?.dipBottom} />
      <div className="mva-excluded"><small>未計入</small><strong>{before?.excluded.length ?? "—"} → {after?.excluded.length ?? "—"}</strong></div>
    </div>
  </section>;
}

function mappingStatusLabel(status: CustomerMappingResult["rows"][number]["status"]) {
  return {
    matched: "配對成功",
    "rd-maintenance-missing": "請 RD 維護",
    "rd-maintenance-mismatch": "R欄料號不一致",
    "location-unmatched": "Location 未匹配",
    "mpn-unmatched": "MPN 未匹配",
    "missing-mpn": "MPN 資料不足",
    ambiguous: "多重候選",
  }[status];
}

function CustomerMappingTable({ version, result }: { version: "舊版" | "新版"; result: CustomerMappingResult | null }) {
  if (!result) return <div className="customer-empty">{version}客戶 BOM 尚未匯入</div>;
  return <div className="customer-table-wrap"><table className="customer-table">
    <thead><tr><th>狀態</th><th>客戶 TPN</th><th>Location</th><th>客戶 MPN</th><th>公司主替料 MPN</th><th>說明</th></tr></thead>
    <tbody>{result.rows.map((row) => <tr className={row.status} key={`${version}-${row.record.sourceRow}-${row.record.customerPartNumber}`}>
      <td><span className={`mapping-status ${row.status}`}>{mappingStatusLabel(row.status)}</span></td>
      <td><strong>{row.record.customerPartNumber || "—"}</strong><small>Excel 第 {row.record.sourceRow} 列</small></td>
      <td>{row.record.positions.join("、") || "—"}</td>
      <td className={row.status === "mpn-unmatched" || row.status === "missing-mpn" ? "mismatch" : ""}>{row.record.manufacturerParts.join("\n") || "—"}</td>
      <td className={row.status === "mpn-unmatched" || row.status === "missing-mpn" ? "mismatch" : ""}>{row.companyManufacturerParts.join("\n") || "—"}{row.companyItem && <small><b>BOM R欄</b> {row.companyItem.rdCustomerPartNumbers?.join("、") || "空白"}</small>}</td>
      <td>{row.reason}</td>
    </tr>)}</tbody>
  </table></div>;
}

function CustomerMappingPanel({ beforeName, afterName, before, after, onBefore, onAfter, onClearBefore, onClearAfter }: {
  beforeName: string; afterName: string; before: CustomerMappingResult | null; after: CustomerMappingResult | null;
  onBefore: () => void; onAfter: () => void; onClearBefore: () => void; onClearAfter: () => void;
}) {
  const [version, setVersion] = useState<"before" | "after">("before");
  const result = version === "before" ? before : after;
  return <div className="customer-panel">
    <div className="customer-import-bar">
      <div><strong>客戶 BOM TPN 嚴格對應</strong><small>Location 集合完全相同，且公司每一顆主替料 MPN 都必須逐字匹配</small></div>
      <div className="customer-files"><SupplementalFileChip label="舊版客戶 BOM" name={beforeName} onClick={onBefore} onClear={onClearBefore} /><SupplementalFileChip label="新版客戶 BOM" name={afterName} onClick={onAfter} onClear={onClearAfter} /></div>
    </div>
    <div className="customer-summary">
      {(["before", "after"] as const).map((side) => {
        const current = side === "before" ? before : after;
        return <button type="button" className={version === side ? "active" : ""} onClick={() => setVersion(side)} key={side}>
          <strong>{side === "before" ? "舊版" : "新版"}</strong>
          <span className="ok">成功 {current?.counts.matched ?? 0}</span>
          <span className="bad">需處理 {(current?.rows.length ?? 0) - (current?.counts.matched ?? 0)}</span>
        </button>;
      })}
    </div>
    <CustomerMappingTable version={version === "before" ? "舊版" : "新版"} result={result} />
  </div>;
}

function DiffDetailDrawer({ item, onClose, onLocate }: { item: BomDiff; onClose: () => void; onLocate: (position: string) => void }) {
  const structure = diffStructureLabel(item) || "未標示架構";
  const positions = [...new Set([...(item.replacementPositions ?? []), ...(item.addedPositions ?? []), ...(item.removedPositions ?? []), ...(item.after?.positions ?? []), ...(item.before?.positions ?? [])])];
  return <div className="detail-backdrop" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}>
    <aside className="detail-drawer" role="dialog" aria-modal="true" aria-labelledby="diff-detail-title">
      <header>
        <div><span className={`primary-badge ${primaryTypeTone(item.primaryType)}`}>{primaryTypeLabel(item.primaryType)}</span><small>差異群組詳細資料</small></div>
        <button type="button" onClick={onClose} aria-label="關閉詳細資料"><X size={18} /></button>
      </header>
      <div className="detail-content">
        <section className="detail-hero">
          <small>料號異動</small>
          <h2 id="diff-detail-title">{displayPart(item.before)} <ArrowRight size={17} /> {displayPart(item.after)}</h2>
          <ChangeFields fields={bomDiffDisplayFields(item)} tone={primaryTypeTone(item.primaryType)} />
        </section>
        <div className="detail-metrics">
          <div><small>舊版數量</small><strong>{item.before?.qty ?? 0}</strong></div>
          <span><ArrowRight size={16} /></span>
          <div><small>新版數量</small><strong>{item.after?.qty ?? 0}</strong></div>
        </div>
        <section className="detail-section"><h3>所屬架構</h3><p className="structure-value">{structure}</p></section>
        {item.processChange && <section className="detail-section review-reason"><h3><AlertTriangle size={15} /> 製程別放置異常</h3><p>此料號由 {item.processChange.before} 架構移至 {item.processChange.after} 架構，請確認 RD 是否將料件放錯群組。</p></section>}
        <section className="detail-section"><h3>料號與製造廠商</h3>
          <div className="detail-version-grid">
            <article><span>舊版</span><PartList item={item.before} changedParts={item.removedParts} tone="removed" /></article>
            <article><span>新版</span><PartList item={item.after} changedParts={item.addedParts} tone="added" /></article>
          </div>
        </section>
        <section className="detail-section"><h3>插件位置</h3>
          {positions.length ? <div className="detail-positions">{positions.map((position) => <button type="button" key={position} onClick={() => onLocate(position)} title={`在線路圖定位 ${position}`}><CircuitBoard size={14} />{position}</button>)}</div> : <p className="detail-empty">沒有插件位置資料</p>}
        </section>
        {(item.needsReview || item.matchConfidence === "low") && <section className="detail-section review-reason"><h3><AlertTriangle size={15} /> 待人工確認</h3><p>{item.matchReason}</p></section>}
      </div>
      <footer><button type="button" className="secondary" onClick={onClose}>關閉</button>{positions[0] && <button type="button" className="primary" onClick={() => onLocate(positions[0])}><CircuitBoard size={15} /> 定位第一個插件位置</button>}</footer>
    </aside>
  </div>;
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

function CustomerImportDialog({ pending, onChange, onCancel, onConfirm }: { pending: PendingCustomerImport; onChange: (value: PendingCustomerImport) => void; onCancel: () => void; onConfirm: () => void }) {
  const sheet = pending.sheets[pending.sheetIndex];
  const header = sheet?.matrix[pending.headerIndex] ?? [];
  const records = useMemo(() => parseCustomerBomMatrix(sheet?.matrix ?? [], pending.headerIndex, pending.mapping), [sheet, pending.headerIndex, pending.mapping]);
  const keys = Object.keys(customerColumnLabels) as CustomerBomColumnKey[];
  const changeSheet = (sheetIndex: number) => {
    const matrix = pending.sheets[sheetIndex].matrix;
    const found = findCustomerHeader(matrix);
    const headerIndex = found >= 0 ? found : Math.max(0, matrix.findIndex((row) => row.filter((cell) => String(cell ?? "").trim()).length >= 2));
    onChange({ ...pending, sheetIndex, headerIndex, mapping: detectCustomerColumns(matrix[headerIndex] ?? []) });
  };
  const changeHeader = (headerIndex: number) => onChange({ ...pending, headerIndex, mapping: detectCustomerColumns(sheet.matrix[headerIndex] ?? []) });
  return <div className="modal-backdrop" role="presentation"><section className="import-dialog" role="dialog" aria-modal="true" aria-labelledby="customer-import-title">
    <header><div><small>客戶 BOM TPN 欄位設定</small><h2 id="customer-import-title">{pending.fileName}</h2></div><button aria-label="關閉" onClick={onCancel}><X size={19} /></button></header>
    <div className="import-controls"><label>工作表<select value={pending.sheetIndex} onChange={(event) => changeSheet(Number(event.target.value))}>{pending.sheets.map((item, index) => <option key={item.name} value={index}>{item.name}</option>)}</select></label><label>標題列<select value={pending.headerIndex} onChange={(event) => changeHeader(Number(event.target.value))}>{sheet.matrix.slice(0, 40).map((row, index) => <option key={index} value={index}>第 {index + 1} 列｜{row.filter((cell) => String(cell ?? "").trim()).slice(0, 4).join("、") || "空白"}</option>)}</select></label></div>
    <div className="mapping-grid">{keys.map((key) => <label key={key}><span>{customerColumnLabels[key]}<b>*</b></span><select value={pending.mapping[key] ?? ""} onChange={(event) => onChange({ ...pending, mapping: { ...pending.mapping, [key]: event.target.value === "" ? undefined : Number(event.target.value) } })}><option value="">未指定</option>{header.map((cell, index) => <option key={index} value={index}>{index + 1}. {String(cell || `未命名欄位 ${index + 1}`)}</option>)}</select></label>)}</div>
    <div className="audit-cards"><div><small>TPN 資料列</small><strong>{records.length}</strong></div><div><small>含 Location</small><strong>{records.filter((record) => record.positions.length).length}</strong></div><div><small>含 MPN</small><strong>{records.filter((record) => record.manufacturerParts.length).length}</strong></div></div>
    <div className="issue-list"><p className="issue-ok"><ShieldCheck size={17} /> MPN 僅忽略大小寫及儲存格前後空白；符號、內部空白與尾碼必須完全一致。</p></div>
    <footer><button className="secondary" onClick={onCancel}>取消</button><button className="primary" disabled={!isCompleteCustomerMapping(pending.mapping)} onClick={onConfirm}>確認匯入</button></footer>
  </section></div>;
}

function PlacementImportDialog({ pending, onChange, onCancel, onConfirm }: { pending: PendingPlacementImport; onChange: (value: PendingPlacementImport) => void; onCancel: () => void; onConfirm: () => void }) {
  const sheet = pending.sheets[pending.sheetIndex];
  const header = sheet?.matrix[pending.headerIndex] ?? [];
  const records = useMemo(() => parsePlacementMatrix(sheet?.matrix ?? [], pending.headerIndex, pending.mapping), [sheet, pending.headerIndex, pending.mapping]);
  const keys = Object.keys(placementColumnLabels) as PlacementColumnKey[];
  const changeSheet = (sheetIndex: number) => {
    const matrix = pending.sheets[sheetIndex].matrix;
    const found = findPlacementHeader(matrix);
    const headerIndex = found >= 0 ? found : Math.max(0, matrix.findIndex((row) => row.filter((cell) => String(cell ?? "").trim()).length >= 2));
    onChange({ ...pending, sheetIndex, headerIndex, mapping: detectPlacementColumns(matrix[headerIndex] ?? []) });
  };
  const changeHeader = (headerIndex: number) => onChange({ ...pending, headerIndex, mapping: detectPlacementColumns(sheet.matrix[headerIndex] ?? []) });
  return <div className="modal-backdrop" role="presentation"><section className="import-dialog" role="dialog" aria-modal="true" aria-labelledby="placement-import-title">
    <header><div><small>Pick and Place 欄位設定</small><h2 id="placement-import-title">{pending.fileName}</h2></div><button aria-label="關閉" onClick={onCancel}><X size={19} /></button></header>
    <div className="import-controls"><label>工作表<select value={pending.sheetIndex} onChange={(event) => changeSheet(Number(event.target.value))}>{pending.sheets.map((item, index) => <option key={item.name} value={index}>{item.name}</option>)}</select></label><label>標題列<select value={pending.headerIndex} onChange={(event) => changeHeader(Number(event.target.value))}>{sheet.matrix.slice(0, 50).map((row, index) => <option key={index} value={index}>第 {index + 1} 列｜{row.filter((cell) => String(cell ?? "").trim()).slice(0, 4).join("、") || "空白"}</option>)}</select></label></div>
    <div className="mapping-grid">{keys.map((key) => <label key={key}><span>{placementColumnLabels[key]}<b>*</b></span><select value={pending.mapping[key] ?? ""} onChange={(event) => onChange({ ...pending, mapping: { ...pending.mapping, [key]: event.target.value === "" ? undefined : Number(event.target.value) } })}><option value="">未指定</option>{header.map((cell, index) => <option key={index} value={index}>{index + 1}. {String(cell || `未命名欄位 ${index + 1}`)}</option>)}</select></label>)}</div>
    <div className="audit-cards"><div><small>Designator</small><strong>{records.length}</strong></div><div><small>Top</small><strong>{records.filter((record) => record.side === "Top").length}</strong></div><div><small>Bottom</small><strong>{records.filter((record) => record.side === "Bottom").length}</strong></div><div className={records.some((record) => !record.side) ? "warning" : "ok"}><small>Layer 無法辨識</small><strong>{records.filter((record) => !record.side).length}</strong></div></div>
    <div className="issue-list"><p className="issue-ok"><ShieldCheck size={17} /> 匯入後只計算能對到 BOM 且可判定 SMT／DIP 的唯一插件位置。</p></div>
    <footer><button className="secondary" onClick={onCancel}>取消</button><button className="primary" disabled={!isCompletePlacementMapping(pending.mapping)} onClick={onConfirm}>確認匯入</button></footer>
  </section></div>;
}

function SessionDialog({ records, onClose }: { records: ImportRecord[]; onClose: () => void }) {
  return <div className="modal-backdrop" role="presentation"><section className="session-dialog" role="dialog" aria-modal="true" aria-labelledby="session-title"><header><div><small>只保留在目前頁面記憶體</small><h2 id="session-title">此次工作階段</h2></div><button aria-label="關閉" onClick={onClose}><X size={19} /></button></header>{records.length ? <div className="session-list">{records.map((record, index) => <article key={`${record.importedAt}-${index}`}><FileSpreadsheet size={18} /><div><strong>{record.fileName}</strong><small>{record.sheetName}・{record.importedAt}・{record.audit.groupCount} 組料・{record.audit.issues.length} 項提示</small></div></article>)}</div> : <div className="empty-state">目前還沒有匯入紀錄</div>}<footer><button className="primary" onClick={onClose}>完成</button></footer></section></div>;
}

function PrimaryTypeBadge({ item }: { item: BomDiff }) {
  const showConfidence = item.matchConfidence === "low";
  const label = "低可信";
  const structure = diffStructureLabel(item);
  const confidenceLabel = `${item.needsReview ? "待確認，" : ""}${label}。${structure ? `所屬架構：${structure}。` : ""}配對依據：${item.matchReason}`;
  return <div className="primary-stack">
    <span className={`primary-badge ${primaryTypeTone(item.primaryType)}`}>{primaryTypeLabel(item.primaryType)}</span>
    {showConfidence && <span className="confidence-popover">
      <button type="button" className={`confidence ${item.matchConfidence}`} aria-label={confidenceLabel}>{item.needsReview ? "人工待審核 · " : ""}{label}</button>
      <span className="confidence-tooltip" role="tooltip">
        {structure && <span><b>所屬架構</b>{structure}</span>}
        <span><b>配對依據</b>{item.matchReason}</span>
        {item.needsReview && <em>此群組需要人工確認</em>}
      </span>
    </span>}
  </div>;
}

function ChangeFields({ fields, tone }: { fields: string[]; tone: string }) {
  return <div className={`change-list ${tone}`}>{fields.map((field) =>
    <span className={field.startsWith("製程別放置異常") ? "change-pill process-warning" : "change-pill"} key={field}>{field}</span>,
  )}</div>;
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
      const label = item.primaryType === "substituteAdded" ? "替料" : globallyNew ? "完全新料" : item.replacementPositions.length ? "改用" : item.before ? "替料" : "新增料號";
      return <div className="b-diff added" key={`a-${part.part}`}><span>{label}</span><strong>+ {alternativeLabel(part)}</strong></div>;
    })}
    {item.removedParts.map((part) => {
      const globallyDeleted = deletedKeys.has(canonicalPartNumber(part.part).toUpperCase());
      const label = item.primaryType === "substituteRemoved" ? "替料" : globallyDeleted ? "完全移除" : item.replacementPositions.length ? "停用" : item.after ? "替料" : "刪除料號";
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
      <div><strong>{alternative.part || "—"}</strong><small><b>製造廠商料號</b> {alternative.manufacturerPart || "—"}</small><small><b>製造廠商</b> {alternative.manufacturerName || "—"}</small>{index === 0 && <CustomerPartStatus item={item} />}</div>
    </div>;
  })}</div>;
}

function CustomerPartStatus({ item }: { item: BomItem }) {
  const rdValues = item.rdCustomerPartNumbers ?? item.alternatives.flatMap((alternative) => alternative.rdCustomerPartNumbers ?? []);
  const uniqueRdValues = [...new Set(rdValues)];
  if (item.customerMappingStatus === "matched" && item.customerPartNumber) return <>
    <small className="customer-part matched"><b>BOM R欄 TPN</b> {uniqueRdValues.join("、") || "—"}</small>
    <small className="customer-part matched"><b>客戶 BOM TPN 驗證</b> {item.customerPartNumber}・一致</small>
  </>;
  if ((item.customerMappingStatus === "rd-maintenance-missing" || item.customerMappingStatus === "rd-maintenance-mismatch") && item.customerPartNumber) return <>
    <small className="customer-part rd-warning"><b>BOM R欄 TPN</b> {uniqueRdValues.join("、") || "空白"}</small>
    <small className="customer-part rd-warning"><b>客戶 BOM TPN 驗證</b> {item.customerPartNumber}・請 RD 維護</small>
  </>;
  const labels = {
    "rd-maintenance-missing": "請 RD 維護",
    "rd-maintenance-mismatch": "R欄料號不一致",
    "location-unmatched": "Location 未匹配",
    "mpn-unmatched": "MPN 未匹配",
    "missing-mpn": "MPN 資料不足",
    ambiguous: "多重候選",
    "not-imported": "未匯入",
  } as const;
  const status = item.customerMappingStatus ?? "not-imported";
  if (status === "matched") return <small className="customer-part missing-mpn"><b>客戶 TPN</b> 配對結果缺少 TPN</small>;
  return <small className={`customer-part ${status}`} title={item.customerMappingReason}><b>BOM R欄 TPN</b> {uniqueRdValues.join("、") || "空白"}・{labels[status]}</small>;
}

function PositionSummary({ added, removed, replacement, allPositions, onLocate }: { added: string[]; removed: string[]; replacement: string[]; allPositions: string[]; onLocate: (position: string) => void }) {
  if (!added.length && !removed.length && !replacement.length) return allPositions.length ? <div className="position-list unchanged"><small>位置未變・點擊定位</small>{allPositions.map((position) => <button className="position-chip neutral" title={`在線路圖定位 ${position}`} onClick={(event) => { event.stopPropagation(); onLocate(position); }} key={`n-${position}`}>{position}</button>)}</div> : <span className="no-change">—</span>;
  return <div className="position-list">
    {replacement.map((position) => <button className="position-chip changed" title={`在線路圖定位 ${position}`} onClick={(event) => { event.stopPropagation(); onLocate(position); }} key={`c-${position}`}>{position} 換料</button>)}
    {added.map((position) => <button className="position-chip added" title={`在線路圖定位 ${position}`} onClick={(event) => { event.stopPropagation(); onLocate(position); }} key={`a-${position}`}>+ {position}</button>)}
    {removed.map((position) => <button className="position-chip removed" title={`在線路圖定位 ${position}`} onClick={(event) => { event.stopPropagation(); onLocate(position); }} key={`r-${position}`}>− {position}</button>)}
  </div>;
}

function SchematicPanel({ beforeUrl, afterUrl, beforeFile, afterFile, beforeName, afterName, target, onTargetChange, diffImage, busy, diffs, onFile }: { beforeUrl: string | null; afterUrl: string | null; beforeFile: File | null; afterFile: File | null; beforeName: string; afterName: string; target: string; onTargetChange: (value: string) => void; diffImage: string | null; busy: boolean; diffs: BomDiff[]; onFile: (event: ChangeEvent<HTMLInputElement>, side: "before" | "after") => void }) {
  const [view, setView] = useState<"side" | "diff">("side");
  const [syncEnabled, setSyncEnabled] = useState(true);
  const [expanded, setExpanded] = useState(false);
  const [reportBusy, setReportBusy] = useState(false);
  const [reportProgress, setReportProgress] = useState("");
  const [scales, setScales] = useState<Record<PdfViewerSide, number>>({ before: 1.15, after: 1.15 });
  const [syncState, setSyncState] = useState<PdfScrollSync>({ source: "before", x: 0, y: 0, revision: 0 });
  const isPdf = (name: string) => name.toLowerCase().endsWith(".pdf");
  const changeScale = (side: PdfViewerSide, scale: number) => setScales((current) => syncEnabled ? { before: scale, after: scale } : { ...current, [side]: scale });
  const toggleSync = () => {
    const next = !syncEnabled;
    if (next) setScales((current) => ({ before: current.before, after: current.before }));
    setSyncEnabled(next);
  };
  const canExportReport = Boolean(beforeFile && afterFile && isPdf(beforeName) && isPdf(afterName) && diffs.length);
  const exportReport = async () => {
    if (!beforeFile || !afterFile || !canExportReport || reportBusy) return;
    setReportBusy(true);
    setReportProgress("準備線路圖報告…");
    try {
      const { exportSchematicPdfReport } = await import("./schematic-report");
      await exportSchematicPdfReport({
        beforeFile,
        afterFile,
        beforeName,
        afterName,
        diffs,
        onProgress: ({ current, total, message }) => setReportProgress(`${message}${total ? ` ${current}／${total}` : ""}`),
      });
      setReportProgress("線路圖差異報告已下載");
    } catch (error) {
      setReportProgress("");
      window.alert(error instanceof Error ? error.message : "線路圖報告產生失敗，請重新整理後再試一次。");
    } finally {
      setReportBusy(false);
    }
  };
  useEffect(() => {
    if (!expanded) return;
    const closeOnEscape = (event: KeyboardEvent) => { if (event.key === "Escape") setExpanded(false); };
    window.addEventListener("keydown", closeOnEscape);
    return () => window.removeEventListener("keydown", closeOnEscape);
  }, [expanded]);
  const preview = (url: string | null, file: File | null, name: string, sideLabel: string, side: PdfViewerSide) => !url ? <div className="sheet-empty"><ImageIcon size={28} /><strong>選擇線路圖</strong><small>文字定位支援含文字層的 PDF</small></div> : isPdf(name) && file ? <PdfSchematicViewer file={file} target={target} sideLabel={sideLabel} side={side} scale={scales[side]} syncEnabled={syncEnabled} syncState={syncState} onScaleChange={changeScale} onSyncScroll={(source, x, y) => setSyncState((current) => ({ source, x, y, revision: current.revision + 1 }))} /> : <div className="image-preview"><img src={url} alt={name} />{target && <span>圖片格式無法搜尋 {target}；請使用含文字層的 PDF。</span>}</div>;
  return <div className={expanded ? "schematic-panel expanded" : "schematic-panel"}>
    <div className="schematic-toolbar"><div><strong>線路圖與 BOM 連動</strong><small>支援同步縮放／捲動、旋轉文字與本機索引快取</small></div><label className="schematic-search"><Search size={15} /><input value={target} onChange={(event) => onTargetChange(event.target.value.toUpperCase())} placeholder="輸入插件位置，例如 U45" />{target && <button aria-label="清除搜尋" onClick={() => onTargetChange("")}><X size={14} /></button>}</label><div className="schematic-options"><button className="schematic-export" disabled={!canExportReport || reportBusy} onClick={exportReport} title={!beforeFile || !afterFile ? "請先匯入新舊版線路圖" : !isPdf(beforeName) || !isPdf(afterName) ? "線路圖報告目前需要新舊版皆為 PDF" : !diffs.length ? "目前沒有可匯出的 BOM 差異" : "依 BOM 異動插件位置匯出 PDF 報告"}><Download size={14} />{reportBusy ? "產生報告中" : "匯出線路圖報告"}</button><button className={syncEnabled ? "sync-active" : ""} onClick={toggleSync}>{syncEnabled ? "同步檢視中" : "獨立檢視"}</button><button onClick={clearInactivePdfCache}>清除閒置快取</button><button className={expanded ? "expand-active" : ""} onClick={() => setExpanded((current) => !current)}>{expanded ? <Minimize2 size={14} /> : <Maximize2 size={14} />}{expanded ? "退出放大" : "放大顯示"}</button><div className="view-toggle"><button className={view === "side" ? "active" : ""} onClick={() => setView("side")}>並排定位</button><button className={view === "diff" ? "active" : ""} onClick={() => setView("diff")} disabled={!diffImage}>像素差異</button></div></div></div>
    {reportProgress && <div className={reportBusy ? "schematic-report-progress active" : "schematic-report-progress"}><span>{reportProgress}</span>{reportBusy && <i />}</div>}
    {view === "diff" && diffImage ? <div className="diff-canvas"><div className="diff-note"><span /> 紅色區域代表舊版與新版的像素差異</div><img src={diffImage} alt="線路圖像素差異" /></div> : <div className="sheet-grid">{(["before", "after"] as const).map((side) => { const url = side === "before" ? beforeUrl : afterUrl; const file = side === "before" ? beforeFile : afterFile; const name = side === "before" ? beforeName : afterName; const sideLabel = side === "before" ? "舊版" : "新版"; return <div className="sheet-card" key={side}><div className="sheet-head"><span>{sideLabel}線路圖</span><strong>{name || "尚未選擇檔案"}</strong><label className="sheet-upload" title={`選擇${sideLabel}線路圖`}><input type="file" hidden accept="image/png,image/jpeg,image/webp,application/pdf" onChange={(e) => onFile(e, side)} /><UploadCloud size={17} /></label></div><div className="sheet-preview">{preview(url, file, name, sideLabel, side)}</div></div>; })}</div>}
    {busy && <div className="processing"><span /> 正在產生差異圖…</div>}
  </div>;
}
