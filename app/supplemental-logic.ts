import { bomProcessKind, type BomItem } from "./bom-logic.ts";

export type CustomerBomColumnKey = "customerPartNumber" | "manufacturerParts" | "positions";
export type CustomerBomColumnMapping = Partial<Record<CustomerBomColumnKey, number>>;
export type CustomerMappingStatus = "matched" | "rd-maintenance-missing" | "rd-maintenance-mismatch" | "location-unmatched" | "mpn-unmatched" | "missing-mpn" | "ambiguous" | "not-imported";

export type CustomerBomRecord = {
  customerPartNumber: string;
  manufacturerParts: string[];
  positions: string[];
  sourceRow: number;
};

export type CustomerMappingRow = {
  record: CustomerBomRecord;
  status: Exclude<CustomerMappingStatus, "not-imported">;
  reason: string;
  companyItem?: BomItem;
  companyManufacturerParts: string[];
};

export type CustomerMappingResult = {
  items: BomItem[];
  rows: CustomerMappingRow[];
  counts: Record<Exclude<CustomerMappingStatus, "not-imported">, number>;
};

export type PlacementColumnKey = "designator" | "layer";
export type PlacementColumnMapping = Partial<Record<PlacementColumnKey, number>>;
export type PlacementSide = "Top" | "Bottom";
export type PlacementRecord = {
  designator: string;
  side?: PlacementSide;
  rawLayer: string;
  sourceRow: number;
};

export type MvaIncludedRecord = {
  designator: string;
  side: PlacementSide;
  process: "SMT" | "DIP";
  part: string;
  structure: string;
};

export type MvaExcludedRecord = {
  designator: string;
  rawLayer: string;
  sourceRows: number[];
  reason: string;
};

export type MvaSummary = {
  smtTop: number;
  smtBottom: number;
  dipTop: number;
  dipBottom: number;
  included: MvaIncludedRecord[];
  excluded: MvaExcludedRecord[];
};

const customerAliases: Record<CustomerBomColumnKey, string[]> = {
  customerPartNumber: ["partnumber", "part number", "customer part number", "customer pn", "customer tpn", "tpn", "客戶料號", "客戶 tpn"],
  manufacturerParts: ["mfgpnos", "mfg pnos", "mpn", "manufacturer part number", "製造廠商料號", "製造商料號"],
  positions: ["referencedesignator", "reference designator", "refdes", "location", "locations", "插件位置", "位號"],
};

const placementAliases: Record<PlacementColumnKey, string[]> = {
  designator: ["designator", "reference designator", "refdes", "location", "插件位置", "位號"],
  layer: ["layer", "side", "board side", "板面", "面別"],
};

export const customerColumnLabels: Record<CustomerBomColumnKey, string> = {
  customerPartNumber: "客戶 TPN",
  manufacturerParts: "客戶製造廠商料號",
  positions: "客戶插件位置",
};

export const placementColumnLabels: Record<PlacementColumnKey, string> = {
  designator: "插件位置／Designator",
  layer: "板面／Layer",
};

function text(value: unknown) {
  return value == null ? "" : String(value).trim();
}

function headerKey(value: unknown) {
  return text(value).toLowerCase().replace(/[：:()（）_\-]/g, " ").replace(/\s+/g, " ").trim();
}

function compactHeaderKey(value: unknown) {
  return headerKey(value).replace(/\s+/g, "");
}

function findColumn(header: unknown[], aliases: string[]) {
  const normalized = new Set(aliases.flatMap((alias) => [headerKey(alias), compactHeaderKey(alias)]));
  return header.findIndex((cell) => normalized.has(headerKey(cell)) || normalized.has(compactHeaderKey(cell)));
}

export function detectCustomerColumns(header: unknown[]): CustomerBomColumnMapping {
  const customerPartNumber = findColumn(header, customerAliases.customerPartNumber);
  const manufacturerParts = findColumn(header, customerAliases.manufacturerParts);
  const positions = findColumn(header, customerAliases.positions);
  return {
    ...(customerPartNumber >= 0 ? { customerPartNumber } : {}),
    ...(manufacturerParts >= 0 ? { manufacturerParts } : {}),
    ...(positions >= 0 ? { positions } : {}),
  };
}

export function detectPlacementColumns(header: unknown[]): PlacementColumnMapping {
  const designator = findColumn(header, placementAliases.designator);
  const layer = findColumn(header, placementAliases.layer);
  return {
    ...(designator >= 0 ? { designator } : {}),
    ...(layer >= 0 ? { layer } : {}),
  };
}

export function findCustomerHeader(matrix: unknown[][]) {
  return matrix.findIndex((row) => {
    const mapping = detectCustomerColumns(row);
    return mapping.customerPartNumber != null && mapping.positions != null;
  });
}

export function findPlacementHeader(matrix: unknown[][]) {
  return matrix.findIndex((row) => {
    const mapping = detectPlacementColumns(row);
    return mapping.designator != null && mapping.layer != null;
  });
}

export function normalizeLocation(value: unknown) {
  return text(value).toUpperCase();
}

export function parseCustomerPositions(value: unknown) {
  return [...new Set(text(value)
    .split(/[|,，;；\s]+/)
    .map(normalizeLocation)
    .filter(Boolean))]
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true, sensitivity: "base" }));
}

export function normalizeMpn(value: unknown) {
  return text(value).toUpperCase();
}

export function parseCustomerMpns(value: unknown) {
  return [...new Set(text(value)
    .split(/[|,，;；\r\n]+/)
    .map(normalizeMpn)
    .filter(Boolean))];
}

export function isCompleteCustomerMapping(mapping: CustomerBomColumnMapping) {
  return mapping.customerPartNumber != null && mapping.manufacturerParts != null && mapping.positions != null;
}

export function isCompletePlacementMapping(mapping: PlacementColumnMapping) {
  return mapping.designator != null && mapping.layer != null;
}

export function parseCustomerBomMatrix(matrix: unknown[][], headerIndex: number, mapping: CustomerBomColumnMapping) {
  if (!isCompleteCustomerMapping(mapping)) return [];
  const customerPartColumn = mapping.customerPartNumber!;
  const manufacturerPartColumn = mapping.manufacturerParts!;
  const positionsColumn = mapping.positions!;
  return matrix.slice(headerIndex + 1).flatMap((row, index): CustomerBomRecord[] => {
    const customerPartNumber = text(row[customerPartColumn]);
    const positions = parseCustomerPositions(row[positionsColumn]);
    if (!customerPartNumber && !positions.length) return [];
    return [{
      customerPartNumber,
      manufacturerParts: parseCustomerMpns(row[manufacturerPartColumn]),
      positions,
      sourceRow: headerIndex + index + 2,
    }];
  });
}

function positionKey(positions: string[]) {
  return [...new Set(positions.map(normalizeLocation).filter(Boolean))]
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true, sensitivity: "base" }))
    .join("|");
}

function companyMpns(item: BomItem) {
  return [...new Set(item.alternatives.map((alternative) => normalizeMpn(alternative.manufacturerPart)).filter(Boolean))];
}

function rdCustomerNumbers(item: BomItem) {
  return [...new Set((item.rdCustomerPartNumbers ?? item.alternatives.flatMap((alternative) => alternative.rdCustomerPartNumbers ?? []))
    .map((partNumber) => partNumber.trim().toUpperCase())
    .filter(Boolean))];
}

function maintainedStatus(item: BomItem, record: CustomerBomRecord) {
  const maintained = rdCustomerNumbers(item);
  const expected = record.customerPartNumber.trim().toUpperCase();
  if (!maintained.length) return "rd-maintenance-missing" as const;
  if (!expected || !maintained.includes(expected)) return "rd-maintenance-mismatch" as const;
  return "matched" as const;
}

function allCompanyMpnsMatch(item: BomItem, record: CustomerBomRecord) {
  const company = companyMpns(item);
  const customer = new Set(record.manufacturerParts.map(normalizeMpn));
  return company.length > 0 && record.manufacturerParts.length > 0 && company.every((mpn) => customer.has(mpn));
}

function emptyCounts(): CustomerMappingResult["counts"] {
  return { matched: 0, "rd-maintenance-missing": 0, "rd-maintenance-mismatch": 0, "location-unmatched": 0, "mpn-unmatched": 0, "missing-mpn": 0, ambiguous: 0 };
}

export function mapCustomerBom(items: BomItem[], records: CustomerBomRecord[]): CustomerMappingResult {
  const locationCandidates = records.map((record) => items
    .map((item, index) => ({ item, index }))
    .filter(({ item }) => record.positions.length > 0 && positionKey(item.positions) === positionKey(record.positions)));
  const passingCandidates = locationCandidates.map((candidates, recordIndex) =>
    candidates.filter(({ item }) => allCompanyMpnsMatch(item, records[recordIndex])),
  );
  const companyPassingCounts = new Map<number, number>();
  passingCandidates.forEach((candidates) => candidates.forEach(({ index }) =>
    companyPassingCounts.set(index, (companyPassingCounts.get(index) ?? 0) + 1),
  ));

  const matchedByCompany = new Map<number, CustomerBomRecord>();
  const rows = records.map((record, recordIndex): CustomerMappingRow => {
    const candidates = locationCandidates[recordIndex];
    const passing = passingCandidates[recordIndex];
    const hasMissingMpn = candidates.some(({ item }) => companyMpns(item).length === 0) || record.manufacturerParts.length === 0;
    if (!candidates.length) {
      return { record, status: "location-unmatched", reason: "找不到 Location 集合完全相同的公司料群。", companyManufacturerParts: [] };
    }
    if (!passing.length) {
      const candidate = candidates[0]?.item;
      return {
        record,
        status: hasMissingMpn ? "missing-mpn" : "mpn-unmatched",
        reason: hasMissingMpn ? "客戶或公司 MPN 資料不完整。" : "Location 相同，但公司主替料 MPN 未全部完全匹配。",
        companyItem: candidates.length === 1 ? candidate : undefined,
        companyManufacturerParts: candidates.flatMap(({ item }) => companyMpns(item)),
      };
    }
    const unique = passing.length === 1 && companyPassingCounts.get(passing[0].index) === 1;
    if (!unique) {
      return {
        record,
        status: "ambiguous",
        reason: "Location 與 MPN 同時命中多個候選，為避免誤配不帶入 TPN。",
        companyManufacturerParts: passing.flatMap(({ item }) => companyMpns(item)),
      };
    }
    const match = passing[0];
    matchedByCompany.set(match.index, record);
    const maintenanceStatus = maintainedStatus(match.item, record);
    return {
      record,
      status: maintenanceStatus,
      reason: maintenanceStatus === "matched"
        ? "Location、MPN 與 BOM R欄 TPN 均完全匹配。"
        : maintenanceStatus === "rd-maintenance-missing"
          ? "Location 與 MPN 已匹配，但 BOM R欄沒有維護此 TPN，請 RD 維護。"
          : "Location 與 MPN 已匹配，但客戶 BOM TPN 不在 BOM R欄清單中，請 RD 確認並維護。",
      companyItem: match.item,
      companyManufacturerParts: companyMpns(match.item),
    };
  });

  const mappedItems = items.map((item, index) => {
    const matched = matchedByCompany.get(index);
    if (matched) {
      const maintenanceStatus = maintainedStatus(item, matched);
      const reason = maintenanceStatus === "matched"
        ? "Location、MPN 與 BOM R欄 TPN 完全匹配"
        : maintenanceStatus === "rd-maintenance-missing"
          ? "Location 與 MPN 已匹配，但 BOM R欄空白，請 RD 維護"
          : "Location 與 MPN 已匹配，但 TPN 不在 BOM R欄，請 RD 維護";
      return { ...item, customerPartNumber: matched.customerPartNumber, customerMappingStatus: maintenanceStatus, customerMappingReason: reason };
    }
    const sameLocation = records.filter((record) => record.positions.length > 0 && positionKey(record.positions) === positionKey(item.positions));
    if (!sameLocation.length) {
      return { ...item, customerPartNumber: undefined, customerMappingStatus: "location-unmatched" as const, customerMappingReason: "客戶 BOM 找不到相同 Location 集合" };
    }
    if (!companyMpns(item).length || sameLocation.some((record) => !record.manufacturerParts.length)) {
      return { ...item, customerPartNumber: undefined, customerMappingStatus: "missing-mpn" as const, customerMappingReason: "客戶或公司 MPN 資料不完整" };
    }
    const matchingRows = sameLocation.filter((record) => allCompanyMpnsMatch(item, record));
    if (matchingRows.length > 1) {
      return { ...item, customerPartNumber: undefined, customerMappingStatus: "ambiguous" as const, customerMappingReason: "多個客戶 BOM 列同時完全匹配" };
    }
    return { ...item, customerPartNumber: undefined, customerMappingStatus: "mpn-unmatched" as const, customerMappingReason: "Location 相同，但 MPN 未完全匹配" };
  });

  const counts = emptyCounts();
  rows.forEach((row) => { counts[row.status] += 1; });
  return { items: mappedItems, rows, counts };
}

export function normalizePlacementSide(value: unknown): PlacementSide | undefined {
  const normalized = text(value).toLowerCase().replace(/[\s_-]+/g, "");
  if (normalized === "top" || normalized === "toplayer") return "Top";
  if (normalized === "bottom" || normalized === "bottomlayer" || normalized === "bot" || normalized === "botlayer") return "Bottom";
  return undefined;
}

export function parsePlacementMatrix(matrix: unknown[][], headerIndex: number, mapping: PlacementColumnMapping) {
  if (!isCompletePlacementMapping(mapping)) return [];
  const designatorColumn = mapping.designator!;
  const layerColumn = mapping.layer!;
  return matrix.slice(headerIndex + 1).flatMap((row, index): PlacementRecord[] => {
    const designator = normalizeLocation(row[designatorColumn]);
    const rawLayer = text(row[layerColumn]);
    if (!designator && !rawLayer) return [];
    return [{ designator, side: normalizePlacementSide(rawLayer), rawLayer, sourceRow: headerIndex + index + 2 }];
  });
}

export function calculateMva(items: BomItem[], records: PlacementRecord[]): MvaSummary {
  const summary: MvaSummary = { smtTop: 0, smtBottom: 0, dipTop: 0, dipBottom: 0, included: [], excluded: [] };
  const grouped = new Map<string, PlacementRecord[]>();
  records.forEach((record) => {
    const key = normalizeLocation(record.designator);
    if (!key) return;
    grouped.set(key, [...(grouped.get(key) ?? []), record]);
  });

  grouped.forEach((group, designator) => {
    const sides = [...new Set(group.map((record) => record.side).filter(Boolean))] as PlacementSide[];
    if (group.some((record) => !record.side) || sides.length !== 1) {
      summary.excluded.push({ designator, rawLayer: group.map((record) => record.rawLayer).join(" / "), sourceRows: group.map((record) => record.sourceRow), reason: sides.length > 1 ? "同一插件位置出現衝突的 Top／Bottom" : "Layer 無法辨識" });
      return;
    }
    const candidates = items.filter((item) => item.positions.some((position) => normalizeLocation(position) === designator) && bomProcessKind(item));
    if (candidates.length !== 1) {
      summary.excluded.push({ designator, rawLayer: group[0].rawLayer, sourceRows: group.map((record) => record.sourceRow), reason: candidates.length ? "同時命中多個可判定製程的 BOM 料群" : "BOM 找不到可判定 SMT／DIP 的插件位置" });
      return;
    }
    const item = candidates[0];
    const process = bomProcessKind(item)!;
    const side = sides[0];
    const key = `${process.toLowerCase()}${side}` as "smtTop" | "smtBottom" | "dipTop" | "dipBottom";
    summary[key] += 1;
    summary.included.push({ designator, side, process, part: item.part, structure: item.structurePath?.join(" › ") ?? "" });
  });
  return summary;
}
