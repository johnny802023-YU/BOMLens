import { bomProcessKind, type BomAlternative, type BomCustomerMappingStatus, type BomCustomerPartAssociation, type BomItem } from "./bom-logic.ts";

export type CustomerBomColumnKey = "customerPartNumber" | "manufacturerParts" | "positions";
export type CustomerBomColumnMapping = Partial<Record<CustomerBomColumnKey, number>>;
export type CustomerMappingStatus = BomCustomerMappingStatus;

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
  companyPartNumbers: string[];
  companyRdCustomerPartNumbers: string[];
  customerAssociationEvidence: string[];
};

export type CustomerAlternativeMappingRow = {
  itemIndex: number;
  alternativeIndex: number;
  role: "主料" | "替料";
  part: string;
  manufacturerPart: string;
  positions: string[];
  rdCustomerPartNumbers: string[];
  customerPartNumbers: string[];
  customerAssociationEvidence: string[];
  status: Exclude<CustomerMappingStatus, "not-imported">;
  reason: string;
};

export type CustomerMappingResult = {
  items: BomItem[];
  rows: CustomerMappingRow[];
  alternativeRows: CustomerAlternativeMappingRow[];
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
  return text(value)
    .normalize("NFKC")
    .replace(/[\u200B-\u200D\u2060\uFEFF]/g, "")
    .replace(/\s+/g, "")
    .toUpperCase();
}

export function parseCustomerMpns(value: unknown) {
  return [...new Set(text(value)
    .split(/[|,，;；\r\n]+|(?:[~～]\s*){2,}/)
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
  return unique(item.alternatives.flatMap((alternative) => [
    normalizeMpn(alternative.manufacturerPart),
    ...(alternative.rdCustomerPartAssociations ?? []).map((association) => normalizeMpn(association.manufacturerPart)),
  ]));
}

function rdCustomerNumbers(item: BomItem | BomAlternative) {
  const values = "alternatives" in item
    ? item.rdCustomerPartNumbers ?? item.alternatives.flatMap((alternative) => alternative.rdCustomerPartNumbers ?? [])
    : item.rdCustomerPartNumbers ?? [];
  return [...new Set(values
    .map((partNumber) => partNumber.trim().toUpperCase())
    .filter(Boolean))];
}

type CustomerAssociationIndex = {
  enabled: boolean;
  byTpn: Map<string, BomCustomerPartAssociation[]>;
  invalidTpns: Set<string>;
  invalidReasons: Map<string, string[]>;
};

type MaintenanceResult = {
  status: "matched" | "rd-maintenance-missing" | "rd-maintenance-mismatch" | "tpn-association-mismatch" | "tpn-association-invalid";
  evidence: string[];
  affectedTpns: string[];
};

function buildCustomerAssociationIndex(items: BomItem[]): CustomerAssociationIndex {
  const byTpn = new Map<string, BomCustomerPartAssociation[]>();
  const invalidTpns = new Set<string>();
  const invalidReasons = new Map<string, string[]>();
  let enabled = false;
  items.forEach((item) => item.alternatives.forEach((alternative) => {
    if (alternative.rdCustomerAssociationStatus && alternative.rdCustomerAssociationStatus !== "not-provided") enabled = true;
    (alternative.rdCustomerPartAssociations ?? []).forEach((association) => {
      const tpn = association.customerPartNumber.trim().toUpperCase();
      if (!tpn) return;
      byTpn.set(tpn, [...(byTpn.get(tpn) ?? []), association]);
    });
    (alternative.rdCustomerAssociationUnresolvedTpns ?? []).forEach((tpn) => {
      const normalized = tpn.trim().toUpperCase();
      if (!normalized) return;
      invalidTpns.add(normalized);
      const reason = alternative.rdCustomerAssociationIssue || `${normalized} 的 R／S 欄資料不完整。`;
      invalidReasons.set(normalized, unique([...(invalidReasons.get(normalized) ?? []), reason]));
    });
  }));
  return { enabled, byTpn, invalidTpns, invalidReasons };
}

function associationEvidence(association: BomCustomerPartAssociation) {
  const maker = association.manufacturerName ? `${association.manufacturerName}/` : "";
  return `${association.customerPartNumber} → ${maker}${association.manufacturerPart}（Excel 第 ${association.sourceRow} 列）`;
}

function hasAssociationData(alternative: BomAlternative) {
  return Boolean(alternative.rdCustomerAssociationStatus && alternative.rdCustomerAssociationStatus !== "not-provided");
}

function matchingAssociations(alternative: BomAlternative, record: CustomerBomRecord, associationIndex?: CustomerAssociationIndex) {
  const customerTpn = record.customerPartNumber.trim().toUpperCase();
  const customerMpns = new Set(record.manufacturerParts.map(normalizeMpn).filter(Boolean));
  if (!customerTpn || !customerMpns.size) return [];
  const ownMatches = (alternative.rdCustomerPartAssociations ?? []).filter((association) =>
    association.customerPartNumber.trim().toUpperCase() === customerTpn
    && customerMpns.has(normalizeMpn(association.manufacturerPart)));
  if (ownMatches.length || !associationIndex) return ownMatches;
  const unresolved = new Set((alternative.rdCustomerAssociationUnresolvedTpns ?? []).map((tpn) => tpn.trim().toUpperCase()));
  if (!unresolved.has(customerTpn)) return [];
  return (associationIndex.byTpn.get(customerTpn) ?? []).filter((association) =>
    customerMpns.has(normalizeMpn(association.manufacturerPart)));
}

function matchedRecordMpns(alternative: BomAlternative, record: CustomerBomRecord, associationIndex: CustomerAssociationIndex) {
  const customerMpns = new Set(record.manufacturerParts.map(normalizeMpn).filter(Boolean));
  const associations = matchingAssociations(alternative, record, associationIndex);
  if (associations.length) {
    return unique(associations.map((association) => normalizeMpn(association.manufacturerPart)));
  }
  const manufacturerPart = normalizeMpn(alternative.manufacturerPart);
  return !hasAssociationData(alternative) && manufacturerPart && customerMpns.has(manufacturerPart) ? [manufacturerPart] : [];
}

function recordMaintenanceStatus(alternative: BomAlternative, record: CustomerBomRecord, associationIndex: CustomerAssociationIndex): MaintenanceResult {
  const associations = matchingAssociations(alternative, record, associationIndex);
  if (associations.length) {
    return { status: "matched", evidence: unique(associations.map(associationEvidence)), affectedTpns: [] };
  }
  return maintainedStatus(alternative, [record.customerPartNumber], associationIndex);
}

function aggregateMaintenance(results: MaintenanceResult[]) {
  const statuses = results.map((result) => result.status);
  const status = statuses.includes("rd-maintenance-missing")
    ? "rd-maintenance-missing"
    : statuses.includes("rd-maintenance-mismatch")
      ? "rd-maintenance-mismatch"
      : statuses.includes("tpn-association-invalid")
        ? "tpn-association-invalid"
        : statuses.includes("tpn-association-mismatch")
          ? "tpn-association-mismatch"
          : "matched";
  return {
    status,
    evidence: unique(results.flatMap((result) => result.evidence)),
    affectedTpns: unique(results.flatMap((result) => result.affectedTpns)),
  } satisfies MaintenanceResult;
}

function maintainedStatus(item: BomItem | BomAlternative, expectedTpns: string[], associationIndex: CustomerAssociationIndex): MaintenanceResult {
  const maintained = new Set(rdCustomerNumbers(item));
  const expected = [...new Set(expectedTpns.map((value) => value.trim().toUpperCase()).filter(Boolean))];
  if (!maintained.size) return { status: "rd-maintenance-missing", evidence: [], affectedTpns: expected };
  const missingTpns = expected.filter((tpn) => !maintained.has(tpn));
  if (missingTpns.length) return { status: "rd-maintenance-mismatch", evidence: [], affectedTpns: missingTpns };
  if (!associationIndex.enabled) return { status: "matched", evidence: [], affectedTpns: [] };

  const manufacturerPart = normalizeMpn("manufacturerPart" in item ? item.manufacturerPart : "");
  const evidence: string[] = [];
  const invalidTpns: string[] = [];
  const mismatchedTpns: string[] = [];
  expected.forEach((tpn) => {
    const candidates = associationIndex.byTpn.get(tpn) ?? [];
    const matches = candidates.filter((association) => normalizeMpn(association.manufacturerPart) === manufacturerPart);
    if (matches.length) {
      evidence.push(...matches.map(associationEvidence));
    } else {
      evidence.push(...candidates.map(associationEvidence));
      if (associationIndex.invalidTpns.has(tpn)) {
        invalidTpns.push(tpn);
        evidence.push(...(associationIndex.invalidReasons.get(tpn) ?? []));
      } else {
        mismatchedTpns.push(tpn);
      }
    }
  });
  if (invalidTpns.length) {
    return { status: "tpn-association-invalid", evidence: unique(evidence), affectedTpns: invalidTpns };
  }
  if (mismatchedTpns.length) {
    return { status: "tpn-association-mismatch", evidence: unique(evidence), affectedTpns: mismatchedTpns };
  }
  return { status: "matched", evidence: unique(evidence), affectedTpns: [] };
}

function unique(values: string[]) {
  return [...new Set(values.filter(Boolean))];
}

function emptyCounts(): CustomerMappingResult["counts"] {
  return {
    matched: 0,
    "rd-maintenance-missing": 0,
    "rd-maintenance-mismatch": 0,
    "tpn-association-mismatch": 0,
    "tpn-association-invalid": 0,
    "location-unmatched": 0,
    "mpn-unmatched": 0,
    "missing-mpn": 0,
    "tpn-missing": 0,
    ambiguous: 0,
  };
}

export function mapCustomerBom(items: BomItem[], records: CustomerBomRecord[]): CustomerMappingResult {
  const associationIndex = buildCustomerAssociationIndex(items);
  const alternatives = items.flatMap((item, itemIndex) => item.alternatives.map((alternative, alternativeIndex) => ({
    item,
    itemIndex,
    alternative,
    alternativeIndex,
    mpn: normalizeMpn(alternative.manufacturerPart),
    locationKey: positionKey(item.positions),
  })));
  const matchedRecordsByAlternative = new Map<string, CustomerBomRecord[]>();
  const alternativeKey = (itemIndex: number, alternativeIndex: number) => `${itemIndex}:${alternativeIndex}`;

  const rows = records.map((record): CustomerMappingRow => {
    const customerMpns = new Set(record.manufacturerParts.map(normalizeMpn).filter(Boolean));
    const locationKey = positionKey(record.positions);
    const locationItems = items.filter((item) => locationKey && positionKey(item.positions) === locationKey);
    const matches = alternatives.filter((candidate) =>
      locationKey
      && candidate.locationKey === locationKey
      && matchedRecordMpns(candidate.alternative, record, associationIndex).length > 0);
    const companyManufacturerParts = unique(locationItems.flatMap(companyMpns));
    const companyPartNumbers = unique(matches.map(({ alternative }) => alternative.part));
    const companyRdCustomerPartNumbers = unique(matches.flatMap(({ alternative }) => rdCustomerNumbers(alternative)));
    const companyItem = matches[0]?.item ?? (locationItems.length === 1 ? locationItems[0] : undefined);

    if (!locationItems.length) {
      return { record, status: "location-unmatched", reason: "找不到 Location 集合完全相同的公司料群。", companyManufacturerParts: [], companyPartNumbers: [], companyRdCustomerPartNumbers: [], customerAssociationEvidence: [] };
    }
    if (!record.manufacturerParts.length || locationItems.some((item) => companyMpns(item).length === 0)) {
      return {
        record,
        status: "missing-mpn",
        reason: "客戶或公司 MPN 資料不完整。",
        companyItem,
        companyManufacturerParts,
        companyPartNumbers,
        companyRdCustomerPartNumbers,
        customerAssociationEvidence: [],
      };
    }
    if (!matches.length) {
      return {
        record,
        status: "mpn-unmatched",
        reason: "Location 相同，但沒有任何公司主料／替料 MPN 完全匹配。",
        companyItem,
        companyManufacturerParts,
        companyPartNumbers,
        companyRdCustomerPartNumbers,
        customerAssociationEvidence: [],
      };
    }

    matches.forEach((match) => {
      const key = alternativeKey(match.itemIndex, match.alternativeIndex);
      matchedRecordsByAlternative.set(key, [...(matchedRecordsByAlternative.get(key) ?? []), record]);
    });
    const matchedCustomerMpns = new Set(matches.flatMap((match) => matchedRecordMpns(match.alternative, record, associationIndex)));
    const unmatchedCustomerMpns = [...customerMpns].filter((mpn) => !matchedCustomerMpns.has(mpn));
    if (unmatchedCustomerMpns.length) {
      return {
        record,
        status: "mpn-unmatched",
        reason: `Location 已匹配，但以下客戶 MPN 找不到公司主替料：${unmatchedCustomerMpns.join("、")}。已匹配的料件仍會帶入 TPN。`,
        companyItem,
        companyManufacturerParts,
        companyPartNumbers,
        companyRdCustomerPartNumbers,
        customerAssociationEvidence: [],
      };
    }
    const customerTpn = record.customerPartNumber.trim().toUpperCase();
    if (!customerTpn) {
      return {
        record,
        status: "tpn-missing",
        reason: "客戶 BOM TPN 空白，無法帶入公司主料／替料。",
        companyItem,
        companyManufacturerParts,
        companyPartNumbers,
        companyRdCustomerPartNumbers,
        customerAssociationEvidence: [],
      };
    }
    const maintenance = aggregateMaintenance(matches.map(({ alternative }) => recordMaintenanceStatus(alternative, record, associationIndex)));
    const customerAssociationEvidence = maintenance.evidence;
    const maintenanceStatus = maintenance.status;
    return {
      record,
      status: maintenanceStatus,
      reason: maintenanceStatus === "matched"
        ? associationIndex.enabled
          ? "Location 與每顆主料／替料 MPN 完全匹配，且客戶 TPN 與 BOM S欄 MPN 關聯一致。"
          : "Location 與每顆主料／替料 MPN 完全匹配，客戶 TPN 也存在於各料件 BOM R欄；此 BOM 未提供 S欄關聯資料。"
        : maintenanceStatus === "rd-maintenance-missing"
          ? "Location 與 MPN 已匹配，但至少一顆主料／替料的 BOM R欄沒有維護此 TPN，請 RD 維護。"
          : maintenanceStatus === "rd-maintenance-mismatch"
            ? "Location 與 MPN 已匹配，但至少一顆主料／替料的 BOM R欄缺少此客戶 TPN，請 RD 確認並維護。"
            : maintenanceStatus === "tpn-association-invalid"
              ? "BOM R欄包含此 TPN，但相關 R／S 欄筆數不同或含空白，無法確認 TPN 對應 MPN，請人工確認。"
              : "BOM R欄包含此 TPN，但所有 S欄候選 MPN 均與客戶 BOM MPN 不同。",
      companyItem,
      companyManufacturerParts,
      companyPartNumbers,
      companyRdCustomerPartNumbers,
      customerAssociationEvidence,
    };
  });

  const alternativeRows: CustomerAlternativeMappingRow[] = alternatives.map((candidate) => {
    const matchingRecords = matchedRecordsByAlternative.get(alternativeKey(candidate.itemIndex, candidate.alternativeIndex)) ?? [];
    const customerPartNumbers = unique(matchingRecords.map((record) => record.customerPartNumber.trim().toUpperCase()));
    const rdPartNumbers = rdCustomerNumbers(candidate.alternative);
    const associationMpns = unique((candidate.alternative.rdCustomerPartAssociations ?? []).map((association) => normalizeMpn(association.manufacturerPart)));
    const identifiableMpns = hasAssociationData(candidate.alternative) ? associationMpns : [candidate.mpn].filter(Boolean);
    const relatedRecords = records.filter((record) => positionKey(record.positions) === candidate.locationKey);
    const relatedAssociationEvidence = unique([
      ...relatedRecords
      .flatMap((record) => {
        const tpn = record.customerPartNumber.trim().toUpperCase();
        return (candidate.alternative.rdCustomerPartAssociations ?? [])
          .filter((association) => association.customerPartNumber.trim().toUpperCase() === tpn)
          .map(associationEvidence);
      }),
      ...(relatedRecords.some((record) => (candidate.alternative.rdCustomerAssociationUnresolvedTpns ?? []).includes(record.customerPartNumber.trim().toUpperCase()))
        ? [candidate.alternative.rdCustomerAssociationIssue ?? ""]
        : []),
    ]);
    let status: CustomerAlternativeMappingRow["status"];
    let reason: string;

    if (matchingRecords.length && !customerPartNumbers.length) {
      status = "tpn-missing";
      reason = "Location 與 MPN 已匹配，但客戶 BOM TPN 空白。";
    } else if (customerPartNumbers.length) {
      const maintenance = aggregateMaintenance(matchingRecords.map((record) => recordMaintenanceStatus(candidate.alternative, record, associationIndex)));
      status = maintenance.status;
      reason = status === "matched"
        ? associationIndex.enabled
          ? `已依 Location＋MPN 帶入 ${customerPartNumbers.length} 組 TPN，且 TPN／S欄 MPN 關聯完全一致。`
          : `已依 Location＋MPN 帶入 ${customerPartNumbers.length} 組 TPN，且均存在於此料件 BOM R欄；此 BOM 未提供 S欄。`
        : status === "rd-maintenance-missing"
          ? `已帶入 ${customerPartNumbers.length} 組 TPN，但此料件 BOM R欄空白，請 RD 維護。`
          : status === "rd-maintenance-mismatch"
            ? `已帶入 ${customerPartNumbers.length} 組 TPN，但 BOM R欄缺少：${customerPartNumbers.filter((tpn) => !rdPartNumbers.includes(tpn)).join("、")}。`
            : status === "tpn-association-invalid"
              ? `TPN ${maintenance.affectedTpns.join("、")} 的 R／S 欄順序資料不完整，請人工確認。`
              : `TPN ${maintenance.affectedTpns.join("、")} 在 S欄沒有與公司／客戶 MPN ${candidate.mpn} 完全相同的候選。`;
    } else if (!identifiableMpns.length) {
      const unresolvedTpns = new Set((candidate.alternative.rdCustomerAssociationUnresolvedTpns ?? []).map((value) => value.trim().toUpperCase()));
      const hasRelatedTpn = relatedRecords.some((record) => unresolvedTpns.has(record.customerPartNumber.trim().toUpperCase()));
      if (candidate.alternative.rdCustomerAssociationStatus === "invalid" && hasRelatedTpn) {
        status = "tpn-association-invalid";
        reason = "Location 與 TPN 相同，但 R／S 欄順序資料不完整，無法確認對應。";
      } else {
        status = "missing-mpn";
        reason = "公司 BOM 的 P 欄及 R／S 對應均沒有可用的製造廠商料號。";
      }
    } else {
      const sameMpn = records.filter((record) => record.manufacturerParts.map(normalizeMpn).some((mpn) => identifiableMpns.includes(mpn)));
      const sameLocation = records.filter((record) => positionKey(record.positions) === candidate.locationKey);
      const sameTpnAtLocation = sameLocation.filter((record) => {
        const tpn = record.customerPartNumber.trim().toUpperCase();
        return (candidate.alternative.rdCustomerPartAssociations ?? []).some((association) => association.customerPartNumber.trim().toUpperCase() === tpn)
          || (candidate.alternative.rdCustomerAssociationUnresolvedTpns ?? []).some((value) => value.trim().toUpperCase() === tpn);
      });
      if (sameTpnAtLocation.length && hasAssociationData(candidate.alternative)) {
        status = candidate.alternative.rdCustomerAssociationStatus === "invalid" ? "tpn-association-invalid" : "tpn-association-mismatch";
        reason = status === "tpn-association-invalid"
          ? "Location 與 TPN 相同，但 R／S 欄順序資料不完整，無法確認對應。"
          : "Location 與 TPN 相同，但 S欄沒有與客戶 BOM 完全相同的 MPN。";
      } else if (sameMpn.length) {
        status = "location-unmatched";
        reason = "MPN 完全相同，但 Location 集合不同；未帶入其他位置的 TPN。";
      } else if (sameLocation.some((record) => !record.manufacturerParts.length)) {
        status = "missing-mpn";
        reason = "Location 相同，但客戶 BOM MPN 空白。";
      } else {
        status = "mpn-unmatched";
        reason = "Location 相同，但客戶 BOM 找不到完全相同的 MPN。";
      }
    }
    return {
      itemIndex: candidate.itemIndex,
      alternativeIndex: candidate.alternativeIndex,
      role: candidate.alternativeIndex === 0 ? "主料" : "替料",
      part: candidate.alternative.part,
      manufacturerPart: candidate.alternative.manufacturerPart,
      positions: candidate.item.positions,
      rdCustomerPartNumbers: rdPartNumbers,
      customerPartNumbers,
      customerAssociationEvidence: customerPartNumbers.length
        ? aggregateMaintenance(matchingRecords.map((record) => recordMaintenanceStatus(candidate.alternative, record, associationIndex))).evidence
        : relatedAssociationEvidence,
      status,
      reason,
    };
  });

  const statusPriority: Array<CustomerAlternativeMappingRow["status"]> = ["ambiguous", "tpn-missing", "missing-mpn", "mpn-unmatched", "location-unmatched", "tpn-association-invalid", "tpn-association-mismatch", "rd-maintenance-mismatch", "rd-maintenance-missing", "matched"];
  const mappedItems = items.map((item, itemIndex) => {
    const itemRows = alternativeRows.filter((row) => row.itemIndex === itemIndex);
    const mappedAlternatives = item.alternatives.map((alternative, alternativeIndex) => {
      const row = itemRows.find((candidate) => candidate.alternativeIndex === alternativeIndex)!;
      return {
        ...alternative,
        customerPartNumbers: row.customerPartNumbers,
        customerMappingStatus: row.status,
        customerMappingReason: row.reason,
        customerAssociationEvidence: row.customerAssociationEvidence,
      };
    });
    const customerPartNumbers = unique(itemRows.flatMap((row) => row.customerPartNumbers));
    const customerMappingStatus = statusPriority.find((status) => itemRows.some((row) => row.status === status)) ?? "mpn-unmatched";
    const customerMappingReason = itemRows.filter((row) => row.status !== "matched").map((row) => `${row.role} ${row.part || row.manufacturerPart}：${row.reason}`).join("；")
      || `主料與 ${Math.max(0, itemRows.length - 1)} 顆替料均已完成 TPN 對應。`;
    return {
      ...item,
      alternatives: mappedAlternatives,
      customerPartNumber: customerPartNumbers[0],
      customerPartNumbers,
      customerMappingStatus,
      customerMappingReason,
    };
  });

  const counts = emptyCounts();
  alternativeRows.forEach((row) => { counts[row.status] += 1; });
  return { items: mappedItems, rows, alternativeRows, counts };
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
