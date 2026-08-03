export type BomAlternative = {
  part: string;
  manufacturerPart: string;
  manufacturerName?: string;
  description: string;
  spec: string;
  rdCustomerPartNumbers?: string[];
  rdCustomerPartAssociations?: BomCustomerPartAssociation[];
  rdCustomerAssociationStatus?: BomCustomerAssociationStatus;
  rdCustomerAssociationIssue?: string;
  rdCustomerAssociationUnresolvedTpns?: string[];
  customerPartNumbers?: string[];
  customerMappingStatus?: BomCustomerMappingStatus;
  customerMappingReason?: string;
  customerAssociationEvidence?: string[];
};

export type BomCustomerPartAssociation = {
  customerPartNumber: string;
  manufacturerPart: string;
  manufacturerName?: string;
  sourceRow: number;
};

export type BomCustomerAssociationStatus = "not-provided" | "valid" | "invalid";

export type BomCustomerMappingStatus = "matched" | "rd-maintenance-missing" | "rd-maintenance-mismatch" | "tpn-association-mismatch" | "tpn-association-invalid" | "location-unmatched" | "mpn-unmatched" | "missing-mpn" | "tpn-missing" | "ambiguous" | "not-imported";

export type BomStructureKind = "root69" | "vb-t" | "vb-d" | "board60" | "pcb" | "flat";
export type BomProcessKind = "SMT" | "DIP";
export type BomProcessChange = {
  before: BomProcessKind;
  after: BomProcessKind;
};

export type BomItem = {
  ref: string;
  part: string;
  manufacturerPart: string;
  manufacturerName?: string;
  value: string;
  description: string;
  qty: number;
  positions: string[];
  alternatives: BomAlternative[];
  sourceRows?: number[];
  sourceSheet?: string;
  declaredQty?: number;
  structureKind?: BomStructureKind;
  structurePath?: string[];
  structureKey?: string;
  customerPartNumber?: string;
  customerPartNumbers?: string[];
  customerMappingStatus?: BomCustomerMappingStatus;
  customerMappingReason?: string;
  rdCustomerPartNumbers?: string[];
  rdCustomerPartAssociations?: BomCustomerPartAssociation[];
};

export type DiffKind = "added" | "removed" | "changed" | "same";
export type DiffCategory = Exclude<DiffKind, "same">;
export type DiffPrimaryType = "componentAdded" | "componentRemoved" | "substituteAdded" | "substituteRemoved" | "partReplaced" | "positionChanged" | "same";

export type BomDiff = {
  ref: string;
  kind: DiffKind;
  categories: DiffCategory[];
  primaryType: DiffPrimaryType;
  before?: BomItem;
  after?: BomItem;
  fields: string[];
  addedParts: BomAlternative[];
  removedParts: BomAlternative[];
  newParts: BomAlternative[];
  deletedParts: BomAlternative[];
  addedPositions: string[];
  removedPositions: string[];
  replacementPositions: string[];
  processChange?: BomProcessChange;
  matchConfidence: "high" | "medium" | "low";
  matchReason: string;
  needsReview: boolean;
};

export type CompanyColumnKey = "ref" | "part" | "qty" | "positions" | "description" | "spec" | "manufacturerName" | "manufacturerPart" | "customerPartNumbers" | "customerPartAssociations";
export type CompanyColumnMapping = Partial<Record<CompanyColumnKey, number>>;
export type ImportIssue = { severity: "error" | "warning"; code: string; message: string; rows?: number[] };
export type ImportAudit = {
  headerRow: number;
  mapping: CompanyColumnMapping;
  mappingLabels: Partial<Record<CompanyColumnKey, string>>;
  sourceRows: number;
  groupCount: number;
  positionCount: number;
  issues: ImportIssue[];
};

const aliases = {
  ref: ["ref", "reference", "references", "designator", "refdes", "reference designator", "位號", "位置", "項次"],
  part: ["part", "part number", "part no", "pn", "料號", "主件料號", "零件料號", "公司料號", "2.主件料號"],
  manufacturerPart: ["mpn", "manufacturer part number", "製造商料號", "製造廠商料號", "製造商型號", "16製造廠商料號"],
  manufacturerName: ["manufacturer", "manufacturer name", "mfr", "製造商", "製造商名稱", "製造廠商", "11製造廠商"],
  value: ["value", "component value", "規格", "數值", "值"],
  description: ["description", "desc", "comment", "item description", "說明", "描述", "品名"],
  qty: ["qty", "quantity", "count", "數量", "用量", "組成用量"],
  positions: ["placement", "placements", "location", "locations", "插件位置", "位置"],
  customerPartNumbers: ["customer part number", "customer pn", "customer tpn", "tpn", "客戶料號", "客戶 tpn", "對應客戶料號"],
  customerPartAssociations: [
    "customer part association",
    "tpn mpn association",
    "廠商料號及型態",
    "廠商/料號及型態",
    "廠商／料號及型態",
    "廠商料號及型號",
    "廠商/料號及型號",
    "廠商／料號及型號",
  ],
};

type CompanyColumns = {
  ref: number;
  part: number;
  qty: number;
  positions: number;
  description?: number;
  spec?: number;
  manufacturerName?: number;
  manufacturerPart?: number;
  customerPartNumbers?: number;
  customerPartAssociations?: number;
};

export const companyColumnLabels: Record<CompanyColumnKey, string> = {
  ref: "項次",
  part: "主件料號",
  qty: "數量／組成用量",
  positions: "插件位置",
  description: "品名／描述",
  spec: "規格",
  manufacturerName: "製造廠商",
  manufacturerPart: "製造廠商料號",
  customerPartNumbers: "TPN／對應客戶料號（R欄）",
  customerPartAssociations: "TPN 對應廠商料號（S欄）",
};

export const requiredCompanyColumns: CompanyColumnKey[] = ["ref", "part", "qty", "positions"];

const companyHeaderAliases = {
  ref: ["項次", "序號", "item", "item no", "item number"],
  part: ["主件料號", "2.主件料號", "料號", "公司料號", "零件料號", "part", "part no", "part number", "pn"],
  qty: ["組成用量", "數量", "用量", "qty", "quantity", "count"],
  positions: ["插件位置", "位號", "reference designator", "refdes", "placement", "placements", "location", "locations"],
  description: ["品名", "描述", "說明", "description", "desc", "item description"],
  spec: ["規格", "數值", "值", "spec", "specification", "value"],
  manufacturerName: ["製造廠商", "11製造廠商", "製造商名稱", "製造商", "manufacturer name", "manufacturer", "mfr"],
  manufacturerPart: ["製造廠商料號", "16製造廠商料號", "製造商料號", "製造商型號", "manufacturer part number", "mpn"],
  customerPartNumbers: ["對應客戶料號", "客戶料號", "客戶 tpn", "customer part number", "customer pn", "customer tpn", "tpn"],
  customerPartAssociations: [
    "廠商/料號及型態",
    "廠商／料號及型態",
    "廠商料號及型態",
    "廠商/料號及型號",
    "廠商／料號及型號",
    "廠商料號及型號",
    "tpn 對應廠商料號",
    "customer part association",
    "tpn mpn association",
  ],
};

function text(value: unknown) {
  return value == null ? "" : String(value).trim();
}

function normalizeKey(value: string) {
  return value.trim().toLowerCase().replace(/[：:()（）]/g, "").replace(/[_-]+/g, " ").replace(/\s+/g, " ");
}

function normalizeValue(value: string) {
  return value.trim().toUpperCase().replace(/\s+/g, " ");
}

export function canonicalPartNumber(value: unknown) {
  const compact = text(value).replace(/\s+/g, "");
  return compact.length > 12 ? compact.slice(-12) : compact;
}

export function bomStructureLabel(item?: BomItem) {
  return item?.structurePath?.filter(Boolean).join(" › ") ?? "";
}

export function bomProcessKind(item?: BomItem): BomProcessKind | undefined {
  if (item?.structureKind === "vb-t" || item?.structureKind === "board60") return "SMT";
  if (item?.structureKind === "vb-d") return "DIP";
  return undefined;
}

export function bomDiffDisplayFields(diff: BomDiff) {
  const fields = diff.fields.map((field) => field === "客戶料號差異" ? "TPN 差異" : field);
  if (diff.processChange) {
    const index = fields.indexOf("製程別放置異常");
    const label = `製程別放置異常（${diff.processChange.before} → ${diff.processChange.after}）`;
    if (index >= 0) fields[index] = label;
    else fields.push(label);
  }
  return fields;
}

function structureMarker(rawPart: string): Exclude<BomStructureKind, "flat"> | null {
  const compact = structurePartLabel(rawPart);
  if (/^69-?[A-Z0-9]+$/.test(compact)) return "root69";
  if (/^VB-?[A-Z0-9]+T$/.test(compact)) return "vb-t";
  if (/^VB-?[A-Z0-9]+D$/.test(compact)) return "vb-d";
  if (/^60-?[A-Z0-9]+$/.test(compact)) return "board60";
  if (canonicalPartNumber(compact).toUpperCase().startsWith("08")) return "pcb";
  return null;
}

function structurePartLabel(rawPart: string) {
  return rawPart.toUpperCase().replace(/\s+/g, "").replace(/^[|│├└─-]+/, "");
}

function pick(row: Record<string, unknown>, names: string[]) {
  const entry = Object.entries(row).find(([key]) => names.includes(normalizeKey(key)));
  return text(entry?.[1]);
}

export function parseQuantity(value: unknown) {
  const first = text(value).split("/")[0].replace(/,/g, "");
  const quantity = Number.parseFloat(first);
  return Number.isFinite(quantity) ? quantity : 0;
}

export function parsePositions(value: unknown) {
  return [...new Set(text(value)
    .split(/[|,，;；\s]+/)
    .map((position) => position.trim().toUpperCase())
    .filter(Boolean))];
}

export function parseCustomerPartNumbers(value: unknown) {
  return [...new Set(text(value)
    .split(/[|,，;；\r\n]+/)
    .map((partNumber) => partNumber.trim().toUpperCase())
    .filter(Boolean))];
}

function parseOrderedValues(value: unknown) {
  const raw = text(value);
  if (!raw) return [];
  return raw.split(/[|,，;；\r\n]+/).map((entry) => entry.trim().toUpperCase());
}

function parseCustomerAssociationValue(value: string) {
  const slashIndex = value.lastIndexOf("/");
  return {
    manufacturerName: slashIndex >= 0 ? value.slice(0, slashIndex).trim() : "",
    manufacturerPart: (slashIndex >= 0 ? value.slice(slashIndex + 1) : value).trim(),
  };
}

function parseCustomerPartAssociations(customerPartValue: unknown, associationValue: unknown, sourceRow: number, provided: boolean) {
  const customerPartNumbers = parseOrderedValues(customerPartValue);
  if (!provided) {
    return {
      associations: [] as BomCustomerPartAssociation[],
      status: "not-provided" as const,
      issue: "",
      unresolvedTpns: [] as string[],
    };
  }
  const associationValues = parseOrderedValues(associationValue);
  const invalid = customerPartNumbers.length !== associationValues.length
    || customerPartNumbers.some((value) => !value)
    || associationValues.some((value) => !value);
  if (invalid) {
    return {
      associations: [] as BomCustomerPartAssociation[],
      status: "invalid" as const,
      issue: `Excel 第 ${sourceRow} 列 R／S 欄筆數不同或包含空白項目。`,
      unresolvedTpns: customerPartNumbers.filter(Boolean),
    };
  }
  return {
    associations: customerPartNumbers.map((customerPartNumber, index) => {
      const parsed = parseCustomerAssociationValue(associationValues[index]);
      return {
        customerPartNumber,
        manufacturerPart: parsed.manufacturerPart,
        manufacturerName: parsed.manufacturerName,
        sourceRow,
      };
    }),
    status: "valid" as const,
    issue: "",
    unresolvedTpns: [] as string[],
  };
}

function findHeaderColumn(row: unknown[], names: string[]) {
  const normalizedNames = new Set(names.map(normalizeKey));
  return row.findIndex((cell) => normalizedNames.has(normalizeKey(text(cell))));
}

export function detectCompanyColumns(headerRow: unknown[]): CompanyColumnMapping {
  const ref = findHeaderColumn(headerRow, companyHeaderAliases.ref);
  const part = findHeaderColumn(headerRow, companyHeaderAliases.part);
  const qty = findHeaderColumn(headerRow, companyHeaderAliases.qty);
  const positions = findHeaderColumn(headerRow, companyHeaderAliases.positions);
  const optionalColumn = (names: string[]) => {
    const index = findHeaderColumn(headerRow, names);
    return index >= 0 ? index : undefined;
  };
  return {
    ...(ref >= 0 ? { ref } : {}),
    ...(part >= 0 ? { part } : {}),
    ...(qty >= 0 ? { qty } : {}),
    ...(positions >= 0 ? { positions } : {}),
    description: optionalColumn(companyHeaderAliases.description),
    spec: optionalColumn(companyHeaderAliases.spec),
    manufacturerName: optionalColumn(companyHeaderAliases.manufacturerName),
    manufacturerPart: optionalColumn(companyHeaderAliases.manufacturerPart),
    customerPartNumbers: optionalColumn(companyHeaderAliases.customerPartNumbers),
    customerPartAssociations: optionalColumn(companyHeaderAliases.customerPartAssociations),
  };
}

export function isCompleteCompanyMapping(mapping: CompanyColumnMapping): mapping is CompanyColumns {
  return requiredCompanyColumns.every((key) => mapping[key] != null && mapping[key]! >= 0);
}

function getCompanyColumns(headerRow: unknown[]): CompanyColumns | null {
  const mapping = detectCompanyColumns(headerRow);
  return isCompleteCompanyMapping(mapping) ? mapping : null;
}

export function findCompanyHeader(matrix: unknown[][]) {
  return matrix.findIndex((row) => getCompanyColumns(row) !== null);
}

export function parseCompanyBomMatrix(matrix: unknown[][], headerIndex = findCompanyHeader(matrix), suppliedMapping?: CompanyColumnMapping): BomItem[] {
  if (headerIndex < 0) return [];
  const detected = suppliedMapping ?? detectCompanyColumns(matrix[headerIndex]);
  if (!isCompleteCompanyMapping(detected)) return [];
  const columns = detected;
  const cell = (row: unknown[], index?: number) => index == null ? "" : row[index];

  type GroupDraft = {
    ref: string;
    alternatives: BomAlternative[];
    positions: Set<string>;
    rawQuantity: number;
    sourceRows: number[];
    structureKind: BomStructureKind;
    structurePath: string[];
    structureKey: string;
  };

  const groups: GroupDraft[] = [];
  let currentGroup: GroupDraft | null = null;
  let rootLabel = "";
  let branchLabel = "";
  let branchKind: BomStructureKind = "flat";
  let rootSection = 0;
  let groupSequence = 0;

  for (const [offset, row] of matrix.slice(headerIndex + 1).entries()) {
    const sourceRow = headerIndex + offset + 2;
    const nextRef = text(cell(row, columns.ref));
    const rawPart = text(cell(row, columns.part));
    const part = canonicalPartNumber(rawPart);
    const description = text(cell(row, columns.description));
    const spec = text(cell(row, columns.spec));
    const rawQuantity = parseQuantity(cell(row, columns.qty));
    const positions = parsePositions(cell(row, columns.positions));
    const manufacturerName = text(cell(row, columns.manufacturerName));
    const manufacturerPart = text(cell(row, columns.manufacturerPart));
    const rdCustomerPartNumbers = parseCustomerPartNumbers(cell(row, columns.customerPartNumbers));
    const parsedAssociations = parseCustomerPartAssociations(
      cell(row, columns.customerPartNumbers),
      cell(row, columns.customerPartAssociations),
      sourceRow,
      columns.customerPartAssociations != null,
    );

    if (!part && !manufacturerPart && !positions.length && !rawQuantity) continue;

    const detectedMarker = structureMarker(rawPart);
    const marker = detectedMarker === "root69" || nextRef ? detectedMarker : null;
    if (nextRef || marker === "root69") {
      const markerLabel = structurePartLabel(rawPart);
      if (marker === "root69") {
        rootSection += 1;
        rootLabel = markerLabel;
        branchLabel = "";
        branchKind = "root69";
      } else if (marker === "vb-t" || marker === "vb-d" || marker === "board60" || marker === "pcb") {
        branchLabel = markerLabel;
        branchKind = marker;
      }

      const structureKind = marker ?? branchKind;
      const structurePath = marker === "root69"
        ? [rootLabel]
        : branchLabel
          ? rootLabel ? [rootLabel, branchLabel] : [branchLabel]
          : rootLabel ? [rootLabel] : [];
      groupSequence += 1;
      currentGroup = {
        ref: nextRef || `69-架構-${rootSection}`,
        alternatives: [],
        positions: new Set<string>(),
        rawQuantity: 0,
        sourceRows: [],
        structureKind,
        structurePath,
        structureKey: `${rootSection || 0}:${structureKind}:${groupSequence}`,
      };
      groups.push(currentGroup);
    }
    if (!currentGroup) continue;

    if (part || manufacturerPart) {
      const alternative = {
        part,
        manufacturerPart,
        manufacturerName,
        description,
        spec,
        rdCustomerPartNumbers,
        rdCustomerPartAssociations: parsedAssociations.associations,
        rdCustomerAssociationStatus: parsedAssociations.status,
        rdCustomerAssociationIssue: parsedAssociations.issue,
        rdCustomerAssociationUnresolvedTpns: parsedAssociations.unresolvedTpns,
      };
      const alternativeKey = altKey(alternative);
      const existingAlternative = currentGroup.alternatives.find((candidate) => altKey(candidate) === alternativeKey);
      if (!existingAlternative) {
        currentGroup.alternatives.push(alternative);
      } else {
        existingAlternative.rdCustomerPartNumbers = [...new Set([...(existingAlternative.rdCustomerPartNumbers ?? []), ...rdCustomerPartNumbers])];
        existingAlternative.rdCustomerPartAssociations = [
          ...(existingAlternative.rdCustomerPartAssociations ?? []),
          ...parsedAssociations.associations,
        ];
        existingAlternative.rdCustomerAssociationUnresolvedTpns = [...new Set([
          ...(existingAlternative.rdCustomerAssociationUnresolvedTpns ?? []),
          ...parsedAssociations.unresolvedTpns,
        ])];
        if (existingAlternative.rdCustomerAssociationStatus !== "invalid" && parsedAssociations.status === "invalid") {
          existingAlternative.rdCustomerAssociationStatus = "invalid";
        } else if (existingAlternative.rdCustomerAssociationStatus === "not-provided" && parsedAssociations.status === "valid") {
          existingAlternative.rdCustomerAssociationStatus = "valid";
        }
        existingAlternative.rdCustomerAssociationIssue = [
          existingAlternative.rdCustomerAssociationIssue,
          parsedAssociations.issue,
        ].filter(Boolean).join("；");
      }
    }
    positions.forEach((position) => currentGroup.positions.add(position));
    currentGroup.rawQuantity = Math.max(currentGroup.rawQuantity, rawQuantity);
    currentGroup.sourceRows.push(sourceRow);
  }

  return groups.map((group) => {
    const positions = [...group.positions].sort(naturalSort);
    const first = group.alternatives[0] ?? { part: "", manufacturerPart: "", manufacturerName: "", description: "", spec: "", rdCustomerPartNumbers: [] };
    const rdCustomerPartNumbers = [...new Set(group.alternatives.flatMap((alternative) => alternative.rdCustomerPartNumbers ?? []))];
    const rdCustomerPartAssociations = group.alternatives.flatMap((alternative) => alternative.rdCustomerPartAssociations ?? []);
    return {
      ref: group.ref,
      part: first.part,
      manufacturerPart: first.manufacturerPart,
      manufacturerName: first.manufacturerName,
      value: first.spec,
      description: first.description,
      qty: positions.length || group.rawQuantity,
      positions,
      alternatives: group.alternatives,
      sourceRows: group.sourceRows,
      declaredQty: group.rawQuantity,
      structureKind: group.structureKind,
      structurePath: group.structurePath,
      structureKey: group.structureKey,
      rdCustomerPartNumbers,
      rdCustomerPartAssociations,
    };
  });
}

export function analyzeCompanyBomMatrix(matrix: unknown[][], headerIndex: number, suppliedMapping?: CompanyColumnMapping): { items: BomItem[]; audit: ImportAudit } {
  const header = matrix[headerIndex] ?? [];
  const mapping = suppliedMapping ?? detectCompanyColumns(header);
  const mappingLabels = Object.fromEntries(Object.entries(mapping).map(([key, index]) => [key, text(header[index as number]) || `第 ${(index as number) + 1} 欄`])) as ImportAudit["mappingLabels"];
  const issues: ImportIssue[] = [];
  const missing = requiredCompanyColumns.filter((key) => mapping[key] == null || mapping[key]! < 0);
  if (missing.length) {
    issues.push({ severity: "error", code: "missing-columns", message: `缺少必要欄位：${missing.map((key) => companyColumnLabels[key]).join("、")}` });
  }
  const items = isCompleteCompanyMapping(mapping) ? parseCompanyBomMatrix(matrix, headerIndex, mapping) : [];
  if (!missing.length && !items.length) issues.push({ severity: "error", code: "no-data", message: "標題下方沒有可辨識的 BOM 資料。" });
  const associationIssues = items.flatMap((item) => item.alternatives
    .filter((alternative) => alternative.rdCustomerAssociationStatus === "invalid")
    .map((alternative) => alternative.rdCustomerAssociationIssue)
    .filter(Boolean));
  if (associationIssues.length) {
    issues.push({
      severity: "warning",
      code: "customer-association-invalid",
      message: `發現 ${associationIssues.length} 筆 R／S 欄順序資料不完整；客戶 BOM 對應時將列入人工確認。`,
    });
  }

  const duplicatePositions = new Map<string, Array<{ ref: string; rows: number[] }>>();
  items.forEach((item) => item.positions.forEach((position) => {
    const scopedPosition = `${item.structureKey?.split(":").slice(0, 2).join(":") ?? "flat"}\u0000${position}`;
    const uses = duplicatePositions.get(scopedPosition) ?? [];
    uses.push({ ref: item.ref, rows: item.sourceRows ?? [] });
    duplicatePositions.set(scopedPosition, uses);
  }));
  duplicatePositions.forEach((uses, scopedPosition) => {
    const position = scopedPosition.split("\u0000").at(-1) ?? scopedPosition;
    if (uses.length > 1) issues.push({ severity: "warning", code: "duplicate-position", message: `${position} 同時出現在 ${uses.map((use) => use.ref).join("、")}，請確認是否重複分配。`, rows: uses.flatMap((use) => use.rows) });
  });

  if (isCompleteCompanyMapping(mapping)) {
    const rawParts = new Map<string, Set<string>>();
    const rowScopes = new Map<number, string>();
    items.forEach((item) => item.sourceRows?.forEach((row) => rowScopes.set(row, item.structureKey?.split(":").slice(0, 2).join(":") ?? "flat")));
    matrix.slice(headerIndex + 1).forEach((row, offset) => {
      const rowNumber = headerIndex + offset + 2;
      const rawPart = text(row[mapping.part]).replace(/\s+/g, "");
      const canonical = canonicalPartNumber(rawPart);
      if (canonical && rawPart) {
        const collisionKey = `${rowScopes.get(rowNumber) ?? "flat"}\u0000${normalizeValue(canonical)}`;
        const variants = rawParts.get(collisionKey) ?? new Set<string>();
        variants.add(normalizeValue(rawPart));
        rawParts.set(collisionKey, variants);
      }
    });
    rawParts.forEach((variants, collisionKey) => {
      const canonical = collisionKey.split("\u0000").at(-1) ?? collisionKey;
      if (variants.size > 1) issues.push({ severity: "warning", code: "part-collision", message: `多個原始料號取最右 12 碼後都變成 ${canonical}：${[...variants].join("、")}` });
    });
    items.forEach((item) => {
      if (!item.alternatives.some((part) => part.part)) issues.push({ severity: "warning", code: "missing-part", message: `項次 ${item.ref} 沒有料號。`, rows: item.sourceRows });
      if (item.positions.length && item.declaredQty && item.positions.length !== item.declaredQty) {
        const structure = bomStructureLabel(item);
        issues.push({ severity: "warning", code: "quantity-mismatch", message: `${structure ? `${structure}／` : ""}項次 ${item.ref} 的數量為 ${item.declaredQty}，但插件位置共有 ${item.positions.length} 個。`, rows: item.sourceRows });
      }
    });
  }

  return {
    items,
    audit: {
      headerRow: headerIndex + 1,
      mapping,
      mappingLabels,
      sourceRows: Math.max(0, matrix.length - headerIndex - 1),
      groupCount: items.length,
      positionCount: new Set(items.flatMap((item) => item.positions.map((position) => `${item.structureKey?.split(":").slice(0, 2).join(":") ?? "flat"}\u0000${position}`))).size,
      issues,
    },
  };
}

export function rowsToBom(rows: Record<string, unknown>[]): BomItem[] {
  return rows
    .map((row, index) => {
      const ref = pick(row, aliases.ref) || `ROW-${index + 1}`;
      const part = canonicalPartNumber(pick(row, aliases.part));
      const manufacturerPart = pick(row, aliases.manufacturerPart);
      const manufacturerName = pick(row, aliases.manufacturerName);
      const value = pick(row, aliases.value);
      const description = pick(row, aliases.description);
      const positions = parsePositions(pick(row, aliases.positions));
      const qty = positions.length || parseQuantity(pick(row, aliases.qty)) || 1;
      return {
        ref,
        part,
        manufacturerPart,
        manufacturerName,
        value,
        description,
        qty,
        positions,
        alternatives: part || manufacturerPart ? [{ part, manufacturerPart, manufacturerName, description, spec: value }] : [],
      };
    })
    .filter((row) => row.alternatives.length || row.positions.length || !row.ref.startsWith("ROW-"));
}

export function compareBom(before: BomItem[], after: BomItem[]): BomDiff[] {
  const beforePartKeys = new Set(before.flatMap((item) => item.alternatives.map(partKey)).filter(Boolean));
  const afterPartKeys = new Set(after.flatMap((item) => item.alternatives.map(partKey)).filter(Boolean));
  const candidates: Array<{ beforeIndex: number; afterIndex: number; score: number; processChange?: BomProcessChange }> = [];
  before.forEach((left, beforeIndex) => after.forEach((right, afterIndex) => {
    const sameStructure = structureCompatible(left, right);
    const processChange = sameStructure ? undefined : processPlacementChange(left, right);
    if (!sameStructure && !processChange) return;
    const baseScore = groupMatchScore(left, right);
    if (baseScore <= 0) return;
    const score = baseScore + (sameStructure && (left.structureKind ?? "flat") !== "flat" ? 500 : processChange ? 250 : 0);
    candidates.push({ beforeIndex, afterIndex, score, processChange });
  }));
  candidates.sort((a, b) => b.score - a.score || a.beforeIndex - b.beforeIndex || a.afterIndex - b.afterIndex);

  const usedBefore = new Set<number>();
  const usedAfter = new Set<number>();
  const pairs: Array<{ before?: BomItem; after?: BomItem; score: number; ambiguous: boolean; processChange?: BomProcessChange }> = [];
  for (const candidate of candidates) {
    if (usedBefore.has(candidate.beforeIndex) || usedAfter.has(candidate.afterIndex)) continue;
    const competing = candidates.some((other) => other !== candidate
      && (other.beforeIndex === candidate.beforeIndex || other.afterIndex === candidate.afterIndex)
      && other.score >= candidate.score * 0.9);
    usedBefore.add(candidate.beforeIndex);
    usedAfter.add(candidate.afterIndex);
    pairs.push({ before: before[candidate.beforeIndex], after: after[candidate.afterIndex], score: candidate.score, ambiguous: competing, processChange: candidate.processChange });
  }
  before.forEach((item, index) => { if (!usedBefore.has(index)) pairs.push({ before: item, score: 0, ambiguous: false }); });
  after.forEach((item, index) => { if (!usedAfter.has(index)) pairs.push({ after: item, score: 0, ambiguous: false }); });

  return pairs.map(({ before: a, after: b, score, ambiguous, processChange }) => {
    const ref = a?.ref ?? b?.ref ?? "";
    const addedParts = differenceParts(b?.alternatives ?? [], a?.alternatives ?? []);
    const removedParts = differenceParts(a?.alternatives ?? [], b?.alternatives ?? []);
    const newParts = addedParts.filter((part) => !beforePartKeys.has(partKey(part)));
    const deletedParts = removedParts.filter((part) => !afterPartKeys.has(partKey(part)));
    const addedPositions = differenceValues(b?.positions ?? [], a?.positions ?? []);
    const removedPositions = differenceValues(a?.positions ?? [], b?.positions ?? []);
    const replacementPositions = a && b && addedParts.length && removedParts.length
      ? intersectionValues(a.positions, b.positions)
      : [];
    const match = matchMetadata(a, b, score, ambiguous, processChange);
    if (!a) {
      return { ref, kind: "added", categories: ["added"], primaryType: "componentAdded", after: b, fields: compactFields(addedParts, [], addedPositions, []), addedParts, removedParts: [], newParts, deletedParts: [], addedPositions, removedPositions: [], replacementPositions: [], ...match };
    }
    if (!b) {
      return { ref, kind: "removed", categories: ["removed"], primaryType: "componentRemoved", before: a, fields: compactFields([], removedParts, [], removedPositions), addedParts: [], removedParts, newParts: [], deletedParts, addedPositions: [], removedPositions, replacementPositions: [], ...match };
    }

    let fields = compactFields(addedParts, removedParts, addedPositions, removedPositions, true, true);
    const quantityChanged = a.qty !== b.qty;
    const customerPartNumbersChanged = !overlapStats(
      a.rdCustomerPartNumbers ?? a.alternatives.flatMap((alternative) => alternative.rdCustomerPartNumbers ?? []),
      b.rdCustomerPartNumbers ?? b.alternatives.flatMap((alternative) => alternative.rdCustomerPartNumbers ?? []),
      normalizeValue,
    ).exact;
    if (customerPartNumbersChanged) fields.push("客戶料號差異");
    if (processChange) fields.push("製程別放置異常");
    if (replacementPositions.length) {
      fields = fields.filter((field) => field !== "新增替料" && field !== "刪除替料");
      fields.push("更換料號");
    }
    const substituteOnlyAddition = addedParts.length > 0 && removedParts.length === 0;
    const substituteOnlyRemoval = removedParts.length > 0 && addedParts.length === 0;
    if (newParts.length && !substituteOnlyAddition) fields.push("新增料號");
    if (deletedParts.length && !substituteOnlyRemoval) fields.push("刪除料號");
    const categories: DiffCategory[] = [];
    const positionRelocated = addedPositions.length > 0
      && removedPositions.length > 0
      && addedParts.length === 0
      && removedParts.length === 0
      && !quantityChanged;
    if (positionRelocated || processChange || customerPartNumbersChanged) {
      categories.push("changed");
    } else {
      if (addedParts.length || addedPositions.length) categories.push("added");
      if (removedParts.length || removedPositions.length) categories.push("removed");
      if (quantityChanged || replacementPositions.length) categories.push("changed");
    }
    const uniqueFields = [...new Set(fields)];
    const kind: DiffKind = !categories.length
      ? "same"
      : categories.includes("changed") || categories.length > 1
        ? "changed"
        : categories[0];
    const primaryType: DiffPrimaryType = processChange
      ? "positionChanged"
      : replacementPositions.length || (addedParts.length > 0 && removedParts.length > 0)
      ? "partReplaced"
      : addedParts.length > 0
        ? "substituteAdded"
        : removedParts.length > 0
          ? "substituteRemoved"
          : addedPositions.length > 0 || removedPositions.length > 0 || quantityChanged || customerPartNumbersChanged
            ? "positionChanged"
            : "same";
    return {
      ref,
      kind,
      categories,
      primaryType,
      before: a,
      after: b,
      fields: uniqueFields,
      addedParts,
      removedParts,
      newParts,
      deletedParts,
      addedPositions,
      removedPositions,
      replacementPositions,
      ...(processChange ? { processChange } : {}),
      ...match,
    };
  }).sort((a, b) => naturalSort(diffSortKey(a), diffSortKey(b)));
}

function matchMetadata(before: BomItem | undefined, after: BomItem | undefined, score: number, ambiguous: boolean, processChange?: BomProcessChange) {
  if (!before || !after) return { matchConfidence: "high" as const, matchReason: before ? "新版沒有可配對群組" : "舊版沒有可配對群組", needsReview: false };
  const samePositions = before.positions.length > 0 && overlapStats(before.positions, after.positions, normalizeValue).exact;
  const overlappingPositions = overlapStats(before.positions, after.positions, normalizeValue).intersection > 0;
  const sameParts = overlapStats(before.alternatives.map(partKey), after.alternatives.map(partKey), normalizeValue).intersection > 0;
  const confidence = ambiguous || score < 250 ? "low" : score < 700 ? "medium" : "high";
  const sameStructure = structureCompatible(before, after) && (before.structureKind ?? "flat") !== "flat";
  const reasons = [
    processChange ? `同一料號跨製程架構移動（${processChange.before} → ${processChange.after}）` : "",
    ambiguous ? "存在分數接近的多個候選群組" : "",
    !ambiguous && confidence === "low" ? "配對證據不足" : "",
    sameStructure ? "所屬架構相同" : "",
    samePositions ? "插件位置完全相同" : overlappingPositions ? "部分插件位置相同" : "",
    sameParts ? "料號重疊" : "",
  ].filter(Boolean);
  return {
    matchConfidence: confidence,
    matchReason: reasons.join("、") || "依群組相似度配對",
    needsReview: Boolean(processChange) || ambiguous || confidence === "low",
  };
}

function structureCompatible(before: BomItem, after: BomItem) {
  const beforeKind = before.structureKind ?? "flat";
  const afterKind = after.structureKind ?? "flat";
  if (beforeKind !== afterKind) return false;
  const beforeSection = before.structureKey?.split(":")[0];
  const afterSection = after.structureKey?.split(":")[0];
  return !beforeSection || !afterSection || beforeSection === afterSection;
}

function processPlacementChange(before: BomItem, after: BomItem): BomProcessChange | undefined {
  const beforeProcess = bomProcessKind(before);
  const afterProcess = bomProcessKind(after);
  if (!beforeProcess || !afterProcess || beforeProcess === afterProcess) return undefined;

  const beforeSection = before.structureKey?.split(":")[0];
  const afterSection = after.structureKey?.split(":")[0];
  if (beforeSection && afterSection && beforeSection !== afterSection) return undefined;

  const parts = overlapStats(
    before.alternatives.map((item) => canonicalPartNumber(item.part)),
    after.alternatives.map((item) => canonicalPartNumber(item.part)),
    normalizeValue,
  );
  if (!parts.intersection) return undefined;
  return { before: beforeProcess, after: afterProcess };
}

const primaryTypeSortOrder: Record<DiffPrimaryType, number> = {
  componentAdded: 0,
  substituteAdded: 1,
  componentRemoved: 2,
  substituteRemoved: 3,
  positionChanged: 4,
  partReplaced: 5,
  same: 6,
};

export function sortBomDiffsForAll(diffs: BomDiff[]) {
  return [...diffs].sort((a, b) => {
    const typeOrder = primaryTypeSortOrder[a.primaryType] - primaryTypeSortOrder[b.primaryType];
    return typeOrder || naturalSort(diffSortKey(a), diffSortKey(b));
  });
}

function groupMatchScore(before: BomItem, after: BomItem) {
  const position = overlapStats(before.positions, after.positions, normalizeValue);
  const alternatives = overlapStats(before.alternatives.map(altKey), after.alternatives.map(altKey), (value) => value);
  const parts = overlapStats(
    before.alternatives.map((item) => canonicalPartNumber(item.part)),
    after.alternatives.map((item) => canonicalPartNumber(item.part)),
    normalizeValue,
  );

  let score = 0;
  if (position.exact && before.positions.length) score += 1000;
  else if (position.intersection) score += 700 + position.ratio * 100 + position.intersection * 20;

  if (alternatives.exact && before.alternatives.length) score += 500;
  else if (alternatives.intersection) score += 300 + alternatives.ratio * 100;

  if (parts.exact && before.alternatives.length) score += 220;
  else if (parts.intersection) score += 120 + parts.ratio * 50;
  return score;
}

function overlapStats<T>(left: T[], right: T[], normalize: (value: T) => string) {
  const a = new Set(left.map(normalize).filter(Boolean));
  const b = new Set(right.map(normalize).filter(Boolean));
  const intersection = [...a].filter((value) => b.has(value)).length;
  const union = new Set([...a, ...b]).size;
  return {
    intersection,
    exact: a.size === b.size && intersection === a.size,
    ratio: union ? intersection / union : 0,
  };
}

function diffSortKey(diff: BomDiff) {
  return diff.after?.positions[0] ?? diff.before?.positions[0] ?? diff.after?.part ?? diff.before?.part ?? diff.ref;
}

function altKey(alternative: BomAlternative) {
  return `${normalizeValue(alternative.part)}\u0000${normalizeValue(alternative.manufacturerPart)}`;
}

function partKey(alternative: BomAlternative) {
  return normalizeValue(canonicalPartNumber(alternative.part));
}

function differenceParts(source: BomAlternative[], comparison: BomAlternative[]) {
  const comparisonKeys = new Set(comparison.map(partKey));
  const emitted = new Set<string>();
  return source.filter((alternative) => {
    const key = partKey(alternative);
    if (!key || comparisonKeys.has(key) || emitted.has(key)) return false;
    emitted.add(key);
    return true;
  });
}

function differenceValues(source: string[], comparison: string[]) {
  const comparisonValues = new Set(comparison.map(normalizeValue));
  return source.filter((value) => !comparisonValues.has(normalizeValue(value))).sort(naturalSort);
}

function intersectionValues(left: string[], right: string[]) {
  const rightValues = new Set(right.map(normalizeValue));
  return left.filter((value) => rightValues.has(normalizeValue(value))).sort(naturalSort);
}

function compactFields(addedParts: BomAlternative[], removedParts: BomAlternative[], addedPositions: string[], removedPositions: string[], addedAsSubstitute = false, removedAsSubstitute = false) {
  return [
    addedParts.length ? (addedAsSubstitute ? "新增替料" : "新增料號") : "",
    removedParts.length ? (removedAsSubstitute ? "刪除替料" : "刪除料號") : "",
    addedPositions.length ? "新增插件位置" : "",
    removedPositions.length ? "移除插件位置" : "",
  ].filter(Boolean);
}

function naturalSort(a: string, b: string) {
  return a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
}
