export type BomAlternative = {
  part: string;
  manufacturerPart: string;
  manufacturerName?: string;
  description: string;
  spec: string;
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
  matchConfidence: "high" | "medium" | "low";
  matchReason: string;
  needsReview: boolean;
};

export type CompanyColumnKey = "ref" | "part" | "qty" | "positions" | "description" | "spec" | "manufacturerName" | "manufacturerPart";
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
  part: ["part", "part number", "part no", "pn", "料號", "零件料號", "公司料號"],
  manufacturerPart: ["mpn", "manufacturer part number", "製造商料號", "製造商型號"],
  manufacturerName: ["manufacturer", "manufacturer name", "mfr", "製造商", "製造商名稱"],
  value: ["value", "component value", "規格", "數值", "值"],
  description: ["description", "desc", "comment", "item description", "說明", "描述", "品名"],
  qty: ["qty", "quantity", "count", "數量", "用量", "組成用量"],
  positions: ["placement", "placements", "location", "locations", "插件位置", "位置"],
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
};

export const companyColumnLabels: Record<CompanyColumnKey, string> = {
  ref: "項次",
  part: "料號",
  qty: "數量／組成用量",
  positions: "插件位置",
  description: "品名／描述",
  spec: "規格",
  manufacturerName: "製造商名稱",
  manufacturerPart: "製造商料號",
};

export const requiredCompanyColumns: CompanyColumnKey[] = ["ref", "part", "qty", "positions"];

const companyHeaderAliases = {
  ref: ["項次", "序號", "item", "item no", "item number"],
  part: ["料號", "公司料號", "零件料號", "part", "part no", "part number", "pn"],
  qty: ["組成用量", "數量", "用量", "qty", "quantity", "count"],
  positions: ["插件位置", "位號", "reference designator", "refdes", "placement", "placements", "location", "locations"],
  description: ["品名", "描述", "說明", "description", "desc", "item description"],
  spec: ["規格", "數值", "值", "spec", "specification", "value"],
  manufacturerName: ["製造商名稱", "製造商", "manufacturer name", "manufacturer", "mfr"],
  manufacturerPart: ["製造商料號", "製造商型號", "manufacturer part number", "mpn"],
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
    .split(/[,，;；\s]+/)
    .map((position) => position.trim().toUpperCase())
    .filter(Boolean))];
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
  };

  const groups = new Map<string, GroupDraft>();
  let currentRef = "";

  for (const [offset, row] of matrix.slice(headerIndex + 1).entries()) {
    const sourceRow = headerIndex + offset + 2;
    const nextRef = text(cell(row, columns.ref));
    if (nextRef) currentRef = nextRef;
    if (!currentRef) continue;

    const part = canonicalPartNumber(cell(row, columns.part));
    const description = text(cell(row, columns.description));
    const spec = text(cell(row, columns.spec));
    const rawQuantity = parseQuantity(cell(row, columns.qty));
    const positions = parsePositions(cell(row, columns.positions));
    const manufacturerName = text(cell(row, columns.manufacturerName));
    const manufacturerPart = text(cell(row, columns.manufacturerPart));

    if (!part && !manufacturerPart && !positions.length && !rawQuantity) continue;

    const key = normalizeValue(currentRef);
    const group = groups.get(key) ?? {
      ref: currentRef,
      alternatives: [],
      positions: new Set<string>(),
      rawQuantity: 0,
      sourceRows: [],
    };

    if (part || manufacturerPart) {
      const alternative = { part, manufacturerPart, manufacturerName, description, spec };
      const alternativeKey = altKey(alternative);
      if (!group.alternatives.some((candidate) => altKey(candidate) === alternativeKey)) {
        group.alternatives.push(alternative);
      }
    }
    positions.forEach((position) => group.positions.add(position));
    group.rawQuantity = Math.max(group.rawQuantity, rawQuantity);
    group.sourceRows.push(sourceRow);
    groups.set(key, group);
  }

  return [...groups.values()].map((group) => {
    const positions = [...group.positions].sort(naturalSort);
    const first = group.alternatives[0] ?? { part: "", manufacturerPart: "", manufacturerName: "", description: "", spec: "" };
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

  const duplicatePositions = new Map<string, Array<{ ref: string; rows: number[] }>>();
  items.forEach((item) => item.positions.forEach((position) => {
    const uses = duplicatePositions.get(position) ?? [];
    uses.push({ ref: item.ref, rows: item.sourceRows ?? [] });
    duplicatePositions.set(position, uses);
  }));
  duplicatePositions.forEach((uses, position) => {
    if (uses.length > 1) issues.push({ severity: "warning", code: "duplicate-position", message: `${position} 同時出現在 ${uses.map((use) => use.ref).join("、")}，請確認是否重複分配。`, rows: uses.flatMap((use) => use.rows) });
  });

  if (isCompleteCompanyMapping(mapping)) {
    const rawParts = new Map<string, Set<string>>();
    const quantityByRef = new Map<string, { qty: number; rows: number[] }>();
    let currentRef = "";
    matrix.slice(headerIndex + 1).forEach((row, offset) => {
      const rowNumber = headerIndex + offset + 2;
      const nextRef = text(row[mapping.ref]);
      if (nextRef) currentRef = nextRef;
      const rawPart = text(row[mapping.part]).replace(/\s+/g, "");
      const canonical = canonicalPartNumber(rawPart);
      if (canonical && rawPart) {
        const variants = rawParts.get(normalizeValue(canonical)) ?? new Set<string>();
        variants.add(normalizeValue(rawPart));
        rawParts.set(normalizeValue(canonical), variants);
      }
      if (currentRef) {
        const qty = parseQuantity(row[mapping.qty]);
        const previous = quantityByRef.get(normalizeValue(currentRef)) ?? { qty: 0, rows: [] };
        previous.qty = Math.max(previous.qty, qty);
        previous.rows.push(rowNumber);
        quantityByRef.set(normalizeValue(currentRef), previous);
      }
    });
    rawParts.forEach((variants, canonical) => {
      if (variants.size > 1) issues.push({ severity: "warning", code: "part-collision", message: `多個原始料號取最右 12 碼後都變成 ${canonical}：${[...variants].join("、")}` });
    });
    items.forEach((item) => {
      if (!item.alternatives.some((part) => part.part)) issues.push({ severity: "warning", code: "missing-part", message: `項次 ${item.ref} 沒有料號。`, rows: item.sourceRows });
      const source = quantityByRef.get(normalizeValue(item.ref));
      if (item.positions.length && source?.qty && item.positions.length !== source.qty) {
        issues.push({ severity: "warning", code: "quantity-mismatch", message: `項次 ${item.ref} 的數量為 ${source.qty}，但插件位置共有 ${item.positions.length} 個。`, rows: source.rows });
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
      positionCount: new Set(items.flatMap((item) => item.positions)).size,
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
  const candidates: Array<{ beforeIndex: number; afterIndex: number; score: number }> = [];
  before.forEach((left, beforeIndex) => after.forEach((right, afterIndex) => {
    const score = groupMatchScore(left, right);
    if (score > 0) candidates.push({ beforeIndex, afterIndex, score });
  }));
  candidates.sort((a, b) => b.score - a.score || a.beforeIndex - b.beforeIndex || a.afterIndex - b.afterIndex);

  const usedBefore = new Set<number>();
  const usedAfter = new Set<number>();
  const pairs: Array<{ before?: BomItem; after?: BomItem; score: number; ambiguous: boolean }> = [];
  for (const candidate of candidates) {
    if (usedBefore.has(candidate.beforeIndex) || usedAfter.has(candidate.afterIndex)) continue;
    const competing = candidates.some((other) => other !== candidate
      && (other.beforeIndex === candidate.beforeIndex || other.afterIndex === candidate.afterIndex)
      && other.score >= candidate.score * 0.9);
    usedBefore.add(candidate.beforeIndex);
    usedAfter.add(candidate.afterIndex);
    pairs.push({ before: before[candidate.beforeIndex], after: after[candidate.afterIndex], score: candidate.score, ambiguous: competing });
  }
  before.forEach((item, index) => { if (!usedBefore.has(index)) pairs.push({ before: item, score: 0, ambiguous: false }); });
  after.forEach((item, index) => { if (!usedAfter.has(index)) pairs.push({ after: item, score: 0, ambiguous: false }); });

  return pairs.map(({ before: a, after: b, score, ambiguous }) => {
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
    const match = matchMetadata(a, b, score, ambiguous);
    if (!a) {
      return { ref, kind: "added", categories: ["added"], primaryType: "componentAdded", after: b, fields: compactFields(addedParts, [], addedPositions, []), addedParts, removedParts: [], newParts, deletedParts: [], addedPositions, removedPositions: [], replacementPositions: [], ...match };
    }
    if (!b) {
      return { ref, kind: "removed", categories: ["removed"], primaryType: "componentRemoved", before: a, fields: compactFields([], removedParts, [], removedPositions), addedParts: [], removedParts, newParts: [], deletedParts, addedPositions: [], removedPositions, replacementPositions: [], ...match };
    }

    let fields = compactFields(addedParts, removedParts, addedPositions, removedPositions, true, true);
    const quantityChanged = a.qty !== b.qty;
    if (quantityChanged) fields.push("數量差異");
    if (replacementPositions.length) {
      fields = fields.filter((field) => field !== "新增替料" && field !== "刪除替料");
      fields.push("更換料號");
    }
    if (newParts.length) fields.push("新增料號");
    if (deletedParts.length) fields.push("刪除料號");
    const categories: DiffCategory[] = [];
    const positionRelocated = addedPositions.length > 0
      && removedPositions.length > 0
      && addedParts.length === 0
      && removedParts.length === 0
      && !quantityChanged;
    if (positionRelocated) {
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
    const primaryType: DiffPrimaryType = replacementPositions.length || (addedParts.length > 0 && removedParts.length > 0)
      ? "partReplaced"
      : addedParts.length > 0
        ? "substituteAdded"
        : removedParts.length > 0
          ? "substituteRemoved"
          : addedPositions.length > 0 || removedPositions.length > 0 || quantityChanged
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
      ...match,
    };
  }).sort((a, b) => naturalSort(diffSortKey(a), diffSortKey(b)));
}

function matchMetadata(before: BomItem | undefined, after: BomItem | undefined, score: number, ambiguous: boolean) {
  if (!before || !after) return { matchConfidence: "high" as const, matchReason: before ? "新版沒有可配對群組" : "舊版沒有可配對群組", needsReview: false };
  const samePositions = before.positions.length > 0 && overlapStats(before.positions, after.positions, normalizeValue).exact;
  const overlappingPositions = overlapStats(before.positions, after.positions, normalizeValue).intersection > 0;
  const sameParts = overlapStats(before.alternatives.map(partKey), after.alternatives.map(partKey), normalizeValue).intersection > 0;
  const confidence = ambiguous || score < 250 ? "low" : score < 700 ? "medium" : "high";
  const reasons = [samePositions ? "插件位置完全相同" : overlappingPositions ? "部分插件位置相同" : "", sameParts ? "料號重疊" : ""].filter(Boolean);
  return {
    matchConfidence: confidence,
    matchReason: reasons.join("、") || "依群組相似度配對",
    needsReview: ambiguous || confidence === "low",
  };
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
    addedParts.length ? (addedAsSubstitute ? "新增替料" : "新增元件") : "",
    removedParts.length ? (removedAsSubstitute ? "刪除替料" : "移除元件") : "",
    addedPositions.length ? "新增插件位置" : "",
    removedPositions.length ? "移除插件位置" : "",
  ].filter(Boolean);
}

function naturalSort(a: string, b: string) {
  return a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" });
}
