import { bomQuantityChanged, type BomDiff } from "./bom-logic.ts";
import { normalizeReference } from "./pdf-search.ts";

export type SchematicReportPlanEntry = {
  reference: string;
  changeLabels: string[];
  detail: string;
};

function unique(values: string[]) {
  return [...new Set(values.map(normalizeReference).filter(Boolean))];
}

function diffPositions(diff: BomDiff) {
  const direct = unique([...diff.addedPositions, ...diff.removedPositions, ...diff.replacementPositions]);
  if (direct.length) return direct;
  const requiresContext = diff.addedParts.length > 0
    || diff.removedParts.length > 0
    || diff.before?.qty !== diff.after?.qty
    || Boolean(diff.processChange);
  return requiresContext ? unique([...(diff.before?.positions ?? []), ...(diff.after?.positions ?? [])]) : [];
}

function entryLabels(diff: BomDiff, reference: string) {
  const labels: string[] = [];
  if (diff.addedPositions.some((position) => normalizeReference(position) === reference)) labels.push("新增插件位置");
  if (diff.removedPositions.some((position) => normalizeReference(position) === reference)) labels.push("移除插件位置");
  if (diff.replacementPositions.some((position) => normalizeReference(position) === reference)) labels.push("更換料號");
  if (diff.addedParts.length) labels.push(diff.primaryType === "substituteAdded" ? "新增替料" : "新增料號");
  if (diff.removedParts.length) labels.push(diff.primaryType === "substituteRemoved" ? "刪除替料" : "刪除料號");
  if (diff.before && diff.after && bomQuantityChanged(diff.before, diff.after)) labels.push("數量差異");
  if (diff.processChange) labels.push(`製程別放置異常 ${diff.processChange.before} → ${diff.processChange.after}`);
  if (!labels.length) labels.push(...diff.fields);
  return [...new Set(labels)];
}

function partSummary(diff: BomDiff) {
  const before = diff.before?.alternatives.map((part) => part.part || part.manufacturerPart).filter(Boolean).join("、") || "無";
  const after = diff.after?.alternatives.map((part) => part.part || part.manufacturerPart).filter(Boolean).join("、") || "無";
  return `舊版 ${before} → 新版 ${after}`;
}

export function createSchematicReportPlan(diffs: BomDiff[]) {
  const entries = new Map<string, SchematicReportPlanEntry>();
  diffs.filter((diff) => diff.kind !== "same").forEach((diff) => {
    diffPositions(diff).forEach((reference) => {
      const current = entries.get(reference);
      const labels = entryLabels(diff, reference);
      if (current) {
        current.changeLabels = [...new Set([...current.changeLabels, ...labels])];
        if (!current.detail.includes(partSummary(diff))) current.detail += `；${partSummary(diff)}`;
      } else {
        entries.set(reference, { reference, changeLabels: labels, detail: partSummary(diff) });
      }
    });
  });
  return [...entries.values()].sort((left, right) => left.reference.localeCompare(right.reference, undefined, { numeric: true, sensitivity: "base" }));
}
