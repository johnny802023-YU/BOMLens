export type PdfTextBox = { page: number; text: string; x: number; y: number; width: number; height: number; rotation?: number };
export type PdfReferenceIndex = Map<string, PdfTextBox[]>;

const referencePattern = /^[A-Z]{1,6}\d+[A-Z]?$/;

export function normalizeReference(value: string) {
  return value.trim().toUpperCase().replace(/\s+/g, "");
}

export function findReferenceHits(boxes: PdfTextBox[], target: string) {
  const normalized = normalizeReference(target);
  if (!normalized) return [];
  return boxes.filter((box) => {
    const text = box.text.toUpperCase();
    const tokens = text.split(/[^A-Z0-9]+/).filter(Boolean);
    return tokens.includes(normalized) || normalizeReference(text) === normalized;
  });
}

function referenceTokens(text: string) {
  const direct = text.toUpperCase().match(/[A-Z]{1,6}\s*\d+\s*[A-Z]?/g) ?? [];
  return [...new Set(direct.map(normalizeReference).filter((value) => referencePattern.test(value)))];
}

function addHit(index: PdfReferenceIndex, reference: string, box: PdfTextBox) {
  const hits = index.get(reference) ?? [];
  const duplicate = hits.some((hit) => hit.page === box.page && Math.abs(hit.x - box.x) < 2 && Math.abs(hit.y - box.y) < 2);
  if (!duplicate) hits.push(box);
  index.set(reference, hits);
}

function isVertical(box: PdfTextBox) {
  const rotation = ((box.rotation ?? 0) % 360 + 360) % 360;
  return Math.abs(rotation - 90) <= 20 || Math.abs(rotation - 270) <= 20;
}

function closeEnough(left: PdfTextBox, right: PdfTextBox, vertical: boolean) {
  if (left.page !== right.page || isVertical(left) !== vertical || isVertical(right) !== vertical) return false;
  if (vertical) {
    const centerDelta = Math.abs((left.x + left.width / 2) - (right.x + right.width / 2));
    const verticalGap = right.y - (left.y + left.height);
    return centerDelta <= Math.max(left.width, right.width) * 0.85
      && verticalGap >= -2
      && verticalGap <= Math.max(20, Math.max(left.width, right.width) * 1.8);
  }
  const centerDelta = Math.abs((left.y + left.height / 2) - (right.y + right.height / 2));
  const horizontalGap = right.x - (left.x + left.width);
  return centerDelta <= Math.max(left.height, right.height) * 0.75
    && horizontalGap >= -2
    && horizontalGap <= Math.max(20, Math.max(left.height, right.height) * 1.8);
}

function addMergedReferences(index: PdfReferenceIndex, boxes: PdfTextBox[], vertical: boolean) {
  const ordered = [...boxes]
    .filter((box) => isVertical(box) === vertical)
    .sort((a, b) => vertical ? a.x - b.x || a.y - b.y : a.y - b.y || a.x - b.x);
  for (let start = 0; start < ordered.length; start += 1) {
    let previous = ordered[start];
    const texts = [previous.text];
    let left = previous.x;
    let top = previous.y;
    let right = previous.x + previous.width;
    let bottom = previous.y + previous.height;
    for (let end = start + 1; end < Math.min(ordered.length, start + 3); end += 1) {
      const current = ordered[end];
      if (!closeEnough(previous, current, vertical)) break;
      texts.push(current.text);
      left = Math.min(left, current.x); top = Math.min(top, current.y);
      right = Math.max(right, current.x + current.width); bottom = Math.max(bottom, current.y + current.height);
      const candidates = [normalizeReference(texts.join("")), normalizeReference([...texts].reverse().join(""))];
      candidates.filter((reference, candidateIndex) => candidateIndex === 0 || reference !== candidates[0]).forEach((reference) => {
        if (referencePattern.test(reference)) addHit(index, reference, { page: current.page, text: reference, x: left, y: top, width: right - left, height: bottom - top, rotation: vertical ? current.rotation : 0 });
      });
      previous = current;
    }
  }
}

export function buildPageReferenceIndex(boxes: PdfTextBox[]) {
  const index: PdfReferenceIndex = new Map();
  boxes.forEach((box) => referenceTokens(box.text).forEach((reference) => addHit(index, reference, box)));

  addMergedReferences(index, boxes, false);
  addMergedReferences(index, boxes, true);
  return index;
}

export function mergeReferenceIndex(target: PdfReferenceIndex, source: PdfReferenceIndex) {
  source.forEach((boxes, reference) => boxes.forEach((box) => addHit(target, reference, box)));
  return target;
}

export function lookupReferenceHits(index: PdfReferenceIndex, target: string) {
  return index.get(normalizeReference(target)) ?? [];
}
