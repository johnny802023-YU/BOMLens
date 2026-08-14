import * as XLSX from "xlsx";

export type WorkbookSheet = { name: string; matrix: unknown[][] };
export type WorkbookSourceFormat = "excel-binary" | "excel-openxml" | "html-xls" | "xml-xls" | "delimited-text";

export type ParsedWorkbook = {
  book: XLSX.WorkBook;
  sheets: WorkbookSheet[];
  sourceFormat: WorkbookSourceFormat;
};

const OLE_MAGIC = [0xd0, 0xcf, 0x11, 0xe0];
const ZIP_MAGIC = [0x50, 0x4b];

function startsWith(bytes: Uint8Array, signature: number[]) {
  return signature.every((value, index) => bytes[index] === value);
}

function decodeWorkbookText(bytes: Uint8Array) {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder("utf-16le").decode(bytes);
  if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    const swapped = new Uint8Array(bytes.length - 2);
    for (let index = 2; index + 1 < bytes.length; index += 2) {
      swapped[index - 2] = bytes[index + 1];
      swapped[index - 1] = bytes[index];
    }
    return new TextDecoder("utf-16le").decode(swapped);
  }
  const utf8 = new TextDecoder("utf-8").decode(bytes);
  const charset = utf8.slice(0, 4096).match(/charset\s*=\s*["']?([^\s"'/>;]+)/i)?.[1]?.toLowerCase();
  if (charset === "big5" || charset === "big-5") {
    try { return new TextDecoder("big5").decode(bytes); } catch { return utf8; }
  }
  return utf8;
}

function sourceFormat(bytes: Uint8Array, text: string): WorkbookSourceFormat {
  if (startsWith(bytes, OLE_MAGIC)) return "excel-binary";
  if (startsWith(bytes, ZIP_MAGIC)) return "excel-openxml";
  const normalized = text.trimStart().toLowerCase();
  if (normalized.startsWith("<") && (normalized.includes("<html") || normalized.includes("<table"))) return "html-xls";
  if (normalized.startsWith("<?xml") || normalized.includes("urn:schemas-microsoft-com:office:spreadsheet")) return "xml-xls";
  return "delimited-text";
}

function ensureWorkbook(book: XLSX.WorkBook) {
  if (!book.SheetNames.length) throw new Error("活頁簿沒有可讀取的工作表。");
  return book;
}

export function parseWorkbookData(data: ArrayBuffer): ParsedWorkbook {
  const bytes = new Uint8Array(data);
  const text = startsWith(bytes, OLE_MAGIC) || startsWith(bytes, ZIP_MAGIC) ? "" : decodeWorkbookText(bytes);
  const format = sourceFormat(bytes, text);
  const readOptions = { cellStyles: true, cellNF: true, cellDates: true } as const;
  const book = ensureWorkbook(format === "html-xls" || format === "xml-xls" || format === "delimited-text"
    ? XLSX.read(text, { ...readOptions, type: "string" })
    : XLSX.read(data, { ...readOptions, type: "array" }));
  return {
    book,
    sourceFormat: format,
    sheets: book.SheetNames.map((name) => ({
      name,
      matrix: XLSX.utils.sheet_to_json<unknown[]>(book.Sheets[name], { header: 1, defval: "", raw: false }),
    })),
  };
}
