import { getDocumentType, getNewUrl } from "../../const";
import {
  AvsFileType,
  DocumentType,
  X2T_CSV_DELIMITER_COMMA,
  X2T_CSV_DELIMITER_SEMICOLON,
  X2T_CSV_DELIMITER_TAB,
  X2T_CSV_ENCODING_GBK,
  X2T_CSV_ENCODING_UTF8,
} from "./types";

export { getDocumentType, getNewUrl };

export function getFileExt(name: string) {
  const type = name.split(".").pop() || "";
  return type.toLowerCase();
}

const x2tSourceFormatByExt: Record<string, AvsFileType> = {
  docx: AvsFileType.AVS_FILE_DOCUMENT_DOCX,
  doc: AvsFileType.AVS_FILE_DOCUMENT_DOC,
  odt: AvsFileType.AVS_FILE_DOCUMENT_ODT,
  rtf: AvsFileType.AVS_FILE_DOCUMENT_RTF,
  txt: AvsFileType.AVS_FILE_DOCUMENT_TXT,
  docm: AvsFileType.AVS_FILE_DOCUMENT_DOCM,
  dotx: AvsFileType.AVS_FILE_DOCUMENT_DOTX,
  dotm: AvsFileType.AVS_FILE_DOCUMENT_DOTM,
  xlsx: AvsFileType.AVS_FILE_SPREADSHEET_XLSX,
  xls: AvsFileType.AVS_FILE_SPREADSHEET_XLS,
  ods: AvsFileType.AVS_FILE_SPREADSHEET_ODS,
  csv: AvsFileType.AVS_FILE_SPREADSHEET_CSV,
  xlsm: AvsFileType.AVS_FILE_SPREADSHEET_XLSM,
  pptx: AvsFileType.AVS_FILE_PRESENTATION_PPTX,
  ppt: AvsFileType.AVS_FILE_PRESENTATION_PPT,
  odp: AvsFileType.AVS_FILE_PRESENTATION_ODP,
  pdf: AvsFileType.AVS_FILE_CROSSPLATFORM_PDF,
};

function getX2tBinFormat(fileType: string) {
  switch (getDocumentType(fileType)) {
    case DocumentType.Cell:
      return AvsFileType.AVS_FILE_CANVAS_SPREADSHEET;
    case DocumentType.Slide:
      return AvsFileType.AVS_FILE_CANVAS_PRESENTATION;
    case DocumentType.Draw:
      return AvsFileType.AVS_FILE_CANVAS + 0x0005;
    default:
      return AvsFileType.AVS_FILE_CANVAS_WORD;
  }
}

export function getX2tConvertFormats(fileType: string) {
  const ext = getFileExt(fileType);
  const formatFrom =
    x2tSourceFormatByExt[ext] ?? AvsFileType.AVS_FILE_DOCUMENT_DOCX;

  return {
    formatFrom,
    formatTo: getX2tBinFormat(fileType),
  };
}

function isValidUtf8(bytes: Uint8Array) {
  try {
    new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    return true;
  } catch {
    return false;
  }
}

function decodeCsvBytes(bytes: Uint8Array, encoding: number) {
  const withoutBom =
    encoding === X2T_CSV_ENCODING_UTF8 &&
    bytes.length >= 3 &&
    bytes[0] === 0xef &&
    bytes[1] === 0xbb &&
    bytes[2] === 0xbf
      ? bytes.subarray(3)
      : bytes;

  if (encoding === X2T_CSV_ENCODING_UTF8) {
    return new TextDecoder("utf-8").decode(withoutBom);
  }

  try {
    return new TextDecoder("gbk").decode(withoutBom);
  } catch {
    return new TextDecoder("latin1").decode(withoutBom);
  }
}

function decodeCsvSample(buffer: ArrayBuffer, encoding: number) {
  const bytes = new Uint8Array(buffer);
  const sample = bytes.subarray(0, Math.min(bytes.length, 8192));
  return decodeCsvBytes(sample, encoding);
}

function decodeCsvBuffer(buffer: ArrayBuffer, encoding: number) {
  return decodeCsvBytes(new Uint8Array(buffer), encoding);
}

function encodeCsvBuffer(text: string, withUtf8Bom: boolean) {
  const encoded = new TextEncoder().encode(text);
  if (!withUtf8Bom) {
    return encoded.slice().buffer;
  }
  const withBom = new Uint8Array(encoded.length + 3);
  withBom.set([0xef, 0xbb, 0xbf], 0);
  withBom.set(encoded, 3);
  return withBom.slice().buffer;
}

function parseCsvLine(line: string, delimiter: string) {
  return parseCsvText(line, delimiter)[0] ?? [];
}

/** RFC 4180 风格解析，支持引号内换行与逗号。 */
function parseCsvText(text: string, delimiter: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];

    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
      continue;
    }

    if (ch === '"') {
      inQuotes = true;
      continue;
    }

    if (ch === delimiter) {
      row.push(field);
      field = "";
      continue;
    }

    if (ch === "\r") {
      row.push(field);
      field = "";
      rows.push(row);
      row = [];
      if (text[i + 1] === "\n") {
        i++;
      }
      continue;
    }

    if (ch === "\n") {
      row.push(field);
      field = "";
      rows.push(row);
      row = [];
      continue;
    }

    field += ch;
  }

  row.push(field);
  if (row.length > 1 || row[0] !== "") {
    rows.push(row);
  }

  return rows;
}

function serializeCsvRow(fields: string[], delimiter: string) {
  return fields.map((field) => serializeCsvField(field, delimiter)).join(delimiter);
}

function serializeCsvField(field: string, delimiter: string) {
  if (
    field.includes(delimiter) ||
    field.includes('"') ||
    field.includes("\n") ||
    field.includes("\r")
  ) {
    return `"${field.replace(/"/g, '""')}"`;
  }
  return field;
}

function toCsvFormulaCell(value: string) {
  if (/^=".*"$/.test(value)) {
    return value;
  }
  return `="${value.replace(/"/g, '""')}"`;
}

/** 仅包裹「以数字结尾且不像日期/时间」的单元格，避免误包 2018/9/6 9:32 引发 x2t 异常。 */
function shouldApplyCsvFormulaWrap(value: string) {
  if (!/[0-9]$/.test(value)) {
    return false;
  }
  if (/^=".*"$/.test(value)) {
    return false;
  }
  if (/[/\-:.Tt]/.test(value)) {
    return false;
  }
  return true;
}

function getCsvDelimiterChar(delimiter: number) {
  if (delimiter === X2T_CSV_DELIMITER_TAB) return "\t";
  if (delimiter === X2T_CSV_DELIMITER_SEMICOLON) return ";";
  return ",";
}

/** 引号内换行会导致按行 split 破坏结构；物理行数明显多于逻辑行数即视为复杂 CSV。 */
export function isMultilineCsv(buffer: ArrayBuffer) {
  const csvEncoding = detectX2tCsvEncoding(buffer);
  const delimiter = getCsvDelimiterChar(
    detectX2tCsvDelimiter(buffer, csvEncoding),
  );
  const text = decodeCsvBuffer(buffer, csvEncoding);
  const physicalLines = text.split(/\r?\n/).filter((line) => line.length > 0).length;
  const logicalRows = parseCsvText(text, delimiter).length;
  return physicalLines > logicalRows + 2;
}

/**
 * 复杂 CSV（多行单元格、嵌套逗号）走 XLSX 再 x2t，避免 x2t CSV 解析崩溃。
 */
export async function convertCsvBufferToXlsxBuffer(buffer: ArrayBuffer) {
  const ExcelJS = (await import("exceljs")).default;
  const csvEncoding = detectX2tCsvEncoding(buffer);
  const delimiter = getCsvDelimiterChar(
    detectX2tCsvDelimiter(buffer, csvEncoding),
  );
  const rows = parseCsvText(decodeCsvBuffer(buffer, csvEncoding), delimiter);
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Sheet1");

  for (const row of rows) {
    worksheet.addRow(row);
  }

  const output = await workbook.xlsx.writeBuffer();
  if (output instanceof ArrayBuffer) {
    return output;
  }
  const bytes = new Uint8Array(output as ArrayLike<number>);
  return bytes.slice().buffer;
}

/**
 * x2t CSV 解析 bug：某列单元格以 ASCII 数字结尾时，下一列会误走 DateReader 并崩溃
 *（如 login1.csv 的「用户3」+「2018/9/6 9:32」）。用 ="value" 包裹前一格可绕过。
 */
export function sanitizeCsvBufferForX2t(buffer: ArrayBuffer) {
  const csvEncoding = detectX2tCsvEncoding(buffer);
  const csvDelimiter = detectX2tCsvDelimiter(buffer, csvEncoding);
  const delimiter = getCsvDelimiterChar(csvDelimiter);
  const bytes = new Uint8Array(buffer);
  const withUtf8Bom =
    csvEncoding === X2T_CSV_ENCODING_UTF8 &&
    bytes.length >= 3 &&
    bytes[0] === 0xef &&
    bytes[1] === 0xbb &&
    bytes[2] === 0xbf;
  const text = decodeCsvBuffer(buffer, csvEncoding);
  const lineEnding = text.includes("\r\n") ? "\r\n" : "\n";
  const rows = parseCsvText(text, delimiter);

  const sanitized = rows.map((fields) => {
    for (let i = 0; i < fields.length - 1; i++) {
      if (shouldApplyCsvFormulaWrap(fields[i])) {
        fields[i] = toCsvFormulaCell(fields[i]);
      }
    }
    return serializeCsvRow(fields, delimiter);
  });

  return encodeCsvBuffer(sanitized.join(lineEnding), withUtf8Bom);
}

function getFirstCsvLine(text: string) {
  const newline = text.search(/\r?\n/);
  return newline === -1 ? text : text.slice(0, newline);
}

/** Detect delimiter from the first CSV row. */
export function detectX2tCsvDelimiter(buffer: ArrayBuffer, encoding: number) {
  const line = getFirstCsvLine(decodeCsvSample(buffer, encoding));
  const counts = { comma: 0, semicolon: 0, tab: 0 };
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        i++;
        continue;
      }
      inQuotes = !inQuotes;
      continue;
    }
    if (inQuotes) continue;
    if (ch === ",") counts.comma++;
    else if (ch === ";") counts.semicolon++;
    else if (ch === "\t") counts.tab++;
  }

  if (counts.tab > counts.comma && counts.tab > counts.semicolon) {
    return X2T_CSV_DELIMITER_TAB;
  }
  if (counts.semicolon > counts.comma) {
    return X2T_CSV_DELIMITER_SEMICOLON;
  }
  return X2T_CSV_DELIMITER_COMMA;
}

/** Detect OnlyOffice x2t CSV encoding index from BOM / byte patterns. */
export function detectX2tCsvEncoding(buffer: ArrayBuffer) {
  const bytes = new Uint8Array(buffer);
  if (
    bytes.length >= 3 &&
    bytes[0] === 0xef &&
    bytes[1] === 0xbb &&
    bytes[2] === 0xbf
  ) {
    return X2T_CSV_ENCODING_UTF8;
  }
  return isValidUtf8(bytes) ? X2T_CSV_ENCODING_UTF8 : X2T_CSV_ENCODING_GBK;
}

export function getX2tCsvConvertOptions(buffer: ArrayBuffer) {
  const csvEncoding = detectX2tCsvEncoding(buffer);
  return {
    csvEncoding,
    csvDelimiter: detectX2tCsvDelimiter(buffer, csvEncoding),
  };
}

export function getX2tExportFormats(fileType: string) {
  const ext = getFileExt(fileType);
  const formatTo =
    x2tSourceFormatByExt[ext] ?? AvsFileType.AVS_FILE_DOCUMENT_DOCX;

  return {
    formatFrom: getX2tBinFormat(fileType),
    formatTo,
  };
}
