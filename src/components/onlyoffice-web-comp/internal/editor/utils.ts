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

function decodeCsvSample(buffer: ArrayBuffer, encoding: number) {
  const bytes = new Uint8Array(buffer);
  const sample = bytes.subarray(0, Math.min(bytes.length, 8192));
  const withoutBom =
    encoding === X2T_CSV_ENCODING_UTF8 &&
    sample.length >= 3 &&
    sample[0] === 0xef &&
    sample[1] === 0xbb &&
    sample[2] === 0xbf
      ? sample.subarray(3)
      : sample;

  if (encoding === X2T_CSV_ENCODING_UTF8) {
    return new TextDecoder("utf-8").decode(withoutBom);
  }

  try {
    return new TextDecoder("gbk").decode(withoutBom);
  } catch {
    return new TextDecoder("latin1").decode(withoutBom);
  }
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
