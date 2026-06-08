import { getDocumentType, getNewUrl } from "../../const";
import { AvsFileType, DocumentType } from "./types";

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

export function getX2tExportFormats(fileType: string) {
  const ext = getFileExt(fileType);
  const formatTo =
    x2tSourceFormatByExt[ext] ?? AvsFileType.AVS_FILE_DOCUMENT_DOCX;

  return {
    formatFrom: getX2tBinFormat(fileType),
    formatTo,
  };
}
