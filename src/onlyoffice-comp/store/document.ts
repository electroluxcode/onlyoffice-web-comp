import { getFileExt } from "../internal/editor/utils";

export type OnlyOfficeDocumentState = {
  isNew: boolean;
  fileName: string;
  fileType?: string;
  file?: File;
  url?: string;
  loader?: (url: string) => Promise<ArrayBuffer>;
};

export type SetOnlyOfficeDocumentStateInput = Omit<
  OnlyOfficeDocumentState,
  "isNew" | "fileName"
> & {
  isNew?: boolean;
  fileName: string;
};

let documentState: OnlyOfficeDocumentState = {
  isNew: true,
  fileName: "New Document.docx",
  fileType: "docx",
};

export function setDocmentObj(state: SetOnlyOfficeDocumentStateInput) {
  documentState = {
    ...state,
    isNew: state.isNew ?? !state.file,
    fileType: state.fileType || getFileExt(state.fileName) || "docx",
  };
}

export function getDocmentObj() {
  return documentState;
}

export function clearDocmentObj() {
  documentState = {
    isNew: true,
    fileName: "New Document.docx",
    fileType: "docx",
  };
}

export function setNewDocument(fileType = "docx") {
  setDocmentObj({
    isNew: true,
    fileName: `New Document.${fileType}`,
    fileType,
  });
}

export function setDocumentFile(file: File, fileName = file.name) {
  setDocmentObj({
    isNew: false,
    file,
    fileName,
    fileType: getFileExt(fileName) || getFileExt(file.name) || "docx",
  });
}

export function setDocumentUrl(
  url: string,
  {
    fileType,
    fileName,
    loader,
  }: {
    fileType?: string;
    fileName?: string;
    loader?: (url: string) => Promise<ArrayBuffer>;
  } = {},
) {
  const name = fileName || decodeURIComponent(url.split("/").pop() || "Document");

  setDocmentObj({
    isNew: false,
    url,
    loader,
    fileName: name,
    fileType: fileType || getFileExt(name) || "docx",
  });
}
