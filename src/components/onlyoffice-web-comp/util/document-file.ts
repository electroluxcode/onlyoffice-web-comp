import type { EditorManager } from "../core/editor-manager";
import { FILE_TYPE, type FileType } from "../const";
import { convertBinToDocument } from "./x2t";

const ZIP_SIGNATURE = 0x04034b50; // PK\x03\x04

export function bufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  const chunkSize = 8192;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, Math.min(i + chunkSize, bytes.length));
    for (let j = 0; j < chunk.length; j++) {
      binary += String.fromCharCode(chunk[j]);
    }
  }
  return btoa(binary);
}

export function base64ToBuffer(base64: string): ArrayBuffer {
  const normalized = base64.trim().replace(/\s/g, "");
  const binary = atob(normalized);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
}

export function getOnlyOfficeMimeType(fileExtension?: string): string {
  const extension = (fileExtension || "").toLowerCase();
  if (["xlsx", "xls"].includes(extension)) {
    return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  }
  if (extension === "csv") return "text/csv";
  if (extension === "docm") return "application/vnd.ms-word.document.macroEnabled.12";
  if (["docx"].includes(extension)) {
    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document;charset=UTF-8";
  }
  if (["pptx", "ppt"].includes(extension)) {
    return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  }
  if (extension === "txt") return "text/plain";
  return "application/octet-stream";
}

function getFileTypeConstant(fileExtension?: string): FileType {
  const extension = (fileExtension || "").toLowerCase();
  if (["docx", "docm", "doc"].includes(extension)) return FILE_TYPE.DOCX;
  if (["pptx", "ppt"].includes(extension)) return FILE_TYPE.PPTX;
  return FILE_TYPE.XLSX;
}

function resolveBinaryBuffer(fileData: ArrayBuffer, extension: string): ArrayBuffer {
  const isTextFile = ["csv", "txt"].includes(extension);
  if (isTextFile) return fileData;

  if (fileData.byteLength >= 4) {
    const view = new DataView(fileData, 0, 4);
    const signature = view.getUint32(0, true);
    if (signature === ZIP_SIGNATURE) return fileData;

    try {
      const text = new TextDecoder("utf-8").decode(fileData);
      const base64Regex = /^[A-Za-z0-9+/]*={0,2}$/;
      if (base64Regex.test(text.trim())) {
        return base64ToBuffer(text);
      }
    } catch {
      // fall through
    }
  }

  return fileData;
}

/** 将 useFileData 下载的数据转为 OnlyOffice 可打开的 File */
export function createFileFromDownloadedBuffer(
  fileData: ArrayBuffer | string | BlobPart,
  fileName: string,
  fileExtension = "docx",
): File {
  const extension = fileExtension.toLowerCase();
  const mimeType = getOnlyOfficeMimeType(extension);
  const isTextFile = ["csv", "txt"].includes(extension);

  if (fileData instanceof ArrayBuffer) {
    if (isTextFile) {
      const text = new TextDecoder("utf-8").decode(fileData);
      return new File([text], fileName, { type: mimeType });
    }
    const finalBuffer = resolveBinaryBuffer(fileData, extension);
    return new File([finalBuffer], fileName, { type: mimeType });
  }

  if (typeof fileData === "string") {
    return new File([fileData], fileName, { type: mimeType });
  }

  return new File([fileData], fileName, { type: mimeType });
}

/** 导出编辑器文档并转为可上传的 File */
export async function exportEditorDocumentAsFile(
  manager: EditorManager,
  fileName: string,
  fileExtension = "docx",
): Promise<File> {
  const exported = await manager.export();
  const fileTypeConstant = getFileTypeConstant(fileExtension);
  const result = await convertBinToDocument(
    exported.binData,
    exported.fileName || fileName,
    fileTypeConstant,
    exported.media,
  );
  const mimeType = getOnlyOfficeMimeType(fileExtension);
  const bytes = new Uint8Array(result.data);
  return new File([bytes], fileName || result.fileName, { type: mimeType });
}

/** 导出编辑器文档并转为 base64 字符串 */
export async function exportEditorDocumentAsBase64(
  manager: EditorManager,
  fileName: string,
  fileExtension = "docx",
): Promise<string> {
  const file = await exportEditorDocumentAsFile(manager, fileName, fileExtension);
  const buffer = await file.arrayBuffer();
  return bufferToBase64(buffer);
}
