import { FILE_TYPE, type FileType } from "../const";

const OFFICE_MIME_TYPES: Record<FileType, string> = {
  [FILE_TYPE.DOCX]:
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  [FILE_TYPE.XLSX]:
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  [FILE_TYPE.PPTX]:
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
};

export function getOfficeMimeType(fileType: FileType): string {
  return OFFICE_MIME_TYPES[fileType];
}

/** 触发浏览器下载 Blob，不刷新页面 */
export function downloadBlob(blob: Blob, fileName: string) {
  const url = window.URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.style.display = "none";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  window.URL.revokeObjectURL(url);
}
