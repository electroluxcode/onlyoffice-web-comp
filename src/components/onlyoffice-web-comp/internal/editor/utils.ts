export { getDocumentType, getNewUrl } from "../../const";

export function getFileExt(name: string) {
  const type = name.split(".").pop() || "";
  return type.toLowerCase();
}
