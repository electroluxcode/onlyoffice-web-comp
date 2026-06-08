import { converter } from "../internal/editor/x2t";
import { getX2tExportFormats } from "../internal/editor/utils";
import {
  type CreateEditorViewOptions,
  editorManagerFactory,
} from "../core/editor-manager";
import type { FileType } from "../const";

export async function createEditorView(options: CreateEditorViewOptions) {
  const manager =
    options.editorManager ||
    (options.containerId
      ? editorManagerFactory.get(options.containerId)
      : editorManagerFactory.getDefault());

  return manager.create(options);
}

export async function convertBinToDocument(
  binData: Uint8Array,
  fileName: string,
  fileType: FileType,
  media?: Record<string, Uint8Array>,
) {
  const targetExt = fileType.toLowerCase();
  const data = new Uint8Array(binData).buffer;
  const { formatFrom, formatTo } = getX2tExportFormats(targetExt);
  const result = await converter.convert({
    data,
    fileFrom: "Editor.bin",
    fileTo: `doc.${targetExt}`,
    formatFrom,
    formatTo,
    media,
  });

  if (!result.output) {
    throw new Error("Failed to convert OnlyOffice bin document");
  }

  return {
    fileName,
    data: result.output,
  };
}
