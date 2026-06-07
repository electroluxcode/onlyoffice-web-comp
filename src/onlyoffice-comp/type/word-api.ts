export type AscWordApiMethod =
  | "asc_onAddComment"
  | "asc_onChangeCommentData"
  | "asc_onRemoveComment"
  | "asc_onShowComment"
  | "asc_onHideComment"
  /** 工具栏当前字体族（对应编辑器内部 N9e → Zb） */
  | "asc_onFontFamily"
  | "asc_onShowRevisionsChange"
  | "asc_onOnTrackRevisionsChange"
  | "asc_onDocumentModifiedChanged"
  | "asc_onSaveCallback"
  | "asc_onError"
  | "asc_onInfo"
  | "asc_onWarning"
  | (string & {});

export type AscWordApiCallback = (...args: unknown[]) => void;
