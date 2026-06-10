export type RevisionData = {
  Type?: number;
  TypeName?: string;
  UserId?: string;
  UserName?: string;
  DateTime?: string;
  Value?: string;
  X?: number;
  Y?: number;
  MoveType?: number;
  MoveId?: string;
  Locked?: boolean;
  LockUserId?: string;
};

export type RevisionItem = {
  /** `rev-${index}` 或与 SDK 元素 id 一致 */
  Id: string;
  Index: number;
  Data: RevisionData;
  /** SDK 原始修订对象，供 accept / reject / goTo 使用 */
  Raw: SdkRevisionChange;
};

export type RevisionChangeHandlers = {
  onShowChanges?: (items: RevisionItem[]) => void;
  onTrackRevisionsChange?: (enabled: boolean) => void;
};

/** v7+ SDK 修订对象（AscCommon 内部类） */
export type SdkRevisionChange = {
  get_Type?: () => number;
  get_UserId?: () => string;
  get_UserName?: () => string;
  get_DateTime?: () => string;
  get_Value?: () => string;
  get_X?: () => number;
  get_Y?: () => number;
  get_MoveType?: () => number;
  get_MoveId?: () => string;
  get_StartPos?: () => number;
  get_EndPos?: () => number;
  get_LockUserId?: () => string | null;
  sca?: string;
  tE?: number;
};

/** 文档内修订元素（Yq）或弹层对象（tv） */
type SdkRevisionElement = SdkRevisionChange & {
  Element?: unknown;
  /** 9.3+ 选区起止 */
  wa?: number | null;
  xa?: number | null;
  /** 9.3+ 移动修订分组 */
  OL?: SdkRevisionElement[];
  vva?: string;
  nW?: number;
  sa?: number;
  ra?: number;
  yD?: unknown[];
  Gc?: () => number;
  Qy?: () => string;
  wV?: () => string;
  nX?: () => string;
  Ym?: () => unknown;
  ed?: () => string;
  sh?: () => unknown;
  ee?: () => unknown;
  SU?: () => number;
  h5?: () => number;
  j9?: () => number;
  yza?: () => number;
  hZd?: () => number;
  $Zd?: () => number;
  Vtb?: () => number;
};

/** 9.3+ 修订管理器（doc.Wq） */
type RevisionWq = {
  z9a?: () => void;
  TLh?: (change: SdkRevisionElement) => void;
  Znf?: () => void;
  u0d?: (elements: unknown[]) => boolean;
  /** 重置 j5d 导航游标 */
  $_d?: () => void;
  /** 等待修订索引后再显示气泡 */
  t0d?: () => void;
  Ydf?: (keepPosition?: boolean) => void;
  /** j5d 导航后的当前修订 */
  kz?: SdkRevisionElement | null;
  wih?: () => Record<string, SdkRevisionElement[]>;
  Tpa?: Array<{ Yb?: () => string; ed?: () => string }>;
};

type RevisionUm = {
  T8?: Array<{ ed?: () => string }>;
  MQc?: () => void;
  Dmf?: () => void;
  yif?: () => Record<string, SdkRevisionElement[]>;
};

type WordLogicDocument = LogicDocumentWithRevisions & {
  lc?: () => void;
  ec?: () => void;
  Pf?: (show?: boolean) => void;
  yf?: (update?: boolean) => void;
  ze?: (mode?: boolean, force?: boolean) => void;
  Ek?: { f8?: (wq: RevisionWq) => void };
  qde?: (id: null, redraw: boolean) => void;
  pb?: { p8a?: () => void };
  zj?: { W5b?: boolean; Ptb?: unknown };
  Ihc?: (
    moveId: string,
    fromMove: boolean,
    scroll?: boolean,
    select?: boolean,
  ) => void;
  Um?: RevisionUm;
  Wq?: RevisionWq;
};

export type RevisionsEditorApi = {
  /** v9+ 逻辑文档（asc_docs_api.te） */
  te?: () => WordLogicDocument | null;
  Gg?: () => LogicDocumentWithRevisions | null;
  Ge?: () => LogicDocumentWithRevisions | null;
  asc_BeginViewModeInReview?: (finalMode?: boolean) => void;
  asc_SetDisplayModeInReview?: (mode: unknown) => void;
  pluginMethod_SetDisplayModeInReview?: (mode: string) => void;
  pluginMethod_MoveToNextReviewChange?: (next: boolean) => void;
  sync_BeginCatchRevisionsChanges?: () => void;
  sync_EndCatchRevisionsChanges?: (show?: boolean) => void;
  asc_GetRevisionsChangesStack?: () => SdkRevisionChange[];
  asc_GetTrackRevisionsReportByAuthors?: () => Record<
    string,
    SdkRevisionElement[]
  >;
  asc_GetNextRevisionsChange?: () => unknown;
  asc_GetPrevRevisionsChange?: () => unknown;
  asc_AcceptChanges?: (change?: SdkRevisionChange) => void;
  asc_RejectChanges?: (change?: SdkRevisionChange) => void;
  asc_HaveRevisionsChanges?: (all?: boolean) => boolean;
  asc_FollowRevisionMove?: (change: SdkRevisionChange) => void;
  asc_SetGlobalTrackRevisions?: (enabled: boolean) => void;
  asc_GetGlobalTrackRevisions?: () => boolean;
  asc_SetLocalTrackRevisions?: (enabled: boolean) => void;
  asc_GetLocalTrackRevisions?: () => boolean;
  ViewScrollToX?: (x: number) => void;
  ViewScrollToY?: (y: number) => void;
};

type LogicDocumentWithRevisions = {
  Um?: {
    MQc?: () => void;
    Dmf?: () => void;
    NVd?: () => void;
  };
};

type AscRevisionEnums = {
  TextAdd?: number;
  TextRem?: number;
  ParaAdd?: number;
  ParaRem?: number;
  TextPr?: number;
  ParaPr?: number;
  TablePr?: number;
  RowsAdd?: number;
  RowsRem?: number;
  TableRowPr?: number;
  MoveMark?: number;
  Unknown?: number;
  MoveTo?: number;
  MoveFrom?: number;
  NoMove?: number;
};

function getAscRevisionEnums(frameWin: Window): AscRevisionEnums {
  const Asc = (frameWin as Window & { Asc?: Record<string, unknown> })?.Asc;
  const fu = Asc?.c_oAscRevisionsChangeType as Record<string, number> | undefined;
  const Bs = Asc?.c_oAscRevisionsMove as Record<string, number> | undefined;
  return {
    TextAdd: fu?.TextAdd,
    TextRem: fu?.TextRem,
    ParaAdd: fu?.ParaAdd,
    ParaRem: fu?.ParaRem,
    TextPr: fu?.TextPr,
    ParaPr: fu?.ParaPr,
    TablePr: fu?.TablePr,
    RowsAdd: fu?.RowsAdd,
    RowsRem: fu?.RowsRem,
    TableRowPr: fu?.TableRowPr,
    MoveMark: fu?.MoveMark,
    Unknown: fu?.Unknown,
    MoveTo: Bs?.MoveTo,
    MoveFrom: Bs?.MoveFrom,
    NoMove: Bs?.NoMove,
  };
}

function revisionTypeName(type: number, enums: AscRevisionEnums): string {
  const entries: [string, number | undefined][] = [
    ["TextAdd", enums.TextAdd],
    ["TextRem", enums.TextRem],
    ["ParaAdd", enums.ParaAdd],
    ["ParaRem", enums.ParaRem],
    ["TextPr", enums.TextPr],
    ["ParaPr", enums.ParaPr],
    ["TablePr", enums.TablePr],
    ["RowsAdd", enums.RowsAdd],
    ["RowsRem", enums.RowsRem],
    ["TableRowPr", enums.TableRowPr],
    ["MoveMark", enums.MoveMark],
    ["Unknown", enums.Unknown],
  ];
  for (const [name, val] of entries) {
    if (val !== undefined && val === type) return name;
  }
  return String(type);
}

function getLogicDocument(
  api: RevisionsEditorApi,
): WordLogicDocument | null {
  if (typeof api.te === "function") return api.te() ?? null;
  if (typeof api.Ge === "function") return (api.Ge() as WordLogicDocument) ?? null;
  if (typeof api.Gg === "function") return (api.Gg() as WordLogicDocument) ?? null;
  return null;
}

function revisionStart(raw: SdkRevisionElement): number | undefined {
  if (typeof raw.get_StartPos === "function") {
    const pos = raw.get_StartPos();
    if (typeof pos === "number") return pos;
  }
  if (typeof raw.SU === "function") {
    const pos = raw.SU();
    if (typeof pos === "number") return pos;
  }
  if (typeof raw.wa === "number") return raw.wa;
  if (typeof raw.j9 === "function") return raw.j9();
  if (typeof raw.sa === "number") return raw.sa;
  return undefined;
}

function revisionEnd(raw: SdkRevisionElement): number | undefined {
  if (typeof raw.get_EndPos === "function") {
    const pos = raw.get_EndPos();
    if (typeof pos === "number") return pos;
  }
  if (typeof raw.h5 === "function") {
    const pos = raw.h5();
    if (typeof pos === "number") return pos;
  }
  if (typeof raw.xa === "number") return raw.xa;
  if (typeof raw.yza === "function") return raw.yza();
  if (typeof raw.ra === "number") return raw.ra;
  return undefined;
}

function isRevisionGroupedMove(raw: SdkRevisionElement): boolean {
  return Array.isArray(raw.OL) && raw.OL.length > 0;
}

function revisionIdentity(raw: SdkRevisionElement): string | null {
  if (typeof raw.ed === "function") return String(raw.ed());
  return null;
}

function isSameRevision(
  a: SdkRevisionElement,
  b: SdkRevisionElement,
): boolean {
  if (a === b) return true;
  const idA = revisionIdentity(a);
  const idB = revisionIdentity(b);
  return idA != null && idA === idB;
}

/** z9a 后从 Wq 索引取带 Element/wa/xa 的修订对象（缓存 Raw 可能不可用于选区） */
function resolveFreshRevisionRaw(
  doc: WordLogicDocument,
  raw: SdkRevisionElement,
): SdkRevisionElement {
  const id = revisionIdentity(raw);
  if (!id) return raw;

  const qt = doc.Wq?.wih?.();
  if (!qt) return raw;

  for (const group of Object.values(qt)) {
    if (!Array.isArray(group)) continue;
    for (const entry of group) {
      if (revisionIdentity(entry) === id) return entry;
    }
  }

  return raw;
}

/** 对齐 SDK 9.3 zL：使用 Element + wa/xa 选区 */
function applyRevisionSelection(
  doc: WordLogicDocument,
  raw: SdkRevisionElement,
  frameWin: Window,
): boolean {
  if (isRevisionGroupedMove(raw) && raw.vva) {
    const enums = getAscRevisionEnums(frameWin);
    doc.lc?.();
    doc.Ihc?.(raw.vva, raw.nW === enums.MoveFrom, false, false);
    return true;
  }

  const element =
    raw.Element ?? (typeof raw.ee === "function" ? raw.ee() : undefined);
  const start =
    typeof raw.wa === "number" ? raw.wa : revisionStart(raw);
  if (element == null || start == null) return false;

  const end =
    typeof raw.xa === "number" ? raw.xa : (revisionEnd(raw) ?? start);

  try {
    doc.lc?.();

    const el = element as {
      Dp?: (start: number, a: boolean, b: number, c: number) => void;
      Gs?: (start: number, end: number, flag: boolean) => void;
      Ft?: (flag: boolean) => void;
      qp?: (start: number, a: boolean, b: number, c: number) => void;
      mo?: (start: number, end: number, flag: boolean) => void;
      hq?: (flag: boolean) => void;
      Selection?: { Na?: boolean; La?: boolean };
    };

    if (typeof el.Dp === "function" && typeof el.Gs === "function") {
      if (doc.Wq?.u0d?.([element])) return false;
      el.Dp(start, false, -1, -1);
      if (el.Selection) el.Selection.Na = true;
      el.Gs(start, end, false);
      el.Ft?.(false);
      return true;
    }

    if (typeof el.qp === "function" && typeof el.mo === "function") {
      doc.ec?.();
      el.qp(start, false, -1, -1);
      if (el.Selection) el.Selection.La = true;
      el.mo(start, end, false);
      el.hq?.(false);
      doc.Pf?.(false);
      doc.yf?.(true);
      return true;
    }

    return false;
  } catch {
    return false;
  }
}

/** 关闭 OnlyOffice 内置 Review Changes 浮窗（showReviewChanges 为 true 时会出现） */
function dismissReviewChangesDialog(
  frameWin?: Window | null,
): void {
  const doc = frameWin?.document;
  if (!doc) return;

  doc.querySelectorAll(".asc-window.review-changes.modal-dlg").forEach((el) => {
    const close = el.querySelector(".close, .btn-close");
    if (close instanceof HTMLElement) {
      close.click();
      return;
    }
    (el as HTMLElement).style.display = "none";
  });
}

/**
 * 进入修订审阅显示模式（页面初始化时调用一次即可）。
 * 禁止 asc_BeginViewModeInReview：SDK $hc 切到 original/final 会 kAc/NKc 批量处理全部修订。
 */
export function prepareRevisionReviewDisplay(
  api: RevisionsEditorApi,
  frameWin?: Window | null,
): void {
  api.asc_SetLocalTrackRevisions?.(true);
  // simple(ewa) 模式下 SDK 不渲染修订气泡；须用 markup(AKa)
  if (typeof api.pluginMethod_SetDisplayModeInReview === "function") {
    api.pluginMethod_SetDisplayModeInReview("markup");
    dismissReviewChangesDialog(frameWin);
    return;
  }
  if (typeof api.asc_SetDisplayModeInReview === "function") {
    const Asc = (frameWin as Window & {
      Asc?: { c_oAscReviewDisplay?: { Edit?: unknown; Markup?: unknown }; Xja?: { AKa?: unknown } };
    })?.Asc;
    const markup =
      Asc?.c_oAscReviewDisplay?.Edit ??
      Asc?.c_oAscReviewDisplay?.Markup ??
      Asc?.Xja?.AKa;
    if (markup !== undefined) {
      api.asc_SetDisplayModeInReview(markup);
    }
  }

  dismissReviewChangesDialog(frameWin);
}

function revisionType(raw: SdkRevisionElement): number | undefined {
  if (typeof raw.get_Type === "function") return raw.get_Type();
  if (typeof raw.Gc === "function") return raw.Gc();
  return undefined;
}

function revisionValue(raw: SdkRevisionElement): string | undefined {
  if (typeof raw.get_Value === "function") {
    const v = raw.get_Value();
    if (typeof v === "string") return v;
    if (v != null) return String(v);
  }
  if (typeof raw.Ym === "function") {
    const v = raw.Ym();
    if (typeof v === "string") return v;
    if (v != null) return String(v);
  }
  return undefined;
}

function revisionChangeToPlain(
  raw: SdkRevisionElement,
  enums: AscRevisionEnums,
): RevisionData {
  const type = revisionType(raw);
  const lockUserId = raw.get_LockUserId?.();
  return {
    Type: type,
    TypeName: type !== undefined ? revisionTypeName(type, enums) : undefined,
    UserId:
      (typeof raw.get_UserId === "function" ? raw.get_UserId() : undefined) ??
      (typeof raw.Qy === "function" ? raw.Qy() : undefined),
    UserName:
      (typeof raw.get_UserName === "function"
        ? raw.get_UserName()
        : undefined) ??
      (typeof raw.wV === "function" ? raw.wV() : undefined),
    DateTime:
      (typeof raw.get_DateTime === "function"
        ? raw.get_DateTime()
        : undefined) ??
      (typeof raw.nX === "function" ? raw.nX() : undefined),
    Value: revisionValue(raw),
    X:
      (typeof raw.get_X === "function" ? raw.get_X() : undefined) ??
      (typeof raw.$Zd === "function" ? raw.$Zd() : undefined),
    Y:
      (typeof raw.get_Y === "function" ? raw.get_Y() : undefined) ??
      (typeof raw.Vtb === "function" ? raw.Vtb() : undefined),
    MoveType:
      (typeof raw.get_MoveType === "function" ? raw.get_MoveType() : undefined) ??
      (typeof raw.hZd === "function" ? raw.hZd() : undefined),
    MoveId: raw.get_MoveId?.(),
    Locked: lockUserId != null && lockUserId !== "",
    LockUserId: lockUserId ?? undefined,
  };
}

let syncingRevisionStack = false;

function ensureRevisionsIndexed(api: RevisionsEditorApi): void {
  const doc = getLogicDocument(api);
  doc?.Wq?.z9a?.();
  doc?.Um?.MQc?.();
}

function collectRevisionRawsFromWq(
  doc: WordLogicDocument,
): SdkRevisionElement[] {
  const wq = doc.Wq;
  const qt = wq?.wih?.();
  if (!qt || typeof qt !== "object" || Object.keys(qt).length === 0) {
    return [];
  }

  const raws: SdkRevisionElement[] = [];
  const orderedIds: string[] = [];

  for (const el of wq?.Tpa ?? []) {
    const id =
      typeof el.Yb === "function"
        ? String(el.Yb())
        : typeof el.ed === "function"
          ? String(el.ed())
          : "";
    if (id && qt[id] && !orderedIds.includes(id)) {
      orderedIds.push(id);
    }
  }

  const rest = Object.keys(qt)
    .filter((id) => !orderedIds.includes(id))
    .sort((a, b) => {
      const ga = qt[a]?.[0];
      const gb = qt[b]?.[0];
      if (!ga || !gb) return 0;
      return revisionSortKey(ga) - revisionSortKey(gb);
    });
  orderedIds.push(...rest);

  for (const id of orderedIds.length > 0 ? orderedIds : Object.keys(qt)) {
    const group = qt[id];
    if (!Array.isArray(group)) continue;
    raws.push(
      ...(group.length > 1
        ? [...group].sort((a, b) => revisionSortKey(a) - revisionSortKey(b))
        : group),
    );
  }

  return raws;
}

function readRevisionStack(
  api: RevisionsEditorApi,
  allowSyncEnd: boolean,
): SdkRevisionElement[] {
  const existing = (api.asc_GetRevisionsChangesStack?.() ??
    []) as SdkRevisionElement[];
  if (existing.length > 0 || !allowSyncEnd || syncingRevisionStack) {
    return existing;
  }

  syncingRevisionStack = true;
  try {
    api.sync_BeginCatchRevisionsChanges?.();
    getLogicDocument(api)?.Um?.Dmf?.();
    api.sync_EndCatchRevisionsChanges?.(false);
    return (api.asc_GetRevisionsChangesStack?.() ?? []) as SdkRevisionElement[];
  } finally {
    syncingRevisionStack = false;
  }
}

function revisionSortKey(raw: SdkRevisionElement): number {
  const y = typeof raw.$Zd === "function" ? raw.$Zd() : 0;
  const x = typeof raw.Vtb === "function" ? raw.Vtb() : 0;
  return y * 1e9 + x * 1e3 + (revisionStart(raw) ?? 0);
}

function mapRevisionItems(
  raws: SdkRevisionElement[],
  enums: AscRevisionEnums,
  idOf: (raw: SdkRevisionElement, index: number) => string,
): RevisionItem[] {
  return raws.map((raw, index) => ({
    Id: idOf(raw, index),
    Index: index,
    Data: revisionChangeToPlain(raw, enums),
    Raw: raw,
  }));
}

export type CollectRevisionItemsOptions = {
  /** 为 false 时不调用 sync_EndCatchRevisionsChanges，避免 asc_onShowRevisionsChange 递归 */
  allowSyncEnd?: boolean;
};

/** v9 Wq → v7 Yq+T8 → 报告 → 修订栈 */
export function collectRevisionItems(
  api: RevisionsEditorApi,
  frameWin: Window,
  options: CollectRevisionItemsOptions = {},
): RevisionItem[] {
  const { allowSyncEnd = true } = options;
  ensureRevisionsIndexed(api);
  const enums = getAscRevisionEnums(frameWin);
  const revId = (raw: SdkRevisionElement, i: number) =>
    `rev-${typeof raw.ed === "function" ? String(raw.ed()) : i}`;

  const doc = getLogicDocument(api);
  const wqRaws = doc ? collectRevisionRawsFromWq(doc) : [];
  if (wqRaws.length > 0) {
    return mapRevisionItems(wqRaws, enums, revId);
  }

  const um = doc?.Um;
  const yq = um?.yif?.();
  if (yq && Object.keys(yq).length > 0) {
    const ids: string[] = [];
    for (const el of um?.T8 ?? []) {
      const id = typeof el.ed === "function" ? String(el.ed()) : "";
      if (id && yq[id]) ids.push(id);
    }
    const rest = Object.keys(yq).filter((id) => !ids.includes(id));
    rest.sort((a, b) => {
      const ga = yq[a]?.[0];
      const gb = yq[b]?.[0];
      if (!ga || !gb) return 0;
      return revisionSortKey(ga) - revisionSortKey(gb);
    });
    ids.push(...rest);

    const raws: SdkRevisionElement[] = [];
    for (const id of ids.length > 0 ? ids : Object.keys(yq)) {
      const group = yq[id];
      if (!Array.isArray(group)) continue;
      raws.push(
        ...(group.length > 1
          ? [...group].sort((a, b) => revisionSortKey(a) - revisionSortKey(b))
          : group),
      );
    }
    if (raws.length > 0) return mapRevisionItems(raws, enums, revId);
  }

  const report = api.asc_GetTrackRevisionsReportByAuthors?.();
  if (report && typeof report === "object") {
    const raws: SdkRevisionElement[] = [];
    for (const key of Object.keys(report)) {
      const group = report[key];
      if (Array.isArray(group)) raws.push(...group);
    }
    if (raws.length > 0) return mapRevisionItems(raws, enums, revId);
  }

  const stack = readRevisionStack(api, allowSyncEnd);
  return mapRevisionItems(stack, enums, (_, i) => `rev-stack-${i}`);
}

/** asc_onShowRevisionsChange 回调：栈有值时归一化，否则全量收集 */
export function resolveRevisionShowChanges(
  stack: unknown,
  api: RevisionsEditorApi,
  frameWin: Window,
): RevisionItem[] {
  if (Array.isArray(stack) && stack.length > 0) {
    const enums = getAscRevisionEnums(frameWin);
    return (stack as SdkRevisionElement[]).map((raw, index) => ({
      Id: `rev-stack-${index}`,
      Index: index,
      Data: revisionChangeToPlain(raw, enums),
      Raw: raw,
    }));
  }
  return collectRevisionItems(api, frameWin, { allowSyncEnd: false });
}

function findRevisionItem(
  items: RevisionItem[],
  id: string,
): RevisionItem | undefined {
  return items.find((item) => item.Id === id);
}

function resolveRevisionTarget(
  target: string | RevisionItem,
  api: RevisionsEditorApi,
  frameWin: Window,
  cachedItems?: RevisionItem[],
): RevisionItem | undefined {
  if (typeof target !== "string") {
    return target.Raw ? target : undefined;
  }

  const cached = cachedItems?.find((item) => item.Id === target);
  if (cached) return cached;

  return findRevisionItem(
    collectRevisionItems(api, frameWin, { allowSyncEnd: false }),
    target,
  );
}

function scrollToRevisionPosition(
  api: RevisionsEditorApi,
  data: RevisionData,
): void {
  if (data.Y != null) api.ViewScrollToY?.(data.Y);
  if (data.X != null) api.ViewScrollToX?.(data.X);
}

function scrollRevisionIntoView(
  doc: WordLogicDocument,
  api: RevisionsEditorApi,
  data: RevisionData,
): void {
  if (doc.zj?.W5b && doc.zj?.Ptb && typeof doc.qde === "function") {
    doc.qde(null, false);
    doc.pb?.p8a?.();
    return;
  }
  scrollToRevisionPosition(api, data);
}

/** Ydf → Ek.f8 → Znf → ze（TLh 在选区前已调用，激活时再设当前修订） */
function activateRevision(
  doc: WordLogicDocument,
  raw: SdkRevisionElement,
): void {
  const wq = doc.Wq;
  if (!wq) return;

  wq.t0d?.();
  wq.TLh?.(raw);
  wq.Ydf?.(true);
  doc.Ek?.f8?.(wq);
  wq.Znf?.();
  doc.ze?.(true);
}

/** j5d 回退：按 Index 逐步 GetNext，内部会 zL 激活 */
function navigateToRevisionByStep(
  api: RevisionsEditorApi,
  doc: WordLogicDocument,
  targetRaw: SdkRevisionElement,
  targetIndex: number,
): boolean {
  if (typeof api.asc_GetNextRevisionsChange !== "function") return false;

  doc.Wq?.$_d?.();
  const steps = Math.max(targetIndex + 1, 1);
  for (let i = 0; i < steps; i++) {
    api.asc_GetNextRevisionsChange();
    const current = doc.Wq?.kz;
    if (current && isSameRevision(current, targetRaw)) return true;
  }

  const current = doc.Wq?.kz;
  return current != null && isSameRevision(current, targetRaw);
}

/** 对齐 SDK j5d：z9a → TLh → zL → wK → 激活（单次 ze，避免相邻修订闪烁） */
function focusRevisionInDocument(
  api: RevisionsEditorApi,
  doc: WordLogicDocument,
  raw: SdkRevisionElement,
  data: RevisionData,
  frameWin: Window,
  index: number,
): boolean {
  ensureRevisionsIndexed(api);

  const resolvedRaw = resolveFreshRevisionRaw(doc, raw);
  doc.Wq?.TLh?.(resolvedRaw);

  if (applyRevisionSelection(doc, resolvedRaw, frameWin)) {
    scrollRevisionIntoView(doc, api, data);
    activateRevision(doc, resolvedRaw);
    return true;
  }

  if (navigateToRevisionByStep(api, doc, resolvedRaw, index)) {
    scrollRevisionIntoView(doc, api, data);
    return true;
  }

  scrollToRevisionPosition(api, data);
  activateRevision(doc, resolvedRaw);
  return true;
}

function isMoveRevision(
  raw: SdkRevisionElement,
  frameWin: Window,
): boolean {
  const enums = getAscRevisionEnums(frameWin);
  const moveType =
    (typeof raw.get_MoveType === "function" ? raw.get_MoveType() : undefined) ??
    (typeof raw.hZd === "function" ? raw.hZd() : undefined);
  if (moveType === undefined) return false;
  return moveType === enums.MoveTo || moveType === enums.MoveFrom;
}

/**
 * 定位到指定修订（仅导航+高亮，不接受/拒绝）。
 */
export function goToRevision(
  target: string | RevisionItem,
  api: RevisionsEditorApi,
  frameWin: Window,
  cachedItems?: RevisionItem[],
): boolean {
  const item = resolveRevisionTarget(target, api, frameWin, cachedItems);
  if (!item?.Raw) return false;

  const doc = getLogicDocument(api);
  if (!doc) return false;

  const { Raw: raw, Data: data, Index: index } = item;

  if (isMoveRevision(raw, frameWin)) {
    api.asc_FollowRevisionMove?.(raw);
    scrollToRevisionPosition(api, data);
  }

  return focusRevisionInDocument(api, doc, raw, data, frameWin, index);
}
