import {
  installOnlyOfficeProxies,
  installReporterWindowHook,
  type OnlyOfficeProxyWindow,
  type ReporterHookWindow,
} from "../internal/editor/install-proxies";
import { EditorServer } from "../internal/editor/server";
import io, { MockSocket, type MockSocketOptions } from "../internal/editor/socket";
import {
  type DocEditor,
  DocumentType,
  type OfficeTheme,
  type PluginMode,
  type User,
} from "../internal/editor/types";
import { getDocumentType } from "../const";
import type { AscWordApiCallback, AscWordApiMethod } from "../type/word-api";
import {
  type OnlyOfficeIframeWindow,
} from "../type/sdk-internal";
import {
  ONLYOFFICE_CONTAINER_CONFIG,
  ONLYOFFICE_EVENT_KEYS,
  ONLYOFFICE_ID,
  ASC_RESTRICTION_NONE,
  ASC_RESTRICTION_VIEW,
  OFFICE_EDITOR_LOGO,
} from "../const";
import {
  type CommentChangeHandlers,
  type CommentData,
  type CommentInput,
  type CommentItem,
  isResolvedComment,
  normalizeCommentInput,
  toPluginCommentPayload,
} from "../feature/comments";
import { onlyofficeEventbus } from "./eventbus";
import {
  type RevisionChangeHandlers,
  type RevisionItem,
  normalizeRevisionItems,
  flattenRevisionsReportByAuthors,
} from "../feature/revisions";
import { getOnlyOfficeLang, type OnlyOfficeLang } from "../store/lang";
import { initializeOnlyOffice } from "../util/initialize";

export type CreateEditorViewOptions = {
  isNew: boolean;
  fileName: string;
  file?: File;
  url?: string;
  loader?: (url: string) => Promise<ArrayBuffer>;
  fileType?: string;
  readOnly?: boolean;
  user?: User;
  lang?: string;
  containerId?: string;
  editorManager?: EditorManager;
  editing?: boolean;
  theme?: OfficeTheme;
  plugins?: PluginMode;
  /** 由 EditorManagerFactory.beginLoadSession 生成，用于丢弃过期的异步初始化 */
  loadSession?: number;
};

let instanceIndex = 0;

function getFileType(fileName: string, fileType?: string) {
  return fileType || fileName.split(".").pop()?.toLowerCase() || "docx";
}

function getNewInstanceId() {
  instanceIndex += 1;
  return `onlyoffice-${instanceIndex}`;
}

type OnlyOfficeSdkApi = {
  i1f?: (priority?: number) => void;
  asyncFontsDocumentEndLoaded?: (priority?: number) => void;
  ra?: { Ghj?: () => void };
  asc_registerCallback?: (type: string, fn: AscWordApiCallback) => void;
  asc_unregisterCallback?: (type: string, fn: AscWordApiCallback) => void;
  asc_addComment?: (data: CommentData) => string | undefined;
  asc_changeComment?: (id: string, data: CommentData) => void;
  asc_removeComment?: (id: string) => void;
  sync_ChangeCommentData?: (id: string, data: unknown, ...args: unknown[]) => unknown;
  __ONLYOFFICE_RESOLVE_PATCHED__?: boolean;
  asc_selectComment?: (id: string) => void;
  asc_showComment?: (id: string) => void;
  asc_showComments?: () => void;
  asc_hideComments?: () => void;
  asc_SetGlobalTrackRevisions?: (enabled: boolean) => void;
  asc_GetGlobalTrackRevisions?: () => boolean;
  asc_GetRevisionsChangesStack?: () => unknown[];
  asc_HaveRevisionsChanges?: () => boolean;
  asc_GetTrackRevisionsReportByAuthors?: () => Record<string, unknown[]>;
  asc_BeginViewModeInReview?: (finalMode?: boolean) => void;
  asc_EndViewModeInReview?: () => void;
  asc_AcceptChanges?: (change?: unknown) => void;
  asc_RejectChanges?: (change?: unknown) => void;
  asc_AcceptChangesBySelection?: (all?: boolean) => void;
  asc_RejectChangesBySelection?: (all?: boolean) => void;
  pluginMethod_GetAllComments?: () => Array<{ Id: string; Data: CommentData }>;
  pluginMethod_AddComment?: (data: CommentData) => string | null;
  pluginMethod_ChangeComment?: (id: string, data: CommentData) => void;
};

/** iframe 内运行时；混淆字段见 type/sdk-internal.ts，asc_* 为公开 API */
type OnlyOfficeWindow = OnlyOfficeIframeWindow & {
  Asc?: Omit<NonNullable<OnlyOfficeIframeWindow["Asc"]>, "editor"> & {
    editor?: OnlyOfficeSdkApi &
      {
        asc_Recalculate?: () => void;
      };
  };
};

type ScopedIoFactory = (url?: string, options?: MockSocketOptions) => MockSocket;

type ShellMainController = {
  mode?: { isEdit?: boolean; canEdit?: boolean };
};

type WordHeaderView = {
  btnDocMode?: { setVisible?: (visible: boolean) => void };
  btnPDFMode?: { setVisible?: (visible: boolean) => void };
};

type OnlyOfficeParentWindow = Window & {
  __ONLYOFFICE_SCOPED_IO__?: Record<string, ScopedIoFactory>;
};

export class EditorManager {
  private editor: DocEditor | null = null;
  private server: EditorServer;
  private dirty = false;
  private readOnly = false;
  private editorLang: OnlyOfficeLang = getOnlyOfficeLang();
  private uiTheme: OfficeTheme = "theme-white";
  private instanceId = getNewInstanceId();
  private containerId: string;
  private plugins: PluginMode = "featured";
  private fileName = "New Document.docx";
  private fileType = "docx";
  private media: Record<string, Uint8Array> = {};
  private comments = new Map<string, CommentData>();
  private revisions: RevisionItem[] = [];
  private wordContentSyncPromise: Promise<void> | null = null;
  private wordContentSyncTeardown: (() => void) | null = null;

  constructor(containerId = ONLYOFFICE_ID) {
    this.containerId = containerId;
    this.server = new EditorServer({
      getState: () => ({ plugins: "none", readOnly: this.readOnly }),
      onUserSave: (snapshot) => {
        this.dirty = false;
        this.notifyUserSave(snapshot);
      },
    });
  }


  private createScopedIo() {
    return (url?: string, options: MockSocketOptions = {}) => {
      const socket = io(url, options);

      socket.on("connect", () => {
        this.server.handleConnect({ socket });
      });
      socket.on("disconnect", () => {
        this.server.handleDisconnect({ socket });
      });

      return socket;
    };
  }

  private getEditorFrameElement() {
    const containerFrame = document
      .getElementById(this.containerId)
      ?.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');

    if (containerFrame) {
      return containerFrame;
    }

    const frames = Array.from(
      document.querySelectorAll<HTMLIFrameElement>('iframe[name="frameEditor"]'),
    );
    const matchedFrame = frames.find((frame) => {
      try {
        const url = new URL(frame.src, window.location.origin);
        return url.searchParams.get("frameEditorId") === this.containerId;
      } catch {
        return false;
      }
    });

    if (matchedFrame) {
      return matchedFrame;
    }

    if (this.containerId === ONLYOFFICE_ID) {
      return frames[0];
    }

    return document
      .querySelector<HTMLElement>(
        `${ONLYOFFICE_CONTAINER_CONFIG.PARENT_SELECTOR}[data-onlyoffice-container-id="${this.containerId}"]`,
      )
      ?.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
  }

  private installProxiesOnWindow(win: OnlyOfficeProxyWindow) {
    installOnlyOfficeProxies(win, this.server, this.createScopedIo());
  }

  /**
   * 劫持 iframe 内 XHR/fetch/io，将协作与 downloadAs 请求路由到 mock EditorServer。
   * 必须在 downloadAs 前安装，否则 export 无法收到 /downloadas/ 分片。
   */
  private installIframeProxies() {
    const iframe = this.getEditorFrameElement();
    const win = iframe?.contentWindow as
      | (OnlyOfficeWindow & ReporterHookWindow)
      | undefined;
    const iframeDoc = iframe?.contentDocument;

    if (!iframeDoc || !win) {
      throw new Error("Iframe not loaded");
    }

    if (win.__ONLYOFFICE_PROXIES_INSTALLED__) {
      return;
    }

    this.installProxiesOnWindow(win);
    installReporterWindowHook(win, (target) => {
      this.installProxiesOnWindow(target as OnlyOfficeProxyWindow);
    });
    this.installSaveShortcutBlocker();
  }

  private getEditorFrameWindow() {
    const iframe = this.getEditorFrameElement();

    return iframe?.contentWindow as OnlyOfficeWindow | undefined;
  }

  private getSdkApi() {
    return this.getEditorFrameWindow()?.Asc?.editor;
  }

  private requireSdkApi() {
    const api = this.getSdkApi();

    if (!api) {
      throw new Error("OnlyOffice SDK API is not ready");
    }

    return api;
  }

  private installCommentResolveCleanup() {
    const api = this.getSdkApi();

    if (!api || api.__ONLYOFFICE_RESOLVE_PATCHED__) {
      return;
    }

    api.__ONLYOFFICE_RESOLVE_PATCHED__ = true;

    const removeResolvedComment = (id: unknown) => {
      // OnlyOffice resolves comments through an internal change event first.
      // Removing synchronously during that event can race its own render pass,
      // so schedule the delete for the next tick after the resolved state lands.
      window.setTimeout(() => {
        api.asc_removeComment?.(String(id));
        this.comments.delete(String(id));
      }, 0);
    };

    const originalSyncChange = api.sync_ChangeCommentData?.bind(api);
    if (originalSyncChange) {
      api.sync_ChangeCommentData = (id, data, ...args) => {
        const result = originalSyncChange(id, data, ...args);

        if (isResolvedComment(data)) {
          removeResolvedComment(id);
        }

        return result;
      };
    }

    api.asc_registerCallback?.("asc_onChangeCommentData", (id, data) => {
      if (!isResolvedComment(data)) {
        return;
      }

      removeResolvedComment(id);
    });
  }

  private getDocumentPermissions(editing: boolean) {
    const doc = this.server.getDocument();
    return {
      edit: editing && doc.fileType !== "pdf",
      chat: false,
      rename: editing,
      protect: editing,
      // 允许接受/拒绝文档内已有修订；不自动进入「修订」录制模式
      review: true,
      print: false,
    };
  }

  /** 关闭 autosave 与保存按钮；保存快捷键由 installSaveShortcutBlocker 拦截。 */
  private buildEditorCustomization() {
    return {
      uiTheme: this.uiTheme,
      autosave: false,
      layout: {
        header: {
          save: false,
          // Word 头部「编辑 / 审阅 / 查看」切换（PPT/Excel 无此入口）
          editMode: false,
        },
        toolbar: {
          file: {
            save: false,
          },
          save: false,
        },
      },
      review: {
        trackChanges: false,
        showReviewChanges: false,
      },
      features: {
        spellcheck: {
          change: false,
        },
      },
      logo: {
        image: OFFICE_EDITOR_LOGO.image,
        imageDark: OFFICE_EDITOR_LOGO.imageDark,
      },
    };
  }

  /** 禁用 Ctrl/Cmd+S 与工具栏保存，避免与 export/downloadAs 共用管道冲突。 */
  private installSaveShortcutBlocker() {
    const win = this.getEditorFrameWindow();
    const doc = win?.document;

    if (!doc || win?.__ONLYOFFICE_SAVE_BLOCKED__) {
      return;
    }

    const blockSaveShortcut = (event: KeyboardEvent) => {
      const key = event.key.toLowerCase();
      if ((event.ctrlKey || event.metaKey) && key === "s") {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
      }
    };

    doc.addEventListener("keydown", blockSaveShortcut, true);
    win.__ONLYOFFICE_SAVE_BLOCKED__ = true;
  }

  /** 文档若自带 w:trackRevisions，OnlyOffice 默认会跟进入修订模式；接入层强制关闭录制。 */
  private applyDefaultReviewSettings() {
    const api = this.getSdkApi();
    api?.asc_SetGlobalTrackRevisions?.(false);
  }

  private mergeCommentItems(items: CommentItem[]) {
    for (const item of items) {
      if (isResolvedComment(item.Data)) {
        this.comments.delete(item.Id);
        continue;
      }

      this.comments.set(item.Id, item.Data);
    }
  }

  private fetchCommentsFromSdk(): CommentItem[] {
    const raw = this.getSdkApi()?.pluginMethod_GetAllComments?.();
    if (!Array.isArray(raw)) {
      return [];
    }

    return raw
      .map((item, index) => {
        const source =
          item && typeof item === "object"
            ? (item as Record<string, unknown>)
            : {};
        const id = String(source.Id ?? source.id ?? `comment-${index}`);
        const data = (source.Data ?? source.data ?? source) as CommentData;

        return { Id: id, Data: data };
      })
      .filter((item) => !isResolvedComment(item.Data));
  }

  private refreshCommentsFromSdk() {
    this.mergeCommentItems(this.fetchCommentsFromSdk());
  }

  private refreshRevisionsFromSdk() {
    const api = this.getSdkApi();
    const stackItems = normalizeRevisionItems(
      api?.asc_GetRevisionsChangesStack?.(),
    );
    const reportItems = flattenRevisionsReportByAuthors(
      api?.asc_GetTrackRevisionsReportByAuthors?.(),
    );

    if (
      stackItems.length > 0 &&
      (reportItems.length === 0 || stackItems.length >= reportItems.length)
    ) {
      this.revisions = stackItems;
      return;
    }

    if (reportItems.length > 0) {
      this.revisions = reportItems;
      return;
    }

    if (stackItems.length > 0) {
      this.revisions = stackItems;
    }
  }

  private syncRevisionsAfterMutation() {
    this.refreshRevisionsFromSdk();

    if (!this.haveRevisionsChanges()) {
      this.revisions = [];
    }
  }

  private applyAllRevisionChanges(mode: "accept" | "reject") {
    const api = this.requireSdkApi();
    const applyOne =
      mode === "accept"
        ? (raw: unknown) => api.asc_AcceptChanges?.(raw)
        : (raw: unknown) => api.asc_RejectChanges?.(raw);
    const applyBulk =
      mode === "accept"
        ? () => api.asc_AcceptChanges?.()
        : () => api.asc_RejectChanges?.();

    let guard = 0;
    let stagnant = 0;
    let lastCount = this.getAllRevisions().length;

    while (this.haveRevisionsChanges() && guard++ < 20) {
      this.refreshRevisionsFromSdk();
      const [first] = this.revisions;

      if (!first) {
        applyBulk();
        this.syncRevisionsAfterMutation();
        continue;
      }

      applyOne(first.Raw);
      this.syncRevisionsAfterMutation();

      const nextCount = this.getAllRevisions().length;
      if (nextCount >= lastCount && this.haveRevisionsChanges()) {
        stagnant += 1;
        if (stagnant >= 3) {
          applyBulk();
          this.syncRevisionsAfterMutation();
          stagnant = 0;
        }
      } else {
        stagnant = 0;
      }

      lastCount = nextCount;
    }

    this.syncRevisionsAfterMutation();
  }

  private teardownWordContentSync() {
    this.wordContentSyncTeardown?.();
    this.wordContentSyncTeardown = null;
    this.wordContentSyncPromise = null;
  }

  private scheduleWordContentSync() {
    window.setTimeout(() => {
      this.refreshCommentsFromSdk();
      this.refreshRevisionsFromSdk();
    }, 0);
  }

  private ensureWordContentSync() {
    if (this.fileType !== "docx" && getDocumentType(this.fileType) !== "word") {
      return Promise.resolve();
    }

    if (this.wordContentSyncPromise) {
      return this.wordContentSyncPromise;
    }

    this.wordContentSyncPromise = (async () => {
      const api = this.requireSdkApi();

      this.refreshCommentsFromSdk();
      this.refreshRevisionsFromSdk();

      const unsubscribers = await Promise.all([
        this.subscribe({
          type: "asc_onAddComment",
          fn: (id, data) => {
            const commentId = String(id);
            const commentData = data as CommentData;
            if (isResolvedComment(commentData)) {
              this.comments.delete(commentId);
              return;
            }

            this.comments.set(commentId, commentData);
          },
        }),
        this.subscribe({
          type: "asc_onChangeCommentData",
          fn: (id, data) => {
            const commentId = String(id);
            const commentData = data as CommentData;
            if (isResolvedComment(commentData)) {
              this.comments.delete(commentId);
              return;
            }

            this.comments.set(commentId, commentData);
          },
        }),
        this.subscribe({
          type: "asc_onRemoveComment",
          fn: (id) => {
            this.comments.delete(String(id));
          },
        }),
        this.subscribe({
          type: "asc_onShowRevisionsChange",
          fn: () => {
            this.refreshRevisionsFromSdk();
          },
        }),
      ]);

      this.wordContentSyncTeardown = () => {
        unsubscribers.forEach((unsubscribe) => unsubscribe?.());
      };

      this.scheduleWordContentSync();
      api.asc_Recalculate?.();
    })().catch(() => {
      this.teardownWordContentSync();
    });

    return this.wordContentSyncPromise;
  }

  private getRestrictionSdkApi() {
    return this.getSdkApi() as {
      asc_setRestriction?: (type: number) => void;
      asc_removeRestriction?: (type: number) => void;
      asc_setCanSendChanges?: (enabled: boolean) => void;
    } | undefined;
  }

  private getShellMainController() {
    const win = this.getEditorFrameWindow() as OnlyOfficeIframeWindow & {
      PE?: { getController?: (name: string) => ShellMainController };
      DE?: { getController?: (name: string) => ShellMainController };
      getApplication?: () => {
        getController?: (name: string) => ShellMainController;
      };
    };

    return (
      win?.PE?.getController?.("Main") ??
      win?.getApplication?.()?.getController?.("Main") ??
      win?.DE?.getController?.("Main")
    );
  }

  private getWordHeaderView() {
    const win = this.getEditorFrameWindow() as OnlyOfficeIframeWindow & {
      DE?: {
        getController?: (name: string) => {
          getView?: (name: string) => WordHeaderView;
        };
      };
    };

    return win?.DE?.getController?.("Viewport")?.getView?.("Common.Views.Header");
  }

  /** 隐藏 Word 头部「编辑 / 审阅 / 查看」切换（customization + 运行时兜底）。 */
  private hideWordDocModeSwitcher() {
    if (getDocumentType(this.fileType) !== DocumentType.Word) {
      return;
    }

    const header = this.getWordHeaderView();
    header?.btnDocMode?.setVisible?.(false);
    header?.btnPDFMode?.setVisible?.(false);

    const hedset = this.getEditorFrameWindow()?.document?.querySelector(
      "[data-layout-name=\"header-editMode\"]",
    );
    if (hedset instanceof HTMLElement) {
      hedset.style.display = "none";
    }
  }

  private scheduleWordDocModeHide() {
    this.hideWordDocModeSwitcher();
    window.setTimeout(() => this.hideWordDocModeSwitcher(), 0);
  }

  /** 同步 web-apps 工具栏/侧栏的 editing:disable（viewMode 与只读一致）。 */
  private syncShellEditingDisable(
    disabled: boolean,
    documentType = getDocumentType(this.fileType),
  ) {
    const nc = this.getEditorFrameWindow()?.Common?.NotificationCenter;
    if (!nc?.trigger) {
      return;
    }

    if (documentType === DocumentType.Slide) {
      nc.trigger("editing:disable", disabled, {
        viewMode: disabled,
        allowSignature: !disabled,
        rightMenu: { clear: false, disable: true },
        statusBar: true,
        leftMenu: { disable: disabled, previewMode: disabled },
        fileMenu: false,
        comments: { disable: false, previewMode: disabled },
        chat: false,
        review: true,
        viewport: disabled,
        documentHolder: { clear: disabled, disable: true },
        toolbar: true,
        header: { search: false },
        shortcuts: disabled ? false : undefined,
      });
      return;
    }

    if (documentType === DocumentType.Word) {
      if (disabled) {
        nc.trigger("editing:disable", true, {
          viewMode: true,
          reviewMode: false,
          fillFormMode: false,
          viewDocMode: false,
          allowMerge: false,
          allowSignature: false,
          allowProtect: false,
          rightMenu: { clear: true, disable: true },
          statusBar: true,
          leftMenu: { disable: true, previewMode: true },
          fileMenu: { protect: true, history: false },
          navigation: { disable: false, previewMode: true },
          comments: { disable: false, previewMode: true },
          chat: false,
          review: true,
          viewport: true,
          documentHolder: { clear: true, disable: true },
          toolbar: true,
          plugins: true,
          protect: true,
          header: { search: false, startfill: false },
          shortcuts: false,
        });
      } else {
        nc.trigger("editing:disable", false, {
          viewMode: false,
          reviewMode: false,
          fillFormMode: false,
          viewDocMode: false,
          allowMerge: true,
          allowSignature: false,
          allowProtect: false,
          rightMenu: { clear: false, disable: true },
          statusBar: true,
          leftMenu: { disable: false, previewMode: false },
          fileMenu: false,
          navigation: { disable: false, previewMode: false },
          comments: { disable: false, previewMode: false },
          chat: false,
          review: true,
          viewport: false,
          documentHolder: { clear: false, disable: true },
          toolbar: true,
          plugins: true,
          protect: true,
        });
      }

      this.scheduleWordDocModeHide();
      return;
    }

    nc.trigger("editing:disable", disabled, {
      viewMode: disabled,
      reviewMode: false,
      fillFormMode: false,
      viewDocMode: false,
      allowMerge: true,
      allowSignature: false,
      allowProtect: false,
      rightMenu: { clear: false, disable: true },
      statusBar: true,
      leftMenu: { disable: false, previewMode: disabled },
      fileMenu: false,
      navigation: { disable: false, previewMode: disabled },
      comments: { disable: false, previewMode: disabled },
      chat: false,
      review: true,
      viewport: false,
      documentHolder: { clear: false, disable: true },
      toolbar: true,
    });
  }

  /** PPT 在通用只读逻辑之上追加：锁定 Main 控制器 + 禁用幻灯片侧栏。 */
  private syncSlideReadOnlyExtras(locked: boolean) {
    if (getDocumentType(this.fileType) !== DocumentType.Slide || !locked) {
      return;
    }
    this.lockShellEditMode();
  }

  /**
   * processRightsChange(true) 在 OnlyOffice 内无效果；false 会 asc_coAuthoringDisconnect 且 mode.isEdit=false。
   * 本地/mock 场景用 asc_setRestriction + 外壳 UI 同步，避免切回编辑仍停留在只读。
   */
  private restoreShellEditMode() {
    const main = this.getShellMainController();

    if (main?.mode) {
      main.mode.isEdit = true;
      main.mode.canEdit = true;
    }

    this.getRestrictionSdkApi()?.asc_setCanSendChanges?.(true);
  }

  private lockShellEditMode() {
    const main = this.getShellMainController();

    if (main?.mode) {
      main.mode.isEdit = false;
      main.mode.canEdit = false;
    }

    this.getRestrictionSdkApi()?.asc_setCanSendChanges?.(false);
  }

  /** 兜底：拦截 SDK 层新增/复制幻灯片（toolbar 锁定之外）。 */
  private installSlideStructureEditBlocker() {
    if (getDocumentType(this.fileType) !== DocumentType.Slide) {
      return;
    }

    const patchApi = (
      api: {
        AddSlide?: (...args: unknown[]) => unknown;
        DublicateSlide?: (...args: unknown[]) => unknown;
        __ONLYOFFICE_SLIDE_BLOCK_PATCHED__?: boolean;
      } | undefined,
    ) => {
      if (!api || api.__ONLYOFFICE_SLIDE_BLOCK_PATCHED__) {
        return;
      }

      const guard = <T extends (...args: unknown[]) => unknown>(fn?: T) => {
        if (!fn) {
          return fn;
        }

        const bound = fn.bind(api);
        return (...args: unknown[]) => {
          if (this.readOnly) {
            return undefined;
          }
          return bound(...args);
        };
      };

      api.AddSlide = guard(api.AddSlide);
      api.DublicateSlide = guard(api.DublicateSlide);
      api.__ONLYOFFICE_SLIDE_BLOCK_PATCHED__ = true;
    };

    patchApi(
      this.getSdkApi() as unknown as {
        AddSlide?: (...args: unknown[]) => unknown;
        DublicateSlide?: (...args: unknown[]) => unknown;
        __ONLYOFFICE_SLIDE_BLOCK_PATCHED__?: boolean;
      },
    );
    patchApi(
      (this.getShellMainController() as {
        api?: {
          AddSlide?: (...args: unknown[]) => unknown;
          DublicateSlide?: (...args: unknown[]) => unknown;
          __ONLYOFFICE_SLIDE_BLOCK_PATCHED__?: boolean;
        };
      })?.api,
    );
  }

  /** downloadAs → /downloadas/ → 更新 fsMap 中的 Editor.bin。 */
  private async captureDocumentSnapshot() {
    if (!this.editor) {
      return this.server.getDocumentSnapshot();
    }

    return await this.server.captureCurrentDocument(() => {
      this.installIframeProxies();
      this.editor?.downloadAs("bin");
    });
  }

  /**
   * 只读模式下 downloadAs 可能被 SDK 拦截；导出前临时恢复编辑权再抓取。
   */
  private async captureDocumentSnapshotAllowingReadOnly() {
    if (!this.editor) {
      return this.server.getDocumentSnapshot();
    }

    const locked = this.readOnly;
    if (locked) {
      this.syncEditingRights(true);
    }

    try {
      return await this.captureDocumentSnapshot();
    } finally {
      if (locked) {
        this.syncEditingRights(false);
        this.syncSlideReadOnlyExtras(true);
      }
    }
  }

  private async captureDocumentIfDirty() {
    if (!this.editor || this.readOnly || !this.dirty) {
      return;
    }

    await this.captureDocumentSnapshot();
    this.dirty = false;
  }

  private destroyDocEditorInstance() {
    this.editor?.destroyEditor?.();
    this.editor = null;
    this.comments.clear();
    this.revisions = [];
    this.teardownWordContentSync();
  }

  /**
   * 初始只读与运行时切只读走同一套 asc_setRestriction。
   * 挂载阶段若 permissions.edit=false，xlsx 等会在打开时样式/格式异常；
   * 因此挂载时保持完整编辑权限，documentReady 后再施加只读限制。
   */
  private applyInitialReadOnlyState(documentType: DocumentType) {
    this.installSlideStructureEditBlocker();
    this.syncEditingRights(false);

    if (documentType === DocumentType.Slide) {
      // 工具栏 delayed render 后再锁一次，确保「新增幻灯片」按钮被 DisableToolbar 处理。
      window.setTimeout(() => {
        if (this.readOnly) {
          this.syncShellEditingDisable(true, documentType);
        }
      }, 0);
    }

    if (documentType === DocumentType.Cell) {
      this.getSdkApi()?.asc_Recalculate?.();
    }
  }

  /** 语言写在 iframe URL 的 lang 参数里，运行时 refreshFile 不会更新界面语言。 */
  private mountDocEditor() {
    const doc = this.server.getDocument();
    const user = this.server.getUser();
    const documentType = getDocumentType(doc.fileType);

    this.server.setClient({
      buildVersion: window.DocsAPI!.DocEditor.version(),
    });

    this.editor = new window.DocsAPI!.DocEditor(this.containerId, {
      document: {
        fileType: doc.fileType,
        key: doc.key,
        title: doc.title,
        url: doc.url,
        permissions: this.getDocumentPermissions(true),
      },
      documentType,
      editorConfig: {
        lang: this.editorLang,
        coEditing: {
          mode: "fast",
          change: false,
        },
        user: {
          ...user,
        },
        customization: this.buildEditorCustomization(),
      },
      events: {
        onAppReady: () => {
          // 尽早安装代理，供 WebSocket auth 与后续 downloadAs 使用。
          this.installIframeProxies();
        },
        onDocumentReady: () => {
          this.installSaveShortcutBlocker();
          this.installCommentResolveCleanup();
          this.installSlideStructureEditBlocker();
          this.applyDefaultReviewSettings();
          void this.ensureWordContentSync();
          onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.DOCUMENT_READY, {
            fileName: doc.title,
            fileType: doc.fileType,
            instanceId: this.instanceId,
          });

          if (this.readOnly) {
            this.applyInitialReadOnlyState(documentType);
          } else if (documentType === DocumentType.Word) {
            this.scheduleWordDocModeHide();
          }
        },
        onDocumentStateChange: (event: { data: boolean }) => {
          if (event.data) {
            this.dirty = true;
          }
        },
        // 不注册 onSave/onSaveDocument：内部保存已禁用，导出统一走 export() → downloadAs。
        onDownloadAs: () => {
          // Required so DocsAPI.downloadAs can request the current editor binary.
        },
      },
      type: "desktop",
      width: "100%",
      height: "100%",
    });
  }

  private buildRefreshFileConfig(editing: boolean) {
    const doc = this.server.getDocument();
    return {
      document: {
        fileType: doc.fileType,
        key: doc.key,
        title: doc.title,
        url: doc.url,
        permissions: this.getDocumentPermissions(editing),
      },
      documentType: getDocumentType(doc.fileType),
      editorConfig: {
        mode: editing ? "edit" : "view",
        lang: this.editorLang,
      },
      type: "desktop",
      width: "100%",
      height: "100%",
    };
  }

  /** 就地切换编辑权限（主路径 asc_setRestriction；PPT 只读额外 syncSlideReadOnlyExtras）。 */
  private syncEditingRights(editing: boolean) {
    if (!this.editor) {
      return;
    }

    const documentType = getDocumentType(this.fileType);
    const sdk = this.getRestrictionSdkApi();

    if (sdk?.asc_setRestriction) {
      if (editing) {
        sdk.asc_removeRestriction?.(ASC_RESTRICTION_VIEW);
        sdk.asc_setRestriction(ASC_RESTRICTION_NONE);
        this.restoreShellEditMode();
      } else {
        sdk.asc_setRestriction(ASC_RESTRICTION_VIEW);
        this.syncSlideReadOnlyExtras(true);
      }
      this.syncShellEditingDisable(!editing, documentType);
      return;
    }

    if (editing) {
      this.restoreShellEditMode();
      this.syncShellEditingDisable(false, documentType);
      this.editor.refreshFile?.(this.buildRefreshFileConfig(true));
    } else {
      this.syncSlideReadOnlyExtras(true);
      if (documentType !== DocumentType.Slide) {
        this.editor.denyEditingRights?.("");
      }
      this.syncShellEditingDisable(true, documentType);
    }
  }

  private createExportData(snapshot: ReturnType<EditorServer["getDocumentSnapshot"]>) {
    const binData = snapshot.binData;

    if (!binData) {
      throw new Error("No OnlyOffice document data is available to export");
    }

    return {
      fileName: snapshot.fileName || this.fileName,
      fileType: snapshot.fileType || this.fileType,
      binData,
      instanceId: this.instanceId,
      media: {
        ...snapshot.media,
        ...this.media,
      },
    };
  }

  private userSaveTimer: number | null = null;
  private pendingUserSaveSnapshot: ReturnType<
    EditorServer["getDocumentSnapshot"]
  > | null = null;

  /** 用户保存：更新快照并广播 SAVE_DOCUMENT + ONSAVE（同 tick 内合并重复回调）。 */
  private notifyUserSave(
    snapshot?: ReturnType<EditorServer["getDocumentSnapshot"]>,
  ) {
    if (snapshot) {
      this.pendingUserSaveSnapshot = snapshot;
    }

    if (this.userSaveTimer !== null) {
      window.clearTimeout(this.userSaveTimer);
    }

    this.userSaveTimer = window.setTimeout(() => {
      this.userSaveTimer = null;
      this.dirty = false;

      const snap =
        this.pendingUserSaveSnapshot ?? this.server.getDocumentSnapshot();
      this.pendingUserSaveSnapshot = null;

      const data = this.createExportData(snap);
      onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.SAVE_DOCUMENT, data);
      onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.ONSAVE, {
        fileName: data.fileName,
        instanceId: this.instanceId,
      });
    }, 0);
  }

  private isLoadSessionActive(
    containerId: string,
    loadSession?: number,
  ) {
    return (
      loadSession === undefined ||
      editorManagerFactory.isLoadSessionActive(containerId, loadSession)
    );
  }

  async create(options: CreateEditorViewOptions) {
    const containerId = options.containerId || this.containerId || ONLYOFFICE_ID;

    if (!this.isLoadSessionActive(containerId, options.loadSession)) {
      return this;
    }

    this.destroy();
    this.plugins = options.plugins || "featured";
    this.readOnly = !!options.readOnly;
    if (options.user) {
      this.server.setUser(options.user);
    }
    this.containerId = containerId;

    const fileType = getFileType(options.fileName, options.fileType);
    this.fileName = options.fileName;
    this.fileType = fileType;
    this.media = {};
    this.comments.clear();
    this.revisions = [];
    this.teardownWordContentSync();

    if (options.isNew) {
      this.server.openNew(fileType);
    } else if (options.file) {
      await this.server.open(options.file, {
        fileName: options.fileName,
        fileType,
      });
    } else if (options.url) {
      await this.server.openUrl(options.url, {
        fileName: options.fileName,
        fileType,
        loader: options.loader,
      });
    } else {
      throw new Error("OnlyOffice requires a file, url, or new document type");
    }

    if (!this.isLoadSessionActive(containerId, options.loadSession)) {
      return this;
    }

    await initializeOnlyOffice();

    if (!this.isLoadSessionActive(containerId, options.loadSession)) {
      return this;
    }

    this.editorLang = (options.lang as OnlyOfficeLang | undefined) || getOnlyOfficeLang();
    this.uiTheme = options.theme || "theme-white";

    this.mountDocEditor();

    return this;
  }

  exists() {
    return !!this.editor;
  }

  /** 导出链路：downloadAs("bin") → server.resolvePendingExport → SAVE_DOCUMENT 事件。 */
  async export() {
    let snapshot;
    if (this.editor && (!this.readOnly || this.dirty)) {
      snapshot = await this.captureDocumentSnapshotAllowingReadOnly();
      this.dirty = false;
    } else {
      snapshot = this.server.getDocumentSnapshot();
    }
    const data = this.createExportData(snapshot);

    onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.SAVE_DOCUMENT, data);

    return data;
  }

  getUser(): User {
    return this.server.getUser();
  }

  setUser(user: User) {
    this.server.setUser(user);
    this.editor?.setUsers?.([{ id: user.id, name: user.name }]);
  }

  /** 就地切换只读；切到只读前先 downloadAs 落盘，避免后续导出仍是打开时的 Editor.bin。 */
  async setReadOnly(readOnly: boolean) {
    if (this.readOnly === readOnly) {
      return;
    }

    onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, {
      loading: true,
    });

    try {
      if (readOnly && this.editor) {
        await this.captureDocumentSnapshot();
        this.dirty = false;
      }

      this.readOnly = readOnly;
      this.installSlideStructureEditBlocker();
      this.syncEditingRights(!readOnly);
    } finally {
      onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, {
        loading: false,
      });
    }
  }

  getReadOnly() {
    return this.readOnly;
  }

  async setLanguage(lang: OnlyOfficeLang) {
    if (this.editorLang === lang) {
      return;
    }

    this.editorLang = lang;

    if (!this.editor) {
      return;
    }

    onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, {
      loading: true,
    });

    try {
      await this.captureDocumentIfDirty();
      this.destroyDocEditorInstance();
      this.mountDocEditor();
    } finally {
      onlyofficeEventbus.emit(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, {
        loading: false,
      });
    }
  }

  getInstanceId() {
    return this.instanceId;
  }

  getContainerId() {
    return this.containerId;
  }

  getFileName() {
    return this.fileName;
  }

  getContainerParentSelector() {
    return `${ONLYOFFICE_CONTAINER_CONFIG.PARENT_SELECTOR}[data-onlyoffice-container-id="${this.containerId}"], ${ONLYOFFICE_CONTAINER_CONFIG.PARENT_SELECTOR}`;
  }

  getContainerStyle() {
    return ONLYOFFICE_CONTAINER_CONFIG.STYLE;
  }

  updateMedia(key: string, data: Uint8Array) {
    this.media[key] = data;
  }

  getMedia() {
    return { ...this.media };
  }

  isDirty() {
    return this.dirty;
  }

  async subscribe({
    type,
    fn,
  }: {
    type: AscWordApiMethod;
    fn: AscWordApiCallback;
  }) {
    const api = this.requireSdkApi();

    if (!api.asc_registerCallback || !api.asc_unregisterCallback) {
      throw new Error("OnlyOffice callback subscription is not supported");
    }

    api.asc_registerCallback(type, fn);

    return () => {
      api.asc_unregisterCallback?.(type, fn);
    };
  }

  getAllComments(): CommentItem[] {
    this.refreshCommentsFromSdk();

    return Array.from(this.comments.entries()).map(([Id, Data]) => ({
      Id,
      Data,
    }));
  }

  private createSdkCommentPayload(data: CommentData): unknown {
    const asc = this.getEditorFrameWindow()?.Asc as
      | (Record<string, unknown> & {
          asc_CCommentDataWord?: new (value: unknown) => {
            asc_putText?: (value: string) => void;
            asc_putUserName?: (value: string) => void;
            asc_putTime?: (value: string) => void;
            asc_putQuoteText?: (value: string) => void;
            asc_putSolved?: (value: boolean) => void;
            asc_putUserData?: (value: string) => void;
          };
        })
      | undefined;
    const CommentDataWord = asc?.asc_CCommentDataWord;

    if (!CommentDataWord) {
      return data;
    }

    const comment = new CommentDataWord(null);
    const payload = toPluginCommentPayload(data);

    if (payload.Text != null) {
      comment.asc_putText?.(String(payload.Text));
    }
    if (payload.UserName != null) {
      comment.asc_putUserName?.(String(payload.UserName));
    }
    if (payload.Time != null) {
      comment.asc_putTime?.(String(payload.Time));
    }
    if (payload.QuoteText != null) {
      comment.asc_putQuoteText?.(String(payload.QuoteText));
    }
    if (typeof payload.Solved === "boolean") {
      comment.asc_putSolved?.(payload.Solved);
    }
    if (payload.UserData != null) {
      comment.asc_putUserData?.(String(payload.UserData));
    }

    return comment;
  }

  addComment(input: CommentInput) {
    const api = this.requireSdkApi();
    const data = toPluginCommentPayload(normalizeCommentInput(input));
    const id =
      api.pluginMethod_AddComment?.(data) ??
      api.asc_addComment?.(this.createSdkCommentPayload(data) as CommentData);
    if (id) {
      this.comments.set(String(id), data);
    }
    return id ? String(id) : "";
  }

  updateComment(id: string, data: CommentData) {
    if (isResolvedComment(data)) {
      this.removeComment(id);
      return;
    }

    const api = this.requireSdkApi();
    const payload = toPluginCommentPayload(data);

    if (typeof api.pluginMethod_ChangeComment === "function") {
      api.pluginMethod_ChangeComment(id, payload);
    } else {
      api.asc_changeComment?.(
        id,
        this.createSdkCommentPayload(payload) as CommentData,
      );
    }
    this.comments.set(id, payload);
  }

  removeComment(id: string) {
    this.requireSdkApi().asc_removeComment?.(id);
    this.comments.delete(id);
  }

  goToComment(id: string, { showBalloon = false }: { showBalloon?: boolean } = {}) {
    const api = this.requireSdkApi();
    api.asc_selectComment?.(id);
    if (showBalloon) {
      api.asc_showComment?.(id);
    }
  }

  async registerCommentCallbacks(handlers: CommentChangeHandlers) {
    const unsubscribers = await Promise.all([
      handlers.onAdd
        ? this.subscribe({
            type: "asc_onAddComment",
            fn: (id, data) => {
              const commentId = String(id);
              const commentData = data as CommentData;
              this.comments.set(commentId, commentData);
              handlers.onAdd?.(commentId, commentData);
            },
          })
        : undefined,
      handlers.onChange
        ? this.subscribe({
            type: "asc_onChangeCommentData",
            fn: (id, data) => {
              const commentId = String(id);
              const commentData = data as CommentData;
              if (isResolvedComment(commentData)) {
                window.setTimeout(() => {
                  this.removeComment(commentId);
                  handlers.onRemove?.(commentId);
                }, 0);
                return;
              }

              this.comments.set(commentId, commentData);
              handlers.onChange?.(commentId, commentData);
            },
          })
        : undefined,
      handlers.onRemove
        ? this.subscribe({
            type: "asc_onRemoveComment",
            fn: (id) => {
              const commentId = String(id);
              this.comments.delete(commentId);
              handlers.onRemove?.(commentId);
            },
          })
        : undefined,
    ]);

    return () => {
      unsubscribers.forEach((unsubscribe) => unsubscribe?.());
    };
  }

  setTrackRevisions(enabled: boolean) {
    this.requireSdkApi().asc_SetGlobalTrackRevisions?.(enabled);
  }

  isTrackRevisions() {
    return !!this.getSdkApi()?.asc_GetGlobalTrackRevisions?.();
  }

  haveRevisionsChanges() {
    const api = this.getSdkApi();
    if (typeof api?.asc_HaveRevisionsChanges === "function") {
      return !!api.asc_HaveRevisionsChanges();
    }

    return this.getAllRevisions().length > 0;
  }

  getAllRevisions(): RevisionItem[] {
    this.refreshRevisionsFromSdk();

    const api = this.getSdkApi();
    if (
      typeof api?.asc_HaveRevisionsChanges === "function" &&
      !api.asc_HaveRevisionsChanges()
    ) {
      this.revisions = [];
    }

    return this.revisions;
  }

  goToNextRevision() {
    const [revision] = this.getAllRevisions();
    if (revision) this.goToRevision(revision.Id);
  }

  goToPrevRevision() {
    const revisions = this.getAllRevisions();
    const revision = revisions[revisions.length - 1];
    if (revision) this.goToRevision(revision.Id);
  }

  goToRevision(id: string) {
    const revision = this.getAllRevisions().find((item) => item.Id === id);
    const raw = revision?.Raw;
    if (raw && typeof raw === "object") {
      const candidate = raw as Record<string, unknown>;
      const goTo =
        candidate.GoTo ||
        candidate.goTo ||
        candidate.Select ||
        candidate.select;

      if (typeof goTo === "function") {
        goTo.call(raw);
      }
    }
  }

  acceptRevision(revision: RevisionItem | string) {
    const item =
      typeof revision === "string"
        ? this.getAllRevisions().find((entry) => entry.Id === revision)
        : revision;
    if (item) {
      this.requireSdkApi().asc_AcceptChanges?.(item.Raw);
      this.syncRevisionsAfterMutation();
    }
  }

  rejectRevision(revision: RevisionItem | string) {
    const item =
      typeof revision === "string"
        ? this.getAllRevisions().find((entry) => entry.Id === revision)
        : revision;
    if (item) {
      this.requireSdkApi().asc_RejectChanges?.(item.Raw);
      this.syncRevisionsAfterMutation();
    }
  }

  acceptAllRevisions() {
    this.applyAllRevisionChanges("accept");
  }

  rejectAllRevisions() {
    this.applyAllRevisionChanges("reject");
  }

  acceptRevisionsBySelection(all?: boolean) {
    this.requireSdkApi().asc_AcceptChangesBySelection?.(all);
  }

  rejectRevisionsBySelection(all?: boolean) {
    this.requireSdkApi().asc_RejectChangesBySelection?.(all);
  }

  async registerRevisionCallbacks(handlers: RevisionChangeHandlers) {
    const unsubscribers = await Promise.all([
      handlers.onShowChanges
        ? this.subscribe({
            type: "asc_onShowRevisionsChange",
            fn: (items) =>
              handlers.onShowChanges?.(normalizeRevisionItems(items)),
          })
        : undefined,
      handlers.onTrackRevisionsChange
        ? this.subscribe({
            type: "asc_onOnTrackRevisionsChange",
            fn: (enabled) => handlers.onTrackRevisionsChange?.(!!enabled),
          })
        : undefined,
    ]);
    return () => {
      unsubscribers.forEach((unsubscribe) => unsubscribe?.());
    };
  }

  destroy() {
    if (this.userSaveTimer !== null) {
      window.clearTimeout(this.userSaveTimer);
      this.userSaveTimer = null;
    }
    this.pendingUserSaveSnapshot = null;
    this.teardownWordContentSync();
    this.editor?.destroyEditor?.();
    this.editor = null;
    this.dirty = false;
    this.comments.clear();
    this.revisions = [];
    this.server.reset();
  }
}

class EditorManagerFactory {
  private defaultManager = new EditorManager();
  private managers = new Map<string, EditorManager>();
  private loadSessions = new Map<string, number>();

  beginLoadSession(containerId: string) {
    const next = (this.loadSessions.get(containerId) ?? 0) + 1;
    this.loadSessions.set(containerId, next);
    return next;
  }

  isLoadSessionActive(containerId: string, loadSession: number) {
    return this.loadSessions.get(containerId) === loadSession;
  }

  getDefault() {
    return this.defaultManager;
  }

  create(containerId: string) {
    const manager = this.managers.get(containerId) || new EditorManager(containerId);
    this.managers.set(containerId, manager);
    return manager;
  }

  get(containerId: string) {
    return this.managers.get(containerId) || this.create(containerId);
  }

  getAll() {
    return [this.defaultManager, ...this.managers.values()];
  }

  destroy(containerId: string) {
    const manager = this.managers.get(containerId);
    manager?.destroy();
    this.managers.delete(containerId);
  }

  destroyAll() {
    this.defaultManager.destroy();
    for (const manager of this.managers.values()) {
      manager.destroy();
    }
    this.managers.clear();
  }
}
export const editorManagerFactory = new EditorManagerFactory();
export const editorManager = editorManagerFactory.getDefault();
if (typeof window !== "undefined") {
  (window as any).editorManagerFactory = editorManagerFactory;
} 
