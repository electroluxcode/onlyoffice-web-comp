/**
 * 面向业务页面的 OnlyOffice 门面：收敛 initialize / 开文档 / 导出 / 只读 / 语言 等调用。
 *
 * 底层仍由 EditorManager 驱动；本类负责 document store 与 x2t 转换编排。
 */
import {
  ONLYOFFICE_ID,
  ONLYOFFICE_EVENT_KEYS,
  ONLYOFFICE_LANG_KEY,
  type FileType,
} from "../const";
import { getDocmentObj, setDocmentObj } from "../store/document";
import {
  getCurrentLang,
  getOnlyOfficeLang,
  setCurrentLang,
  type OnlyOfficeLang,
} from "../store/lang";
import { downloadBlob, getOfficeMimeType } from "../util/download";
import { initializeOnlyOffice } from "../util/initialize";
import { convertBinToDocument } from "../util/x2t";
import {
  EditorManager,
  editorManagerFactory,
} from "./editor-manager";
import { onlyofficeEventbus, type LoadingChangeData } from "./eventbus";

export type OnlyOfficeManagerOptions = {
  /** DOM 容器 id，默认 ONLYOFFICE_ID */
  containerId?: string;
  /** 导出时使用的 Office 主格式 */
  fileType: FileType;
  /** 首次 bootstrap 打开的默认文件名 */
  defaultFileName: string;
  readOnly?: boolean;
  lang?: OnlyOfficeLang;
  /** 页面初始化会话，用于忽略路由切换后过期的 openDocument */
  loadSession?: number;
};

export type OpenDocumentInput = {
  fileName: string;
  file?: File;
  isNew?: boolean;
  readOnly?: boolean;
  loadSession?: number;
};

export class OnlyOfficeManager {
  readonly containerId: string;
  readonly fileType: FileType;

  private editor: EditorManager;
  private readOnly: boolean;
  private ready = false;

  private constructor(
    editor: EditorManager,
    options: OnlyOfficeManagerOptions & { containerId: string },
  ) {
    this.containerId = options.containerId;
    this.fileType = options.fileType;
    this.editor = editor;
    this.readOnly = options.readOnly ?? false;
    if (options.lang) {
      setCurrentLang(options.lang);
    }
  }

  /** 多实例场景：绑定已有 EditorManager，不自动 open */
  static fromEditor(
    editor: EditorManager,
    options: OnlyOfficeManagerOptions & { containerId: string },
  ): OnlyOfficeManager {
    return new OnlyOfficeManager(editor, options);
  }

  /** 加载 DocsAPI 并打开 defaultFileName（新建或空白） */
  static async create(
    options: OnlyOfficeManagerOptions,
  ): Promise<OnlyOfficeManager> {
    const containerId = options.containerId ?? ONLYOFFICE_ID;
    await initializeOnlyOffice();

    const editor = editorManagerFactory.get(containerId);
    const manager = new OnlyOfficeManager(editor, { ...options, containerId });

    await manager.openDocument({
      fileName: options.defaultFileName,
      isNew: true,
      readOnly: options.readOnly,
      loadSession: options.loadSession,
    });

    return manager;
  }

  /** 加载 DocsAPI 并直接打开已有 File（先取文件再挂载） */
  static async createWithFile(
    options: OnlyOfficeManagerOptions,
    file: File,
  ): Promise<OnlyOfficeManager> {
    const containerId = options.containerId ?? ONLYOFFICE_ID;
    await initializeOnlyOffice();

    const editor = editorManagerFactory.get(containerId);
    const manager = new OnlyOfficeManager(editor, { ...options, containerId });

    await manager.openDocument({
      fileName: file.name,
      file,
      readOnly: options.readOnly,
      loadSession: options.loadSession,
    });

    return manager;
  }

  /** 打开/切换文档（上传、新建、重开） */
  async openDocument(input: OpenDocumentInput) {
    const readOnly = input.readOnly ?? this.readOnly;

    setDocmentObj({
      fileName: input.fileName,
      file: input.file,
      isNew: input.isNew ?? !input.file,
    });

    const { fileName, file } = getDocmentObj();

    await this.editor.create({
      file,
      fileName,
      isNew: !file,
      readOnly,
      lang: getOnlyOfficeLang(),
      containerId: this.containerId,
      editorManager: this.editor,
      loadSession: input.loadSession,
    });

    this.readOnly = readOnly;
    this.ready = true;
  }

  async openNew(fileName: string, readOnly?: boolean) {
    await this.openDocument({ fileName, isNew: true, readOnly });
  }

  async openFile(file: File, readOnly?: boolean) {
    await this.openDocument({ fileName: file.name, file, readOnly });
  }

  isReady() {
    return this.ready;
  }

  getReadOnly() {
    return this.readOnly;
  }

  setReadOnly(readOnly: boolean) {
    this.readOnly = readOnly;
    this.editor.setReadOnly(readOnly);
  }

  toggleReadOnly() {
    return this.setReadOnly(!this.readOnly);
  }

  getLanguage() {
    return getCurrentLang();
  }

  async setLanguage(lang: OnlyOfficeLang) {
    setCurrentLang(lang);
    await this.editor.setLanguage(lang);
  }

  /** 在中/英之间切换并应用到 iframe */
  async toggleLanguage() {
    const nextLang =
      getCurrentLang() === ONLYOFFICE_LANG_KEY.ZH
        ? ONLYOFFICE_LANG_KEY.EN
        : ONLYOFFICE_LANG_KEY.ZH;
    await this.setLanguage(nextLang);
    return nextLang;
  }

  getEditor() {
    return this.editor;
  }

  async exportDocument() {
    return this.editor.export();
  }

  /** 导出为 Office 文件 Blob：Editor.bin → x2t → doc.{fileType}。 */
  async exportAsBlob() {
    const binData = await this.editor.export();
    const result = await convertBinToDocument(
      binData.binData,
      binData.fileName,
      this.fileType,
      binData.media,
    );

    return {
      blob: new Blob([result.data as any], {
        type: getOfficeMimeType(this.fileType),
      }),
      fileName: result.fileName,
    };
  }

  /** 触发浏览器下载；内部先走 exportAsBlob 完整链路。 */
  async downloadExport() {
    const { blob, fileName } = await this.exportAsBlob();
    downloadBlob(blob, fileName);
  }

  onLoadingChange(handler: (data: LoadingChangeData) => void) {
    onlyofficeEventbus.on(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, handler);
    return () => {
      onlyofficeEventbus.off(ONLYOFFICE_EVENT_KEYS.LOADING_CHANGE, handler);
    };
  }

  destroy() {
    this.editor.destroy();
    this.ready = false;
  }
}

/** 多容器场景（如三栏 Word/Excel/PPT）按 containerId 缓存门面实例 */
export class OnlyOfficeManagerFactory {
  private managers = new Map<string, OnlyOfficeManager>();

  async open(
    options: OnlyOfficeManagerOptions,
    document: OpenDocumentInput,
  ): Promise<OnlyOfficeManager> {
    const containerId = options.containerId ?? ONLYOFFICE_ID;
    let manager = this.managers.get(containerId);

    if (!manager) {
      await initializeOnlyOffice();
      const editor = editorManagerFactory.get(containerId);
      manager = OnlyOfficeManager.fromEditor(editor, {
        ...options,
        containerId,
      });
      this.managers.set(containerId, manager);
    }

    await manager.openDocument({
      ...document,
      readOnly: document.readOnly ?? options.readOnly,
    });

    return manager;
  }

  get(containerId: string) {
    return this.managers.get(containerId);
  }

  destroy(containerId: string) {
    this.managers.get(containerId)?.destroy();
    this.managers.delete(containerId);
  }

  destroyAll() {
    for (const manager of this.managers.values()) {
      manager.destroy();
    }
    this.managers.clear();
  }
}

export const onlyOfficeManagerFactory = new OnlyOfficeManagerFactory();

/** 单页预览默认门面（懒创建，需先 OnlyOfficeManager.create 或 factory.open） */
let defaultManager: OnlyOfficeManager | null = null;

export function getDefaultOnlyOfficeManager() {
  return defaultManager;
}

export function setDefaultOnlyOfficeManager(manager: OnlyOfficeManager | null) {
  defaultManager = manager;
}
