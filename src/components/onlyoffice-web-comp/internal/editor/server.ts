import { converter } from "./x2t";
import { MockSocket } from "./socket";
import { User, Participant, AscSaveTypes, ServerOptions } from "./types";
import { emptyDocx, emptyPdf, emptyPptx, emptyXlsx } from "./empty";
import { convertCsvBufferToXlsxBuffer } from "./csv-to-xlsx";
import {
  getDocumentType,
  getFileExt,
  getX2tConvertFormats,
  getX2tCsvConvertOptions,
  sanitizeCsvBufferForX2t,
  isMultilineCsv,
} from "./utils";
import { allPlugins, featuredPlugins, getPluginsData } from "./plugins";

/**
 * Mock OnlyOffice 协作服务：维护 fsMap（Editor.bin + media），处理 WebSocket 与 /downloadas/ HTTP。
 *
 * 关键链路：
 * 打开 — loadDocument：x2t doc.* → Editor.bin
 * 导出 — captureCurrentDocument + downloadAs → /downloadas/ → resolvePendingExport
 * 保存 — 同 URL 无 pendingExport 时 commitUserSave（UI 已禁用，兜底保留）
 */

function mergeBuffers(buffers: Uint8Array[]) {
  const totalLength = buffers.reduce((acc, buffer) => acc + buffer.length, 0);
  const mergedBuffer = new Uint8Array(totalLength);
  let offset = 0;
  for (const buffer of buffers) {
    mergedBuffer.set(buffer, offset);
    offset += buffer.length;
  }
  return mergedBuffer;
}

/** OnlyOffice 画布 bin 魔数；x2t 只接受 XLSY/DOCY/PPTY，非法数据会导致导出失败。 */
function isValidEditorBin(data: Uint8Array) {
  if (data.length < 4) {
    return false;
  }

  const magic = String.fromCharCode(data[0], data[1], data[2], data[3]);
  return magic === "XLSY" || magic === "DOCY" || magic === "PPTY";
}

function randomId() {
  return Math.random().toString(36).substring(2, 9);
}

type CoAuthoringLockBlock =
  | string
  | number
  | { guid?: string; [key: string]: unknown };

function normalizeCoAuthoringLockBlocks(block: unknown): CoAuthoringLockBlock[] {
  if (Array.isArray(block)) {
    return block as CoAuthoringLockBlock[];
  }
  return [block as CoAuthoringLockBlock];
}

function getCoAuthoringLockKey(
  block: CoAuthoringLockBlock,
  isSpreadsheet: boolean,
) {
  if (
    isSpreadsheet &&
    typeof block === "object" &&
    block !== null &&
    block.guid != null
  ) {
    return String(block.guid);
  }
  return String(block);
}

function buildCoAuthoringLocks(
  block: unknown,
  fileType: string,
  userId?: string,
) {
  const isSpreadsheet = getDocumentType(fileType) === "cell";
  const time = +new Date();
  const locks: Record<
    string,
    { time: number; user?: string; block: CoAuthoringLockBlock }
  > = {};

  for (const item of normalizeCoAuthoringLockBlocks(block)) {
    const key = getCoAuthoringLockKey(item, isSpreadsheet);
    locks[key] = {
      time,
      user: userId,
      block: item,
    };
  }

  return locks;
}

function getUrl(data: Uint8Array, type?: string) {
  const blob = new Blob([data as Uint8Array<ArrayBuffer>], {
    type: type || "application/octet-stream",
  });
  return URL.createObjectURL(blob);
}

export class EditorServer {
  private id = "";
  private sockets = new Set<MockSocket>();
  private messageHandlers = new WeakMap<
    MockSocket,
    (msg: unknown, ...args: unknown[]) => void
  >();
  private sessionId: string = "session-id";
  private user: User = {
    id: "uid",
    name: "Me",
  };
  private client = {
    buildVersion: "9.3.0",
    buildNumber: 8,
  };
  private participants: Participant[] = [];
  private syncChangesIndex = 0;
  private loadPromise: Promise<void> | null = null;

  private file: File | null = null;
  private fileType: string = "docx";
  private title: string = "";
  private fsMap: Map<string, Uint8Array> = new Map();
  private urlsMap: Map<string, string> = new Map();

  private downloadId: string = "";
  /** downloadAs multipart 分片缓冲；保存与导出共用 HTTP 管道，需与 pendingExport 配合区分意图。 */
  private downloadParts: Uint8Array[] = [];
  /** 用户保存 downloadAs 进行中时阻塞 export，避免分片交错污染 Editor.bin。 */
  private savingDone: Promise<void> = Promise.resolve();
  private finishSaving: (() => void) | null = null;
  /** export() 调用 downloadAs("bin") 后等待 resolvePendingExport 完成。 */
  private pendingExport:
    | {
        resolve: (snapshot: ReturnType<EditorServer["getDocumentSnapshot"]>) => void;
        reject: (error: Error) => void;
        timer: number;
      }
    | null = null;

  private options: ServerOptions = {};

  constructor(options: ServerOptions = {}) {
    this.options = options;
    this.handleConnect = this.handleConnect.bind(this);
    this.handleMessage = this.handleMessage.bind(this);
  }

  reset() {
    if (this.pendingExport) {
      window.clearTimeout(this.pendingExport.timer);
      this.pendingExport.reject(new Error("Editor server reset"));
      this.pendingExport = null;
    }

    for (const socket of this.sockets) {
      const handler = this.messageHandlers.get(socket);
      if (handler) {
        socket.server.off("message", handler);
      }
    }
    this.sockets.clear();
    this.messageHandlers = new WeakMap();

    if (this.urlsMap.size > 0) {
      this.urlsMap.forEach((url) => URL.revokeObjectURL(url));
    }

    this.id = "";
    this.file = null;
    this.fileType = "docx";
    this.title = "";
    this.fsMap.clear();
    this.urlsMap.clear();
    this.loadPromise = null;
    this.downloadId = "";
    this.downloadParts = [];
    this.endSaving();
    this.syncChangesIndex = 0;
    this.participants = [];
  }

  async open(
    file: File,
    { fileType, fileName }: { fileType?: string; fileName?: string } = {},
  ) {
    const title = fileName || file.name;
    this.fileType = fileType || getFileExt(file.name) || "docx";
    const documentType = getDocumentType(this.fileType);
    this.id = randomId();
    this.file = file;
    this.title = title;
    const buffer = await file.arrayBuffer();
    this.loadPromise = this.loadDocument(buffer, this.fileType);

    return {
      id: this.id,
      documentType,
    };
  }

  openNew(fileType?: string) {
    this.fileType = fileType || "docx";
    this.id = randomId();
    this.file = null;
    this.loadPromise = null;
    this.title = "New Document";
    const documentType = getDocumentType(this.fileType);

    let binData: Uint8Array | null = null;

    switch (documentType) {
      case "word":
        binData = Uint8Array.from(emptyDocx, (v) => v.charCodeAt(0));
        break;
      case "cell":
        binData = Uint8Array.from(emptyXlsx, (v) => v.charCodeAt(0));
        break;
      case "slide":
        binData = Uint8Array.from(emptyPptx, (v) => v.charCodeAt(0));
        break;
      case "pdf":
        binData = Uint8Array.from(emptyPdf, (v) => v.charCodeAt(0));
        break;
    }

    if (!binData) {
      throw new Error("Failed to create new document");
    }

    this.fsMap.set("Editor.bin", binData);
    this.urlsMap.set("Editor.bin", getUrl(binData));

    return {
      id: this.id,
      documentType: documentType,
    };
  }

  async openUrl(
    url: string,
    {
      fileType,
      fileName,
      loader = (url: string) => fetch(url).then((res) => res.arrayBuffer()),
    }: {
      fileType?: string;
      fileName?: string;
      loader?: (url: string) => Promise<ArrayBuffer>;
    } = {},
  ) {
    const title = fileName || decodeURIComponent(url.split("/").pop() || "Document")
    this.fileType = fileType || getFileExt(title) || "docx";
    const documentType = getDocumentType(this.fileType);
    this.id = randomId();
    this.title = title;
    this.loadPromise = this.loadDocument(() => loader(url), this.fileType);

    return {
      id: this.id,
      documentType,
    };
  }

  getDocument() {
    if (!this.id) {
      this.openNew();
    }

    return {
      fileType: this.fileType,
      key: this.id,
      title: this.title,
      url: "/" + this.id,
    };
  }

  getUser() {
    return { ...this.user };
  }

  setUser(user: Partial<User>) {
    this.user = {
      ...this.user,
      ...user,
    };
  }

  getDocumentSnapshot() {
    const binData = this.fsMap.get("Editor.bin");
    const media = Object.fromEntries(
      Array.from(this.fsMap.entries()).filter(([key]) =>
        key.startsWith("media/"),
      ),
    );

    return {
      fileName: this.title,
      fileType: this.fileType,
      binData,
      media,
    };
  }

  /** 用户保存：更新 Editor.bin 并通知接入层，不触发浏览器下载。 */
  commitUserSave(data: Uint8Array) {
    if (!isValidEditorBin(data)) {
      console.warn(
        "[EditorServer] Ignoring invalid Editor.bin from save, length:",
        data.length,
      );
      return;
    }

    this.downloadParts = [];
    this.downloadId = "";
    this.updateEditorBin(data);
    this.options.onUserSave?.(this.getDocumentSnapshot());
  }

  private beginSaving() {
    this.savingDone = new Promise<void>((resolve) => {
      this.finishSaving = resolve;
    });
  }

  private endSaving() {
    this.finishSaving?.();
    this.finishSaving = null;
    this.savingDone = Promise.resolve();
  }

  /**
   * 导出链路：register pendingExport → trigger downloadAs("bin")
   * → iframe XHR/fetch 命中 /downloadas/ → resolvePendingExport 写入 Editor.bin。
   * 开始前 await savingDone 并清空 downloadParts，避免与保存分片冲突。
   */
  async captureCurrentDocument(
    trigger: () => void,
    timeout = 30000,
  ): Promise<ReturnType<EditorServer["getDocumentSnapshot"]>> {
    await this.savingDone;

    this.downloadParts = [];
    this.downloadId = "";

    if (this.pendingExport) {
      window.clearTimeout(this.pendingExport.timer);
      this.pendingExport.reject(
        new DOMException("OnlyOffice export was superseded", "AbortError"),
      );
      this.pendingExport = null;
    }

    return new Promise((resolve, reject) => {
      const timer = window.setTimeout(() => {
        this.pendingExport = null;
        reject(new Error("Timed out waiting for OnlyOffice export data"));
      }, timeout);

      this.pendingExport = { resolve, reject, timer };

      try {
        trigger();
      } catch (err) {
        window.clearTimeout(timer);
        this.pendingExport = null;
        reject(err instanceof Error ? err : new Error("Failed to start export"));
      }
    });
  }

  /** 打开文档：x2t 将 doc.{fileType} 转为 Editor.bin 写入 fsMap，供 iframe 加载。 */
  private async loadDocument(
    buffer: ArrayBuffer | (() => Promise<ArrayBuffer>),
    fileType: string,
  ) {
    if (typeof buffer == "function") {
      buffer = await buffer();
    }

    let output: Uint8Array | null = null;
    let media: { [key: string]: Uint8Array } = {};

    if (fileType == "pdf") {
      output = new Uint8Array(buffer);
    } else if (fileType === "csv") {
      ({ output, media } = await this.loadCsvDocument(buffer));
    } else {
      ({ output, media } = await this.convertBufferToEditorBin(buffer, fileType));
    }

    if (!output) {
      throw new Error(`Failed to convert ${fileType} file`);
    }

    if (this.urlsMap.size > 0) {
      this.urlsMap.forEach((url) => URL.revokeObjectURL(url));
    }
    this.fsMap.set("Editor.bin", output);
    this.urlsMap.set("Editor.bin", getUrl(output));
    for (const name in media) {
      this.addMedia(name, media[name]);
    }
  }

  private async convertBufferToEditorBin(buffer: ArrayBuffer, fileType: string) {
    const { formatFrom, formatTo } = getX2tConvertFormats(fileType);
    const result = await converter.convert({
      data: buffer,
      fileFrom: "doc." + fileType,
      fileTo: "Editor.bin",
      formatFrom,
      formatTo,
    });
    return { output: result.output, media: result.media };
  }

  private async loadCsvDocument(buffer: ArrayBuffer) {
    if (isMultilineCsv(buffer)) {
      return this.convertCsvViaXlsx(buffer);
    }

    const convertBuffer = sanitizeCsvBufferForX2t(buffer);
    const { formatFrom, formatTo } = getX2tConvertFormats("csv");

    try {
      const result = await converter.convert({
        data: convertBuffer,
        fileFrom: "doc.csv",
        fileTo: "Editor.bin",
        formatFrom,
        formatTo,
        ...getX2tCsvConvertOptions(convertBuffer),
      });
      if (result.output?.byteLength) {
        return { output: result.output, media: result.media };
      }
    } catch (err) {
      console.warn("[EditorServer] CSV x2t failed, retry via xlsx:", err);
    }

    return this.convertCsvViaXlsx(buffer);
  }

  private async convertCsvViaXlsx(buffer: ArrayBuffer) {
    const xlsxBuffer = await convertCsvBufferToXlsxBuffer(buffer);
    return this.convertBufferToEditorBin(xlsxBuffer, "xlsx");
  }

  private addMedia(name: string, data: Uint8Array) {
    const pathname = "media/" + name;
    const url = getUrl(data);
    this.fsMap.set(pathname, data);
    this.urlsMap.set(pathname, url);
    return url;
  }

  private updateEditorBin(data: Uint8Array) {
    this.fsMap.set("Editor.bin", data);
    this.urlsMap.set("Editor.bin", getUrl(data));
  }

  private resolvePendingExport(data: Uint8Array) {
    const pendingExport = this.pendingExport;
    if (!pendingExport) return false;

    window.clearTimeout(pendingExport.timer);
    this.pendingExport = null;

    // 校验魔数后再写入 fsMap，避免脏分片进入 x2t 导出链路。
    if (!isValidEditorBin(data)) {
      pendingExport.reject(
        new Error("OnlyOffice export returned invalid document data"),
      );
      return true;
    }

    this.updateEditorBin(data);
    pendingExport.resolve(this.getDocumentSnapshot());
    return true;
  }

  setClient(info: Partial<typeof this.client>) {
    this.client = {
      ...this.client,
      ...info,
    };
  }

  handleConnect({ socket }: { socket: MockSocket }) {
    console.log("connect: ", socket);

    this.sockets.add(socket);
    const { sessionId, client } = this;

    const readOnly = this.options.getState?.()?.readOnly ?? false;

    this.participants = [
      {
        connectionId: this.sessionId,
        encrypted: false,
        id: this.user.id,
        idOriginal: this.user.id,
        indexUser: 1,
        isCloseCoAuthoring: false,
        isLiveViewer: readOnly,
        username: this.user.name,
        view: readOnly,
      },
    ];

    const handler = (msg: unknown, ...args: unknown[]) => {
      void this.handleMessage(socket, msg as Record<string, unknown>, ...args);
    };
    this.messageHandlers.set(socket, handler);
    socket.server.on("message", handler);

    this.sendTo(socket, {
      maxPayload: 100000000,
      pingInterval: 25000,
      pingTimeout: 20000,
      sid: sessionId,
      upgrades: [],
    });

    this.sendTo(socket, {
      type: "license",
      license: {
        type: 3,
        buildNumber: client.buildNumber,
        buildVersion: client.buildVersion,
        light: false,
        mode: 0,
        rights: 1,
        protectionSupport: true,
        isAnonymousSupport: true,
        liveViewerSupport: true,
        branding: false,
        customization: true,
        advancedApi: false,
      },
    });
  }

  handleDisconnect({ socket }: { socket: MockSocket }) {
    console.log("disconnect: ", socket);

    const handler = this.messageHandlers.get(socket);
    if (handler) {
      socket.server.off("message", handler);
      this.messageHandlers.delete(socket);
    }
    this.sockets.delete(socket);
  }

  private sendTo(socket: MockSocket, ...msg: unknown[]) {
    console.log("[ws] >> ", ...msg);
    socket.server.emit("message", ...msg);
  }

  private broadcast(...msg: unknown[]) {
    for (const socket of this.sockets) {
      this.sendTo(socket, ...msg);
    }
  }

  async handleMessage(
    socket: MockSocket,
    msg: Record<string, unknown>,
    ...args: unknown[]
  ) {
    console.log("[ws] << ", msg, args);

    const send = (...payload: unknown[]) => this.sendTo(socket, ...payload);
    const { sessionId, participants, user, client } = this;
    const type =
      typeof msg === "object" && msg && "type" in msg ? msg.type : null;
    switch (type) {
      case "auth": {
        const changes: unknown[] = [];
        const readOnly = this.options.getState?.()?.readOnly ?? false;
        send({
          type: "authChanges",
          changes: changes,
        });
        send({
          type: "auth",
          result: 1,
          sessionId: sessionId,
          participants: participants,
          locks: [],
          //   changes: changes,
          //   changesIndex: 0,
          indexUser: 1,
          buildVersion: client.buildVersion || "9.3.0",
          buildNumber: client.buildNumber || 9,
          licenseType: 3,
          editorType: 2,
          mode: readOnly ? "view" : "edit",
          permissions: {
            comment: true,
            chat: true,
            download: true,
            edit: !readOnly,
            fillForms: false,
            modifyFilter: !readOnly,
            protect: !readOnly,
            print: true,
            review: true,
            copy: true,
          },
        });

        try {
          if (this.loadPromise) {
            await this.loadPromise;
          }
          send({
            type: "documentOpen",
            data: {
              type: "open",
              status: "ok",
              data: {
                ...Object.fromEntries(this.urlsMap),
              },
            },
          });
        } catch (err) {
          console.error(err);
          send({
            type: "documentOpen",
            data: {
              type: "open",
              status: "err",
              data: err instanceof Error ? err.message : String(err),
            },
          });
        }
        break;
      }
      case "isSaveLock":
        send({
          type: "saveLock",
          saveLock: false,
        });
        break;
      case "saveChanges":
        send({
          type: "unSaveLock",
          index:
            typeof msg.startSaveChanges === "number"
              ? msg.startSaveChanges
              : typeof msg.endSaveChanges === "number"
                ? msg.endSaveChanges
                : -1,
          syncChangesIndex: ++this.syncChangesIndex,
          time: +new Date(),
        });
        break;
      case "getLock":
        if (msg.block == null) {
          break;
        }

        {
          const locks = buildCoAuthoringLocks(
            msg.block,
            this.fileType,
            user?.id,
          );
          send({ type: "getLock", locks });
          send({ type: "releaseLock", locks });
        }
        break;
    }
  }

  /**
   * Mock 协作 HTTP 入口。OnlyOffice iframe 内 XHR/fetch 被代理到此。
   *
   * downloadAs 双路径（同一 URL，靠 pendingExport 区分）：
   * - 有 pendingExport → 导出：resolvePendingExport，写入 Editor.bin 供 x2t
   * - 无 pendingExport → 保存：commitUserSave（UI 保存已禁用，保留兜底）
   */
  async handleRequest(req: Request) {
    const u = new URL(req.url);

    const { id: key } = this;
    // console.log("[msg] server: ", u, key);

    if (u.pathname.endsWith("/downloadas/" + key)) {
      const cmd = JSON.parse(u.searchParams.get("cmd") || "{}");
      const buffer = await req.arrayBuffer();

      console.log("downloadAs -> ", cmd, buffer);

      const download = async () => {
        const input = mergeBuffers(this.downloadParts);
        const resolvedExport = this.resolvePendingExport(input);
        if (!resolvedExport) {
          // 用户保存（Ctrl+S / 工具栏保存）：保留 bin，走 EventBus，不触发浏览器下载。
          this.commitUserSave(input);
        }
        this.endSaving();
        return { status: "ok" as const, isExport: resolvedExport };
      };

      let result: { status: "ok"; isExport: boolean } = {
        status: "ok",
        isExport: false,
      };
      let isFinalChunk = false;

      // OnlyOffice downloadAs 按 PartStart → Part* → Complete(All) 分片 POST。
      switch (cmd.savetype) {
        case AscSaveTypes.PartStart:
          if (!this.pendingExport) {
            this.beginSaving();
          }
          this.downloadId = "_" + Math.round(Math.random() * 1000);
          this.downloadParts = [new Uint8Array(buffer)];
          break;
        case AscSaveTypes.Part:
          this.downloadParts.push(new Uint8Array(buffer));
          break;
        case AscSaveTypes.Complete:
          this.downloadParts.push(new Uint8Array(buffer));
          result = await download();
          this.downloadParts = [];
          isFinalChunk = true;
          break;
        case AscSaveTypes.CompleteAll:
          if (!this.pendingExport) {
            this.beginSaving();
          }
          this.downloadId = "_" + Math.round(Math.random() * 1000);
          this.downloadParts = [new Uint8Array(buffer)];
          result = await download();
          this.downloadParts = [];
          isFinalChunk = true;
          break;
      }

      // 仅在最终分片通知 SDK，结束 downloadAs 的「正在下载中」状态。
      // 中间分片若广播 save 会提前消费回调并误报「未知错误」。
      if (isFinalChunk) {
        const downloadUrl = this.urlsMap.get("Editor.bin") || "";
        setTimeout(() => {
          this.broadcast({
            type: "documentOpen",
            data: {
              type: "save",
              status: result.status,
              data: downloadUrl,
              // export 走 downloadAs("bin")，filetype 须为 bin；空 URL 会触发未知错误。
              filetype: result.isExport ? "bin" : this.fileType,
            },
          });
        }, 100);
      }

      return Response.json({
        status: result.status,
        type: "save",
        data: this.downloadId,
      });
    }

    if (u.pathname.endsWith("/upload/" + key)) {
      const buffer = await req.arrayBuffer();
      const data = new Uint8Array(buffer);
      const filename = Date.now() + ".png";
      const pathname = "media/" + filename;
      const url = this.addMedia(filename, data);
      return Response.json({ [pathname]: url });
    }

    if (u.pathname == "/plugins.json") {
      const state = this.options.getState?.();
      if (state?.plugins == "none") {
        return Response.json({ url: "", pluginsData: [], autostart: [] });
      }
      if (state?.plugins == "all") {
        return Response.json(getPluginsData(allPlugins));
      }
      return Response.json(getPluginsData(featuredPlugins));
    }

    return null;
  }
}
