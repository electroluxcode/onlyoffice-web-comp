/// <reference lib="webworker" />

import { AvsFileType, X2tConvertParams, X2tConvertResult } from "./types";
import {
  fetchMaybeBrotliAsset,
  fetchMaybeBrotliScript,
} from "./fetch-brotli";
import { getX2tBaseUrl, resolveSiteUrl, STATIC_RESOURCE } from "../../const";

/**
 * X2T Converter Web Worker
 *
 * This worker handles CPU-intensive document conversion operations
 * off the main thread to prevent UI blocking.
 */

/* eslint-disable no-restricted-globals */

// Worker 内用 origin 拼绝对 URL（见 STATIC_RESOURCE.x2t）
const X2T_ORIGIN = self.location.origin;
const X2T_BASE_URL = getX2tBaseUrl(X2T_ORIGIN);

let x2t: any = null;
let initPromise: Promise<void> | null = null;

const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[i] = c >>> 0;
  }
  return table;
})();

function crc32(data: Uint8Array) {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function readU16(data: Uint8Array, offset: number) {
  return data[offset] | (data[offset + 1] << 8);
}

function readU32(data: Uint8Array, offset: number) {
  return (
    data[offset] |
    (data[offset + 1] << 8) |
    (data[offset + 2] << 16) |
    (data[offset + 3] << 24)
  ) >>> 0;
}

function writeU16(data: Uint8Array, offset: number, value: number) {
  data[offset] = value & 0xff;
  data[offset + 1] = (value >>> 8) & 0xff;
}

function writeU32(data: Uint8Array, offset: number, value: number) {
  data[offset] = value & 0xff;
  data[offset + 1] = (value >>> 8) & 0xff;
  data[offset + 2] = (value >>> 16) & 0xff;
  data[offset + 3] = (value >>> 24) & 0xff;
}

function concatBytes(parts: Uint8Array[]) {
  const size = parts.reduce((total, part) => total + part.length, 0);
  const output = new Uint8Array(size);
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.length;
  }
  return output;
}

async function decompressDeflateRaw(data: Uint8Array) {
  if (!("DecompressionStream" in self)) {
    return null;
  }

  const stream = new Blob([data as Uint8Array<ArrayBuffer>])
    .stream()
    .pipeThrough(new (self as any).DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function shouldRewriteZipEntry(name: string) {
  return (
    /\.(xml|rels)$/i.test(name) &&
    (name.startsWith("word/") ||
      name.startsWith("xl/") ||
      name.startsWith("ppt/"))
  );
}

function rewriteFontText(text: string, aliases?: Record<string, string>) {
  if (!aliases) return text;

  const entries = Object.entries(aliases)
    .filter(([from, to]) => from && to && from !== to)
    .sort(([a], [b]) => b.length - a.length);

  let output = text;
  for (const [from, to] of entries) {
    const escaped = from.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    // 避免「仿宋」替换命中已是「仿宋_GB2312」的前缀，产生 _GB2312_GB2312
    const suffix =
      to.startsWith(from) && to.length > from.length
        ? to.slice(from.length).replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
        : "";
    const pattern = suffix ? `${escaped}(?!${suffix})` : escaped;
    output = output.replace(new RegExp(pattern, "gi"), to);
  }
  return output;
}

function hasOfficeZipExtension(fileName: string) {
  return /\.(docx|xlsx|pptx)$/i.test(fileName);
}

function encodeUtf16Le(value: string) {
  const output = new Uint8Array(value.length * 2);
  for (let i = 0; i < value.length; i++) {
    const code = value.charCodeAt(i);
    output[i * 2] = code & 0xff;
    output[i * 2 + 1] = code >>> 8;
  }
  return output;
}

function replaceBytesWithPadding(
  data: Uint8Array,
  from: Uint8Array,
  to: Uint8Array,
  skipSuffix?: Uint8Array,
) {
  if (!from.length || to.length > from.length) return data;

  const output = data.slice();
  for (let i = 0; i <= output.length - from.length; i++) {
    let matched = true;
    for (let j = 0; j < from.length; j++) {
      if (output[i + j] !== from[j]) {
        matched = false;
        break;
      }
    }
    if (!matched) continue;

    if (skipSuffix && skipSuffix.length) {
      let suffixMatched = true;
      for (let k = 0; k < skipSuffix.length; k++) {
        if (output[i + from.length + k] !== skipSuffix[k]) {
          suffixMatched = false;
          break;
        }
      }
      if (suffixMatched) continue;
    }

    output.fill(0, i, i + from.length);
    output.set(to, i);
    i += from.length - 1;
  }

  return output;
}

function restoreEditorBinFontNames(
  data: Uint8Array,
  restoreAliases?: Record<string, string>,
) {
  if (!restoreAliases) return data;

  const entries = Object.entries(restoreAliases).sort(
    (a, b) => b[0].length - a[0].length,
  );

  let output = data;
  for (const [from, to] of entries) {
    const skipSuffix =
      to.startsWith(from) && to.length > from.length
        ? encodeUtf16Le(to.slice(from.length))
        : undefined;
    output = replaceBytesWithPadding(
      output,
      encodeUtf16Le(from),
      encodeUtf16Le(to),
      skipSuffix,
    );
  }
  return output;
}

async function rewriteZipFontNames(
  data: ArrayBuffer,
  fileFrom: string,
  fontAliases?: Record<string, string>,
) {
  if (
    !fontAliases ||
    !hasOfficeZipExtension(fileFrom) ||
    !("TextDecoder" in self) ||
    !("TextEncoder" in self)
  ) {
    return data;
  }

  const input = new Uint8Array(data);
  let eocd = -1;
  for (let i = input.length - 22; i >= Math.max(0, input.length - 65558); i--) {
    if (readU32(input, i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) return data;

  const entryCount = readU16(input, eocd + 10);
  const centralOffset = readU32(input, eocd + 16);
  const decoder = new TextDecoder();
  const encoder = new TextEncoder();
  const localParts: Uint8Array[] = [];
  const centralParts: Uint8Array[] = [];
  let readOffset = centralOffset;
  let outputOffset = 0;

  for (let i = 0; i < entryCount; i++) {
    if (readU32(input, readOffset) !== 0x02014b50) return data;

    const method = readU16(input, readOffset + 10);
    const modTime = readU16(input, readOffset + 12);
    const modDate = readU16(input, readOffset + 14);
    const crc = readU32(input, readOffset + 16);
    const compressedSize = readU32(input, readOffset + 20);
    const uncompressedSize = readU32(input, readOffset + 24);
    const nameLength = readU16(input, readOffset + 28);
    const extraLength = readU16(input, readOffset + 30);
    const commentLength = readU16(input, readOffset + 32);
    const internalAttrs = readU16(input, readOffset + 36);
    const externalAttrs = readU32(input, readOffset + 38);
    const localOffset = readU32(input, readOffset + 42);
    const nameBytes = input.slice(readOffset + 46, readOffset + 46 + nameLength);
    const name = decoder.decode(nameBytes);

    const localNameLength = readU16(input, localOffset + 26);
    const localExtraLength = readU16(input, localOffset + 28);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    const compressed = input.slice(dataStart, dataStart + compressedSize);
    let outputData = compressed;
    let outputMethod = method;
    let outputCrc = crc;
    let outputUncompressedSize = uncompressedSize;

    if (shouldRewriteZipEntry(name)) {
      const uncompressed =
        method === 0
          ? compressed
          : method === 8
            ? await decompressDeflateRaw(compressed)
            : null;
      if (uncompressed) {
        const rewritten = encoder.encode(
          rewriteFontText(decoder.decode(uncompressed), fontAliases),
        );
        outputData = rewritten;
        outputMethod = 0;
        outputCrc = crc32(rewritten);
        outputUncompressedSize = rewritten.length;
      }
    }

    const localHeader = new Uint8Array(30 + nameBytes.length);
    writeU32(localHeader, 0, 0x04034b50);
    writeU16(localHeader, 4, 20);
    writeU16(localHeader, 6, 0);
    writeU16(localHeader, 8, outputMethod);
    writeU16(localHeader, 10, modTime);
    writeU16(localHeader, 12, modDate);
    writeU32(localHeader, 14, outputCrc);
    writeU32(localHeader, 18, outputData.length);
    writeU32(localHeader, 22, outputUncompressedSize);
    writeU16(localHeader, 26, nameBytes.length);
    writeU16(localHeader, 28, 0);
    localHeader.set(nameBytes, 30);
    localParts.push(localHeader, outputData);

    const centralHeader = new Uint8Array(46 + nameBytes.length);
    writeU32(centralHeader, 0, 0x02014b50);
    writeU16(centralHeader, 4, 20);
    writeU16(centralHeader, 6, 20);
    writeU16(centralHeader, 8, 0);
    writeU16(centralHeader, 10, outputMethod);
    writeU16(centralHeader, 12, modTime);
    writeU16(centralHeader, 14, modDate);
    writeU32(centralHeader, 16, outputCrc);
    writeU32(centralHeader, 20, outputData.length);
    writeU32(centralHeader, 24, outputUncompressedSize);
    writeU16(centralHeader, 28, nameBytes.length);
    writeU16(centralHeader, 30, 0);
    writeU16(centralHeader, 32, 0);
    writeU16(centralHeader, 34, 0);
    writeU16(centralHeader, 36, internalAttrs);
    writeU32(centralHeader, 38, externalAttrs);
    writeU32(centralHeader, 42, outputOffset);
    centralHeader.set(nameBytes, 46);
    centralParts.push(centralHeader);

    outputOffset += localHeader.length + outputData.length;
    readOffset += 46 + nameLength + extraLength + commentLength;
  }

  const centralDirectory = concatBytes(centralParts);
  const end = new Uint8Array(22);
  writeU32(end, 0, 0x06054b50);
  writeU16(end, 8, entryCount);
  writeU16(end, 10, entryCount);
  writeU32(end, 12, centralDirectory.length);
  writeU32(end, 16, outputOffset);

  return concatBytes([...localParts, centralDirectory, end]).buffer;
}

/**
 * 在 Worker 全局作用域执行 Emscripten 脚本。
 * classic worker 用 importScripts；module worker 用间接 eval（全局作用域）。
 */
function executeEmscriptenScript(scriptSource: string): void {
  if (typeof importScripts === "function") {
    const scriptBlob = new Blob([scriptSource], {
      type: "application/javascript",
    });
    const scriptBlobUrl = URL.createObjectURL(scriptBlob);
    try {
      importScripts(scriptBlobUrl);
      return;
    } catch (error) {
      const isImportScriptsUnsupported =
        error instanceof TypeError &&
        String(error.message).includes("importScripts");
      if (!isImportScriptsUnsupported) {
        throw error;
      }
    } finally {
      URL.revokeObjectURL(scriptBlobUrl);
    }
  }

  // Module worker 不支持 importScripts；间接 eval 在 Worker 全局作用域执行。
  (0, eval)(scriptSource);
}

/**
 * Initialize x2t module in Worker context
 */
async function initX2t(): Promise<void> {
  if (x2t) return;

  const scriptUrl = resolveSiteUrl(X2T_ORIGIN, STATIC_RESOURCE.x2t.script);
  const wasmUrl = resolveSiteUrl(X2T_ORIGIN, STATIC_RESOURCE.x2t.wasm);

  const [scriptSource, wasmBinary] = await Promise.all([
    fetchMaybeBrotliScript(scriptUrl),
    fetchMaybeBrotliAsset(wasmUrl),
  ]);

  Object.assign(self, {
    __filename: X2T_BASE_URL,
    wasmBinary,
  });

  executeEmscriptenScript(scriptSource);

  x2t = (self as any).Module;

  // Wait for WASM runtime initialization
  await new Promise<void>((resolve) => {
    x2t.onRuntimeInitialized = () => resolve();
  });

  // Create working directories
  try {
    x2t.FS.mkdir("/working");
    x2t.FS.mkdir("/working/media");
    x2t.FS.mkdir("/working/fonts");
    x2t.FS.mkdir("/working/themes");
  } catch (err) {
    // Directories may already exist
    console.error("[x2t.worker] mkdir error:", err);
  }

  console.log("[x2t.worker] Initialized successfully");
}

/**
 * Ensure x2t is initialized before conversion
 */
async function ensureInit(): Promise<void> {
  if (!initPromise) {
    initPromise = initX2t();
  }
  return initPromise;
}

// Auto-initialize on worker creation
ensureInit().catch((err) => {
  console.error("[x2t.worker] Auto-init failed:", err);
});

/**
 * Clean up temporary files after conversion
 */
function cleanupFiles(files: string[]): void {
  for (const file of files) {
    try {
      x2t.FS.unlink(file);
    } catch (err) {
      console.error(err);
    }
  }
  cleanMedia();
  cleanFonts();
}

function cleanMedia() {
  try {
    const mediaFiles = x2t.FS.readdir("/working/media/");
    for (const file of mediaFiles) {
      if (file !== "." && file !== "..") {
        x2t.FS.unlink("/working/media/" + file);
      }
    }
  } catch (err) {
    console.error(err);
  }
}

function cleanFonts() {
  try {
    const fontFiles = x2t.FS.readdir("/working/fonts/");
    for (const file of fontFiles) {
      if (file !== "." && file !== "..") {
        x2t.FS.unlink("/working/fonts/" + file);
      }
    }
  } catch (err) {
    console.error(err);
  }
}

function writeFonts(fonts?: { [key: string]: Uint8Array }) {
  cleanFonts();
  if (!fonts) {
    return;
  }

  for (const [key, value] of Object.entries(fonts)) {
    try {
      x2t.FS.writeFile("/working/fonts/" + key, value);
    } catch (err) {
      console.error(key, err);
    }
  }
}

/**
 * Read media files from the working directory
 */
function readMedia(): { [key: string]: Uint8Array } {
  const media: { [key: string]: Uint8Array } = {};
  try {
    const files = x2t.FS.readdir("/working/media/");
    for (const file of files) {
      if (file !== "." && file !== "..") {
        const fileData = x2t.FS.readFile("/working/media/" + file, {
          encoding: "binary",
        });
        media[file] = fileData;
      }
    }
  } catch (e) {
    console.error(e);
  }
  return media;
}

const xmlPath = "/working/params.xml";

function writeInputs({
  fileFrom,
  fileTo,
  formatFrom,
  formatTo,
  data,
  media,
  csvEncoding,
  csvDelimiter,
  csvDelimiterChar,
}: X2tConvertParams) {
  const isCsvSource =
    formatFrom === AvsFileType.AVS_FILE_SPREADSHEET_CSV ||
    fileFrom.toLowerCase().endsWith(".csv");

  const params: Record<string, string | number | boolean> = {
    m_sFileFrom: fileFrom,
    m_sThemeDir: "/working/themes",
    m_sFileTo: fileTo,
    m_nFormatFrom: formatFrom ?? 0,
    m_nFormatTo: formatTo ?? 0,
    m_bIsPDFA: formatTo === AvsFileType.AVS_FILE_CROSSPLATFORM_PDFA,
    m_bIsNoBase64: true,
    m_sFontDir: "/working/fonts/",
  };

  if (isCsvSource) {
    params.m_nCsvTxtEncoding = csvEncoding ?? 46;
    params.m_nCsvDelimiter = csvDelimiter ?? 4;
    if (csvDelimiterChar) {
      params.m_nCsvDelimiterChar = csvDelimiterChar;
    }
  }

  const content = Object.entries(params)
    .filter(([, v]) => v !== undefined && v !== null && v !== "")
    .reduce((a, [k, v]) => a + `<${k}>${v}</${k}>\n`, "");

  const xml = `<?xml version="1.0" encoding="utf-8"?>
<TaskQueueDataConvert
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:xsd="http://www.w3.org/2001/XMLSchema"
>
${content}
</TaskQueueDataConvert>`;

  x2t.FS.writeFile(xmlPath, xml);
  if (data) {
    x2t.FS.writeFile(fileFrom, new Uint8Array(data));
  }

  if (media) {
    cleanMedia();
    for (const [key, value] of Object.entries(media)) {
      try {
        x2t.FS.writeFile("/working/" + key, value);
      } catch (err) {
        console.error(key, err);
      }
    }
  }
}

/**
 * Convert document from one format to another
 */
async function convert({
  data,
  fileFrom,
  fileTo,
  formatFrom,
  formatTo,
  media,
  fonts,
  fontAliases,
  fontExportAliases,
  themes,
  csvEncoding,
  csvDelimiter,
  csvDelimiterChar,
}: X2tConvertParams): Promise<X2tConvertResult> {
  const preparedData = data
    ? await rewriteZipFontNames(data, fileFrom, fontAliases)
    : data;
  const fromPath = "/working/" + fileFrom;
  const toPath = "/working/" + fileTo;
  const files = [fromPath, toPath, xmlPath];

  writeFonts(fonts);

  writeInputs({
    fileFrom: fromPath,
    fileTo: toPath,
    formatFrom,
    formatTo,
    data: preparedData,
    media,
    csvEncoding,
    csvDelimiter,
    csvDelimiterChar,
  });

  if (
    fileFrom.endsWith(".doc") ||
    formatFrom == AvsFileType.AVS_FILE_DOCUMENT_DOC
  ) {
    const viaPath = fromPath + ".docx";
    writeInputs({
      fileFrom: fromPath,
      fileTo: viaPath,
      data: null as never,
    });
    x2t.ccall("main1", ["number"], ["string"], [xmlPath]);
    writeInputs({
      fileFrom: viaPath,
      fileTo: toPath,
      data: null as never,
    });
    files.push(viaPath);
  }

  try {
    const pathInfo = x2t.FS.analyzePath(toPath);
    if (pathInfo.exists) {
      x2t.FS.unlink(toPath);
    }
  } catch (err) {}

  try {
    x2t.ccall("main1", ["number"], ["string"], [xmlPath]);
  } catch (e) {
    console.error("ccall", e);
  }

  // Read output file
  let output: Uint8Array | null = null;
  try {
    output = x2t.FS.readFile(toPath);
    if (output && fileTo === "Editor.bin" && fontAliases) {
      output = restoreEditorBinFontNames(output, fontAliases);
    }
    if (output && hasOfficeZipExtension(fileTo) && fontExportAliases) {
      const restored = await rewriteZipFontNames(
        output.slice(0).buffer,
        fileTo,
        fontExportAliases,
      );
      output = new Uint8Array(restored);
    }
  } catch (e) {
    console.error("[x2t.worker] read output failed:", e);
  }

  if (!output) {
    throw new Error(`x2t conversion produced no output (${fileFrom} -> ${fileTo})`);
  }

  // Read media files
  const outputMedia = readMedia();

  // Cleanup temporary files
  setTimeout(() => {
    cleanupFiles(files);
  });

  return { output, media: outputMedia };
}

// Message types
interface WorkerMessage {
  id?: number;
  type: "convert";
  payload?: any;
}

/**
 * Handle incoming messages from main thread
 */
self.onmessage = async (event: MessageEvent<WorkerMessage>) => {
  const { id, type, payload } = event.data;

  try {
    switch (type) {
      case "convert": {
        // Ensure x2t is initialized before conversion
        await ensureInit();
        const result = await convert(payload);

        // Use Transferable objects for zero-copy transfer
        const transferables: Transferable[] = [];
        if (result.output) {
          transferables.push(result.output.buffer);
        }
        Object.values(result.media).forEach((m) =>
          transferables.push(m.buffer),
        );

        self.postMessage(
          { id, type: "convert:done", payload: result },
          { transfer: transferables },
        );
        break;
      }

      default:
        self.postMessage({
          id,
          type: "error",
          error: `Unknown message type: ${type}`,
        });
    }
  } catch (error) {
    self.postMessage({
      id,
      type: "error",
      error: error instanceof Error ? error.message : String(error),
    });
  }
};

// Signal that worker is ready
self.postMessage({ type: "ready" });
