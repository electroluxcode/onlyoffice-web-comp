#!/usr/bin/env node
/**
 * 将原始 TTF/OTF 转为 OnlyOffice public/fonts/{id} catalog 线格式。
 *
 * 源文件默认放在本脚本同目录（onlyoffice-comp/scripts/fonts/），例如 1000.ttf；
 * 产物写入 public/9.3.0/fonts/{id}。
 *
 * 用法:
 *   pnpm font:catalog -- --id 1000 --verify
 *   node onlyoffice-comp/scripts/fonts/ttf-to-catalog-font.mjs --id 1000
 *   node onlyoffice-comp/scripts/fonts/ttf-to-catalog-font.mjs ./其它路径/font.ttf --id 1000
 *   node onlyoffice-comp/scripts/fonts/ttf-to-catalog-font.mjs --decode --id 1000
 */

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ONLYOFFICE_FONT_XOR_KEY = [
  160, 102, 214, 32, 20, 150, 71, 250, 149, 105, 184, 80, 176, 65, 73, 72,
];

const __dirname = path.dirname(fileURLToPath(import.meta.url));
/** 放置待转换的 .ttf / .otf（如 1000.ttf） */
const DEFAULT_WORKSPACE_DIR = __dirname;
const REPO_ROOT = path.resolve(__dirname, "../../..");
const DEFAULT_OUTPUT_DIR = path.join(REPO_ROOT, "public/9.3.0/fonts");
const DEFAULT_ALLFONTS = path.join(REPO_ROOT, "public/9.3.0/sdkjs/common/AllFonts.js");

const FONT_EXTENSIONS = [".ttf", ".otf", ".tte", ".otc", ".ttc"];

function printUsage() {
  console.log(`
用法:
  编码（TTF/OTF → catalog 线格式）:
    pnpm font:catalog -- --id <fileId> [--verify]
    node onlyoffice-comp/scripts/fonts/ttf-to-catalog-font.mjs --id 1000

    未指定输入文件时，从工作目录读取 <id>.ttf / <id>.otf：
      ${DEFAULT_WORKSPACE_DIR}/

    也可显式指定源文件:
    node .../ttf-to-catalog-font.mjs <input.ttf> --id 1000 [--out <path>] [--verify]

  解码（线格式 → 可读 TTF）:
    node .../ttf-to-catalog-font.mjs --decode --id 1000
    node .../ttf-to-catalog-font.mjs --decode <wire-file> --out <output.ttf>

选项:
  --id <id>           OnlyOffice 字体文件 id（输出文件名，无扩展名）
  --workspace <dir>   源字体目录（默认: onlyoffice-comp/scripts/fonts）
  --out <path>        输出路径（编码默认: public/9.3.0/fonts/<id>）
  --fonts-dir <dir>   同 --out 的目录别名（编码产物目录）
  --allfonts <path>   AllFonts.js 路径
  --verify            编码后 XOR 还原并校验 TTF magic
  --decode            将 catalog 线格式还原为原始 TTF
  -h, --help          显示帮助
`);
}

function parseArgs(argv) {
  const args = {
    input: null,
    id: null,
    out: null,
    workspace: DEFAULT_WORKSPACE_DIR,
    fontsDir: DEFAULT_OUTPUT_DIR,
    allfonts: DEFAULT_ALLFONTS,
    verify: false,
    decode: false,
    help: false,
  };

  const positional = [];
  for (let i = 0; i < argv.length; i++) {
    const token = argv[i];
    switch (token) {
      case "-h":
      case "--help":
        args.help = true;
        break;
      case "--id":
        args.id = argv[++i];
        break;
      case "--out":
        args.out = path.resolve(argv[++i]);
        break;
      case "--workspace":
        args.workspace = path.resolve(argv[++i]);
        break;
      case "--fonts-dir":
        args.fontsDir = path.resolve(argv[++i]);
        break;
      case "--allfonts":
        args.allfonts = path.resolve(argv[++i]);
        break;
      case "--verify":
        args.verify = true;
        break;
      case "--decode":
        args.decode = true;
        break;
      default:
        if (token.startsWith("-")) {
          throw new Error(`未知参数: ${token}`);
        }
        positional.push(token);
    }
  }

  if (positional.length > 0) {
    args.input = path.resolve(positional[0]);
  }
  return args;
}

function listWorkspaceFonts(workspaceDir) {
  if (!fs.existsSync(workspaceDir)) {
    return [];
  }
  return fs
    .readdirSync(workspaceDir)
    .filter((name) => FONT_EXTENSIONS.some((ext) => name.toLowerCase().endsWith(ext)))
    .sort();
}

function resolveSourceInWorkspace(workspaceDir, fileId) {
  for (const ext of FONT_EXTENSIONS) {
    const candidate = path.join(workspaceDir, `${fileId}${ext}`);
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }
  return null;
}

function resolveInputPath(args) {
  if (args.input) {
    return args.input;
  }
  if (!args.id) {
    const available = listWorkspaceFonts(args.workspace);
    throw new Error(
      `请指定 --id <fileId>，或将源文件放到工作目录后使用 --id。\n` +
        `工作目录: ${args.workspace}\n` +
        (available.length
          ? `当前可用: ${available.join(", ")}`
          : `（目录为空，可放入如 ${args.id ?? "1000"}.ttf）`),
    );
  }

  const resolved = resolveSourceInWorkspace(args.workspace, args.id);
  if (!resolved) {
    const available = listWorkspaceFonts(args.workspace);
    throw new Error(
      `工作目录中未找到 ${args.id}.ttf / ${args.id}.otf 等源文件。\n` +
        `工作目录: ${args.workspace}\n` +
        (available.length ? `当前可用: ${available.join(", ")}` : "（目录为空）"),
    );
  }
  return resolved;
}

function isRawOpenTypeFont(data) {
  if (data.length < 4) return false;
  const sig = data.readUInt32BE(0);
  if (sig === 0x00010000 || sig === 0x00000100) return true;
  const tag = data.toString("ascii", 0, 4);
  return tag === "OTTO" || tag === "true" || tag === "typ1";
}

function applyXorInPlace(buffer, length = 32) {
  const n = Math.min(length, buffer.length);
  for (let i = 0; i < n; i++) {
    buffer[i] ^= ONLYOFFICE_FONT_XOR_KEY[i % ONLYOFFICE_FONT_XOR_KEY.length];
  }
}

function encodeToWireFormat(inputBuffer) {
  if (!isRawOpenTypeFont(inputBuffer)) {
    const magic = inputBuffer.slice(0, 4).toString("hex");
    throw new Error(
      `输入不是裸 TTF/OTF（magic=${magic}）。若已是 catalog 线格式，请勿重复编码。`,
    );
  }
  const wire = Buffer.from(inputBuffer);
  applyXorInPlace(wire);
  return wire;
}

function decodeFromWireFormat(wireBuffer) {
  const plain = Buffer.from(wireBuffer);
  applyXorInPlace(plain);
  if (!isRawOpenTypeFont(plain)) {
    const magic = plain.slice(0, 4).toString("hex");
    throw new Error(`解码后仍非 TTF/OTF（magic=${magic}），可能不是 OnlyOffice catalog 线格式。`);
  }
  return plain;
}

function readFontInternalFamilies(data) {
  const families = new Set();
  if (data.length < 12) return [];

  const numTables = data.readUInt16BE(4);
  for (let i = 0; i < numTables; i++) {
    const off = 12 + i * 16;
    if (off + 16 > data.length) break;

    const tag = data.toString("ascii", off, off + 4);
    if (tag !== "name") continue;

    const tblOff = data.readUInt32BE(off + 8);
    if (tblOff + 6 > data.length) return [...families];

    const count = data.readUInt16BE(tblOff + 2);
    const stringOffset = data.readUInt16BE(tblOff + 4);

    for (let j = 0; j < count; j++) {
      const rec = tblOff + 6 + j * 12;
      if (rec + 12 > data.length) break;

      const platform = data.readUInt16BE(rec);
      const nameId = data.readUInt16BE(rec + 6);
      const length = data.readUInt16BE(rec + 8);
      const offset = data.readUInt16BE(rec + 10);
      const start = tblOff + stringOffset + offset;
      const end = start + length;
      if (end > data.length || length === 0) continue;
      if (nameId !== 1 && nameId !== 4 && nameId !== 16) continue;

      let text = "";
      if (platform === 3) {
        text = new TextDecoder("utf-16be").decode(data.subarray(start, end));
      } else if (platform === 1) {
        text = new TextDecoder("latin1").decode(data.subarray(start, end));
      } else {
        continue;
      }

      text = text.replace(/\0/g, "").trim();
      if (text) families.add(text);
    }
    break;
  }

  return [...families];
}

function loadAllFontsFiles(allfontsPath) {
  if (!fs.existsSync(allfontsPath)) {
    return null;
  }
  const code = fs.readFileSync(allfontsPath, "utf8");
  const ctx = { window: {} };
  const fn = new Function("window", code + "\nreturn window.__fonts_files;");
  const files = fn(ctx);
  return Array.isArray(files) ? files : null;
}

function patchAllFontsCatalog(allfontsPath, fileId, families) {
  if (!fs.existsSync(allfontsPath)) {
    return false;
  }

  const files = loadAllFontsFiles(allfontsPath);
  if (!files) {
    return false;
  }

  let code = fs.readFileSync(allfontsPath, "utf8");
  let fileIndex = files.indexOf(fileId);
  const primary = families[0] ?? fileId;
  const infoEntry = `["${primary}",${fileIndex >= 0 ? fileIndex : files.length},0,-1,-1,-1,-1,-1,-1]`;

  if (fileIndex < 0) {
    code = code.replace(
      /(window\["__fonts_files"\]\s*=\s*\[[\s\S]*?)(\n\];)/,
      `$1,\n"${fileId}"$2`,
    );
    fileIndex = files.length;
  }

  const infoWithIndex = `["${primary}",${fileIndex},0,-1,-1,-1,-1,-1,-1]`;
  if (!code.includes(infoWithIndex)) {
    const escapedPrimary = primary.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const existing = new RegExp(
      `\\["${escapedPrimary}",\\d+,0,-1,-1,-1,-1,-1,-1\\]`,
    );
    if (existing.test(code)) {
      code = code.replace(existing, infoWithIndex);
    } else {
      code = code.replace(
        /(window\["__fonts_infos"\]\s*=\s*\[[\s\S]*?)(\n\];)/,
        `$1,\n${infoWithIndex}$2`,
      );
    }
  }

  fs.writeFileSync(allfontsPath, code);
  console.log(`\n[AllFonts.js] 已更新 catalog: "${fileId}" 下标 ${fileIndex}`);
  return true;
}

function printCatalogHints(fileId, families, allfontsPath) {
  const files = loadAllFontsFiles(allfontsPath);
  let fileIndex = -1;
  if (files) {
    fileIndex = files.indexOf(fileId);
    if (fileIndex < 0) {
      fileIndex = files.length;
      console.log("\n[AllFonts.js] 尚未包含此 id，需在 __fonts_files 末尾添加:");
      console.log(`  "${fileId}",`);
    } else {
      console.log(`\n[AllFonts.js] "${fileId}" 在 __fonts_files 下标: ${fileIndex}`);
    }
  } else {
    console.log(`\n[AllFonts.js] 未找到 ${allfontsPath}，请手动维护 catalog。`);
    fileIndex = "?";
  }

  const primary = families[0] ?? fileId;
  console.log("\n[name 表] internal family 候选:", families.join(", ") || "(无)");
  console.log("\n[__fonts_infos] 可参考（下标换成上面的 fileIndex）:");
  console.log(`  ["${primary}", ${fileIndex}, 0, -1, -1, -1, -1, -1, -1],`);
  console.log("\n[word-fonts/font-registry.ts] 转换后建议:");
  console.log(`  onlyOfficeFileId: "${fileId}",`);
  console.log(`  aliases 含文档名 + name 表名: ${families.map((n) => `"${n}"`).join(", ") || "(见 TTF)"}`);
}

function runEncode(args) {
  if (!args.id) {
    throw new Error("编码模式需要 --id <fileId>");
  }

  const inputPath = resolveInputPath(args);
  const inputBuffer = fs.readFileSync(inputPath);
  const wire = encodeToWireFormat(inputBuffer);
  const outPath = args.out ?? path.join(args.fontsDir, args.id);

  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, wire);

  const families = readFontInternalFamilies(inputBuffer);
  console.log(`源文件: ${inputPath}`);
  console.log(`已写入 catalog 线格式: ${outPath}`);
  console.log(`  原始: ${inputBuffer.length} 字节 → 线格式: ${wire.length} 字节`);
  console.log(`  线格式头 4 字节: ${wire.slice(0, 4).toString("hex")} (应为非 00010000)`);

  if (args.verify) {
    const roundtrip = decodeFromWireFormat(wire);
    const ok =
      roundtrip.length === inputBuffer.length &&
      roundtrip.equals(inputBuffer);
    console.log(`  往返校验: ${ok ? "通过" : "失败"}`);
    if (!ok) process.exitCode = 1;
  }

  patchAllFontsCatalog(args.allfonts, args.id, families);
  printCatalogHints(args.id, families, args.allfonts);
}

function runDecode(args) {
  if (args.input) {
    if (!args.out) {
      throw new Error("解码模式需要 --out <output.ttf>");
    }
    const wire = fs.readFileSync(args.input);
    const plain = decodeFromWireFormat(wire);
    fs.mkdirSync(path.dirname(args.out), { recursive: true });
    fs.writeFileSync(args.out, plain);
    const families = readFontInternalFamilies(plain);
    console.log(`已解码为 TTF: ${args.out} (${plain.length} 字节)`);
    console.log(`  name 表: ${families.join(", ") || "(无)"}`);
    return;
  }

  if (!args.id) {
    throw new Error("解码模式需要 --id <fileId> 或显式 <wire-file> --out <output.ttf>");
  }

  const wirePath = path.join(args.fontsDir, args.id);
  const outPath = args.out ?? path.join(args.workspace, `${args.id}-decoded.ttf`);
  const wire = fs.readFileSync(wirePath);
  const plain = decodeFromWireFormat(wire);
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, plain);
  const families = readFontInternalFamilies(plain);
  console.log(`线格式: ${wirePath}`);
  console.log(`已解码为 TTF: ${outPath} (${plain.length} 字节)`);
  console.log(`  name 表: ${families.join(", ") || "(无)"}`);
}

function main() {
  let args;
  try {
    args = parseArgs(process.argv.slice(2));
  } catch (err) {
    console.error(err.message);
    printUsage();
    process.exit(1);
  }

  if (args.help) {
    printUsage();
    process.exit(0);
  }

  if (process.argv.length <= 2 && !args.id) {
    printUsage();
    const available = listWorkspaceFonts(DEFAULT_WORKSPACE_DIR);
    if (available.length) {
      console.log(`\n工作目录中的源字体: ${available.join(", ")}`);
      console.log("示例: pnpm font:catalog -- --id 1000 --verify");
    }
    process.exit(1);
  }

  try {
    if (args.decode) {
      runDecode(args);
    } else {
      runEncode(args);
    }
  } catch (err) {
    console.error(err.message);
    process.exit(1);
  }
}

main();
