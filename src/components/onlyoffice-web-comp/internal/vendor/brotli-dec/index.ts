/** Vendored Brotli decompressor (MIT, brotli.js / Google). See dec/decode.js header. */
import { BrotliDecompressBuffer } from "./dec/decode.js";

export function brotliDecompress(input: Uint8Array): Uint8Array {
  return BrotliDecompressBuffer(input) as Uint8Array;
}
