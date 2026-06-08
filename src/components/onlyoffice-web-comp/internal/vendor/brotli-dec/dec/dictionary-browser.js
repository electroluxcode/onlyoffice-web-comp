function base64ToBytes(base64) {
  var binary = atob(base64);
  var bytes = new Uint8Array(binary.length);
  for (var i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

/**
 * Browser-friendly dictionary bootstrap: dictionary.bin.js is a compressed
 * copy of the static dictionary; decompress it on first use.
 */
exports.init = function () {
  var BrotliDecompressBuffer = require("./decode").BrotliDecompressBuffer;
  var compressed = base64ToBytes(require("./dictionary.bin.js"));
  return BrotliDecompressBuffer(compressed);
};
