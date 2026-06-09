import { parseCsvBuffer } from "./utils";

async function loadExcelJS() {
  const mod = await import(
    "exceljs/dist/exceljs.min"
  );
  return mod.default ?? mod;
}

/** 复杂 CSV 先转 XLSX，再交给 x2t，避免 x2t CSV 解析崩溃。 */
export async function convertCsvBufferToXlsxBuffer(buffer: ArrayBuffer) {
  const ExcelJS = await loadExcelJS();
  const rows = parseCsvBuffer(buffer);
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Sheet1");

  for (const row of rows) {
    worksheet.addRow(row);
  }

  const output = await workbook.xlsx.writeBuffer();
  if (output instanceof ArrayBuffer) {
    return output;
  }
  const bytes = new Uint8Array(output as ArrayLike<number>);
  return bytes.slice().buffer;
}
