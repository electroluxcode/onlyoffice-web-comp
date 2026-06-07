import type { Metadata } from "next";
import { OfficePreviewPage } from "@/components/office/office-preview-page";
import { FILE_TYPE } from "@/onlyoffice-comp";

export const metadata: Metadata = {
  title: "Excel Preview — OnlyOffice MVP",
};

export default function ExcelBasePage() {
  return (
    <OfficePreviewPage
      title="Excel 预览"
      badge="E"
      badgeClassName="bg-gradient-to-br from-green-500 to-green-700"
      defaultFileName="New_Spreadsheet.xlsx"
      fileType={FILE_TYPE.XLSX}
      accept=".xlsx,.xls,.ods,.csv"
      newButtonLabel="新建 Excel"
    />
  );
}
