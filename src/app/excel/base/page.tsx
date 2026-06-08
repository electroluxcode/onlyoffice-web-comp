import type { Metadata } from "next";
<<<<<<< HEAD
import { OfficePreviewPage } from "@/components/office/office-preview-page";
import { FILE_TYPE } from "@/onlyoffice-comp";
=======
import { OfficePreviewPage } from "@/components/onlyoffice-web-demo/office-preview-page";
import { FILE_TYPE } from "@/components/onlyoffice-web-comp";
>>>>>>> refactor/v9

export const metadata: Metadata = {
  title: "Excel Preview — OnlyOffice MVP",
};

export default function ExcelBasePage() {
  return (
    <OfficePreviewPage
      title="Excel 预览"
      badge="E"
      badgeClassName="bg-gradient-to-br from-green-500 to-green-700"
<<<<<<< HEAD
      defaultFileName="New_Spreadsheet.xlsx"
=======
      defaultFileName="test.xlsx"
      initialFileUrl="/test.xlsx"
>>>>>>> refactor/v9
      fileType={FILE_TYPE.XLSX}
      accept=".xlsx,.xls,.ods,.csv"
      newButtonLabel="新建 Excel"
    />
  );
}
