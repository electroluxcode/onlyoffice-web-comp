import type { Metadata } from "next";
<<<<<<< HEAD
import { OfficePreviewPage } from "@/components/office/office-preview-page";
import { FILE_TYPE } from "@/onlyoffice-comp";
=======
import { OfficePreviewPage } from "@/components/onlyoffice-web-demo/office-preview-page";
import { FILE_TYPE } from "@/components/onlyoffice-web-comp";
>>>>>>> refactor/v9

export const metadata: Metadata = {
  title: "Word Preview — OnlyOffice MVP",
};

export default function WordBasePage() {
  return (
    <OfficePreviewPage
      title="Word 预览"
      badge="W"
      badgeClassName="bg-gradient-to-br from-blue-500 to-blue-700"
      defaultFileName="New_Document.docx"
      fileType={FILE_TYPE.DOCX}
      accept=".docx,.doc,.odt,.rtf,.txt"
      newButtonLabel="新建 Word"
    />
  );
}
