import type { Metadata } from "next";
import { OfficePreviewPage } from "@/components/onlyoffice-web-demo/office-preview-page";
import { FILE_TYPE } from "@/components/onlyoffice-web-comp";

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
