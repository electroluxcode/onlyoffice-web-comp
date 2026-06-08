import type { Metadata } from "next";
import { OfficePreviewPage } from "@/components/onlyoffice-web-demo/office-preview-page";
import { FILE_TYPE } from "@/components/onlyoffice-web-comp";

export const metadata: Metadata = {
  title: "Web Office — OnlyOffice MVP",
};

export default function EditorPage() {
  return (
    <OfficePreviewPage
      title="Web Office"
      badge="O"
      badgeClassName="bg-gradient-to-br from-indigo-500 to-indigo-700"
      defaultFileName="New_Document.docx"
      fileType={FILE_TYPE.DOCX}
      accept=".docx,.doc,.odt,.rtf,.txt,.xlsx,.xls,.ods,.csv,.pptx,.ppt,.odp"
      newButtonLabel="新建文档"
    />
  );
}
