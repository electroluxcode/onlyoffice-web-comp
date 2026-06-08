import type { Metadata } from "next";
import { OfficePreviewPage } from "@/components/onlyoffice-web-demo/office-preview-page";
import { FILE_TYPE } from "@/components/onlyoffice-web-comp";

export const metadata: Metadata = {
  title: "PowerPoint Preview — OnlyOffice MVP",
};

export default function PptBasePage() {
  return (
    <OfficePreviewPage
      title="PowerPoint 预览"
      badge="P"
      badgeClassName="bg-gradient-to-br from-orange-500 to-orange-700"
      defaultFileName="New_Presentation.pptx"
      fileType={FILE_TYPE.PPTX}
      accept=".pptx,.ppt,.odp"
      newButtonLabel="新建 PowerPoint"
    />
  );
}
