import type { Metadata } from "next";
<<<<<<< HEAD
import { ProgressiveMultiPage } from "@/components/office/progressive-multi-page";
=======
import { ProgressiveMultiPage } from "@/components/onlyoffice-web-demo/progressive-multi-page";
>>>>>>> refactor/v9

export const metadata: Metadata = {
  title: "Progressive Multi Instance — OnlyOffice MVP",
};

export default function MultiBasePage() {
  return <ProgressiveMultiPage />;
}
