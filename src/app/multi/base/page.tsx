import type { Metadata } from "next";
import { ProgressiveMultiPage } from "@/components/office/progressive-multi-page";

export const metadata: Metadata = {
  title: "Progressive Multi Instance — OnlyOffice MVP",
};

export default function MultiBasePage() {
  return <ProgressiveMultiPage />;
}
