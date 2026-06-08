import type { Metadata } from "next";
import { TabsMultiPage } from "@/components/onlyoffice-web-demo/tabs-multi-page";

export const metadata: Metadata = {
  title: "Tabs Multi Instance — OnlyOffice MVP",
};

export default function MultiTabsPage() {
  return <TabsMultiPage />;
}
