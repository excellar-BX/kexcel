import { Spreadsheet } from "@/components/ClientOnly";

export default function SheetPage() {
  // The editor reads the sheet id from the URL itself, in the browser.
  return <Spreadsheet />;
}
