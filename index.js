import { workbook } from "./lib/workbook";
import { download } from "./lib/download";

export { workbook, download };

export function exportexcel(workbookName, worksheets) {
  const wb = workbook(worksheets);
  download(wb, workbookName);
}
