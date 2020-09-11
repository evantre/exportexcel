import workbook from "./workbook";
import { saveAs } from "file-saver";

export default function exportExcel(workbookName, worksheets) {
  if (!Array.isArray(worksheets)) {
    worksheets = [worksheets];
  }

  const excel = workbook(worksheets);

  // 下载 download
  excel.xlsx.writeBuffer().then((buffer) => {
    saveAs(new Blob([buffer]), `${workbookName}.xlsx`);
  });
}
