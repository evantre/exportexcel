import Excel from "exceljs";

// Excel 相关配置文档参考
// https://github.com/exceljs/exceljs

const defaultProperties = {
  // tabColor: "",
  // outlineLevelCol: "",
  // outlineLevelRow: "",
  // defaultRowHeight: "",
  defaultColWidth: 15,
  // dyDescent: "",
  // rowCount: "",
  // actualRowCount: "",
  // columnCount: "",
  // actualColumnCount: "",
};
const defaultPageSetup = {};
const defaultHeaderFooter = {};

// numFmt, font, alignment, border, fill
const defaultBorder = { style: "thin", color: { argb: "00000000" } };
const defaultCellStyle = {
  border: {
    top: defaultBorder,
    left: defaultBorder,
    bottom: defaultBorder,
    right: defaultBorder,
  },
};

export function workbook(worksheets) {
  if (!Array.isArray(worksheets)) {
    worksheets = [worksheets];
  }

  const workbook = new Excel.Workbook();

  for (let i = 0; i < worksheets.length; i++) {
    addWorksheet(workbook, worksheets[i]);
  }

  return workbook;
}

function addWorksheet(workbook, config) {
  const {
    sheetName,
    properties,
    pageSetup,
    headerFooter,
    views,
    autoFilter,
    columns,
    cellStyle = {},
    rows,
  } = config;

  const sheet = workbook.addWorksheet(sheetName, {
    properties: properties || defaultProperties,
    pageSetup: pageSetup || defaultPageSetup,
    headerFooter: headerFooter || defaultHeaderFooter,
  });

  // 冻结拆分视图
  if (views) {
    sheet.views = views;
  }
  // 自动筛选
  if (autoFilter) {
    sheet.autoFilter = autoFilter;
  }
  // 列配置
  if (columns) {
    sheet.columns = columns;
  }

  // 添加行
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const sheetRow = [];

    for (let j = 0; j < row.length; j++) {
      if (typeof row[j] !== "object") {
        row[j] = { value: row[j] };
      }
      let col = row[j];
      sheetRow[col.start || j] = col.value;
    }

    for (let j = 0; j < sheetRow.length; j++) {
      if (!sheetRow[j]) {
        sheetRow[j] = "";
      }
    }

    sheet.addRow(sheetRow);
  }

  // 设置单元格样式、合并单元格
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const sheetRow = sheet.getRow(i + 1);

    for (let j = 0; j < row.length; j++) {
      let { start = j, rowspan, colspan, style = {} } = row[j];
      const cell = sheetRow.getCell(start + 1);

      // 设置单元格样式
      cell.style = {
        numFmt: style.numFmt || cellStyle.numFmt || defaultCellStyle.numFmt,
        font: { ...defaultCellStyle.font, ...cellStyle.font, ...style.font },
        alignment: {
          ...defaultCellStyle.alignment,
          ...cellStyle.alignment,
          ...style.alignment,
        },
        border: {
          ...defaultCellStyle.border,
          ...cellStyle.border,
          ...style.border,
        },
        fill: { ...defaultCellStyle.fill, ...cellStyle.fill, ...style.fill },
      };

      // 设置合并单元格
      if (rowspan || colspan) {
        rowspan = rowspan || 1;
        colspan = colspan || 1;
        sheet.mergeCells(i + 1, start + 1, i + rowspan, start + colspan);
      }
    }
  }
}
