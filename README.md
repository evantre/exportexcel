# exportexcel

浏览器端生成和导出 Excel (Browser-side generation and export of Excel)。支持合并单元格和设置单元格样式。

## Import and API

```js
import { exportexcel, workbook, download } from "fe-export-excel";

// 创建并下载 Excel
exportexcel(workbookName, worksheetConfig);
// exportexcel(workbookName, [worksheetConfig1, worksheetConfig2]);

// 创建 Excel
const wb = workbook(worksheetConfig);
// const wb = workbook([worksheetConfig1, worksheetConfig2]);

// 下载 Excel
download(wb, workbookName);
```

## Examples

```js
const headerStyle = {
  alignment: { vertical: "middle", horizontal: "center" },
};

const data = [
  {
    date: "2020年08月26日",
    count1: "1234",
    number1: "1234",
    count2: "1234",
    number2: "1234",
    weekCount: "1234",
    weekAvg: "1234",
    ratio: "1234",
    count: "1234",
    avg: "1234",
    stock: "1",
  },
];

const worksheet = {
  // 一个 excel 是一个 workbook，里面的每个工作表是一个 worksheet
  sheetName: "sheet1",

  // 详细配置参考: https://github.com/exceljs/exceljs
  // properties: {},
  // pageSetup: {},
  // headerFooter: {},
  // views: {},
  // autoFilter: {},
  // columns: {},

  // 可选，统一给单元格设置样式
  // numFmt, font, alignment, border, fill
  cellStyle: {
    // border: {
    //   top: { style: "thin", color: { argb: "00000000" } },
    //   left: { style: "thin", color: { argb: "00000000" } },
    //   bottom: { style: "thin", color: { argb: "00000000" } },
    //   right: { style: "thin", color: { argb: "00000000" } },
    // },
    // font: {
    //   name: "Comic Sans MS",
    //   family: 4,
    //   size: 16,
    //   bold: true,
    // },
    // alignment: { vertical: "middle", horizontal: "center" },
    // fill: {
    //   type: "pattern",
    //   pattern: "darkVertical",
    //   fgColor: { argb: "FFFF0000" },
    // },
    // numFmt: "0.00%",
  },

  rows: [
    // 行内容(不区分表头行和数据行), 
    [
      // 列内容，对象，可以进行相关配置；没有特殊配置时，也可以直接传字符串或者数字
      // 为对象时候必填，value 当前单元格的值，字符串或者数字
      // 可选，start 当前单元格起始列，从 0 开始；如 start 位置和数组索引位置一致可以省略
      // 可选，rowspan 当前单元格跨多少行
      // 可选，colspan 当前单元格跨多少列
      // 可选，style 为当前单元格的样式，配置同 cellStyle
      { start: 0, rowspan: 2, value: "时间", style: headerStyle },
      { start: 1, colspan: 2, value: "生产", style: headerStyle },
      { start: 3, colspan: 2, value: "联合试转运", style: headerStyle },
      { start: 5, colspan: 5, value: "产量", style: headerStyle },
      { start: 10, rowspan: 2, value: "期末库存", style: headerStyle },
    ],
    [
      { start: 1, value: "处数", style: headerStyle },
      { start: 2, value: "产能（万吨/年）", style: headerStyle },
      { start: 3, value: "处数" },
      { start: 4, value: "产能（万吨/年）", style: headerStyle },
      { start: 5, value: "本周累计", style: headerStyle },
      { start: 6, value: "本周日均", style: headerStyle },
      { start: 7, value: "本周日均环比", style: headerStyle },
      { start: 8, value: "1月1日起累计", style: headerStyle },
      { start: 9, value: "1月1日起日均", style: headerStyle },
    ],
    ...data.map((x) => [
      x.date,
      x.count1,
      x.number1,
      x.count2,
      x.number2,
      x.weekCount,
      x.weekAvg,
      x.ratio,
      x.count,
      x.avg,
      x.stock,
    ]),
  ],
};

// 只有一个 worksheet 的时候可以直接传 worksheet
// exportExcel("测试", worksheet);
exportexcel("测试", [worksheet]);
```
