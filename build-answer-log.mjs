import fs from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = fileURLToPath(new URL(".", import.meta.url));
const workbook = Workbook.create();
const sheet = workbook.worksheets.add("作答紀錄");
sheet.showGridLines = false;
sheet.tabColor = "#2F6D5C";

sheet.getRange("A1:M1").merge();
sheet.getRange("A1").values = [["空污彈弓大作戰：作答紀錄"]];
sheet.getRange("A1:M1").format = {
  font: { name: "Arial", size: 16, bold: true, color: "#17313A" },
  verticalAlignment: "center",
};
sheet.getRange("A1:M1").format.rowHeight = 30;

sheet.getRange("A2:F4").values = [
  ["作答筆數", null, "答對筆數", null, "整體正確率", null],
  [null, null, null, null, null, null],
  ["成功場次", null, "失敗場次", null, "資料用途", "遊戲作答紀錄"],
];
sheet.getRange("B2").formulas = [["=COUNTA(A7:A1006)"]];
sheet.getRange("D2").formulas = [["=COUNTIF(I7:I1006,\"是\")"]];
sheet.getRange("F2").formulas = [["=IF(B2=0,\"\",D2/B2)"]];
sheet.getRange("B4").formulas = [["=COUNTIF(L7:L1006,\"成功\")"]];
sheet.getRange("D4").formulas = [["=COUNTIF(L7:L1006,\"失敗\")"]];
sheet.getRange("A2:F4").format.font = { name: "Arial", size: 10, color: "#23383E" };
sheet.getRange("A2:A4").format.font = { name: "Arial", size: 10, bold: true, color: "#2F6D5C" };
sheet.getRange("C2:C4").format.font = { name: "Arial", size: 10, bold: true, color: "#2F6D5C" };
sheet.getRange("E2:E4").format.font = { name: "Arial", size: 10, bold: true, color: "#2F6D5C" };
sheet.getRange("F2").format.numberFormat = "0.0%";

const headers = [[
  "提交時間",
  "遊戲場次 ID",
  "姓名",
  "題號",
  "題目",
  "選擇選項",
  "選擇內容",
  "正確答案",
  "是否答對",
  "本題得分",
  "剩餘愛心",
  "遊戲結果",
  "裝置類型",
]];
sheet.getRange("A6:M6").values = headers;
sheet.getRange("A6:M6").format = {
  fill: "#2F6D5C",
  font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: true,
  borders: { preset: "inside", style: "thin", color: "#FFFFFF" },
};
sheet.getRange("A6:M6").format.rowHeight = 34;

sheet.getRange("A7:M1006").format = {
  font: { name: "Arial", size: 10, color: "#26373B" },
  verticalAlignment: "center",
};
sheet.getRange("A7:A1006").format.numberFormat = "yyyy-mm-dd hh:mm:ss";
sheet.getRange("D7:D1006").format.numberFormat = "0";
sheet.getRange("J7:K1006").format.numberFormat = "0";
sheet.getRange("D7:D1006").format.horizontalAlignment = "center";
sheet.getRange("F7:F1006").format.horizontalAlignment = "center";
sheet.getRange("H7:M1006").format.horizontalAlignment = "center";

sheet.getRange("I7:I1006").conditionalFormats.add("containsText", {
  text: "是",
  format: { fill: "#E4F3E7", font: { color: "#24643B", bold: true } },
});
sheet.getRange("I7:I1006").conditionalFormats.add("containsText", {
  text: "否",
  format: { fill: "#FCE7E5", font: { color: "#A52E2A", bold: true } },
});
sheet.getRange("L7:L1006").conditionalFormats.add("containsText", {
  text: "失敗",
  format: { fill: "#FCE7E5", font: { color: "#A52E2A", bold: true } },
});

sheet.freezePanes.freezeRows(6);
const widths = {
  A: 145, B: 150, C: 105, D: 55, E: 300, F: 80, G: 190,
  H: 90, I: 80, J: 80, K: 80, L: 90, M: 90,
};
for (const [col, width] of Object.entries(widths)) {
  sheet.getRange(`${col}:${col}`).format.columnWidthPx = width;
}
sheet.getRange("E7:G1006").format.wrapText = true;

sheet.getRange("H4:M4").merge();
sheet.getRange("H4").values = [["網站每次擊中答案時新增一列；姓名只用於本次遊戲識別，不記錄電子郵件或其他個資。"]];
sheet.getRange("H4:M4").format = {
  font: { name: "Arial", size: 9, italic: true, color: "#5F6F72" },
  verticalAlignment: "center",
  wrapText: true,
};

workbook.recalculate();

const inspection = await workbook.inspect({
  kind: "region",
  sheetId: sheet.name,
  range: "A1:M8",
  maxChars: 6000,
});
await fs.writeFile(`${outputDir}/作答紀錄_檢查.json`, JSON.stringify(inspection, null, 2));

const preview = await workbook.render({
  sheetName: "作答紀錄",
  range: "A1:M14",
  scale: 1,
  format: "png",
});
await fs.writeFile(`${outputDir}/作答紀錄_預覽.png`, new Uint8Array(await preview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(`${outputDir}/空污彈弓大作戰_作答紀錄.xlsx`);
