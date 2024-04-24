import * as fs from "fs";
import * as exceljs from "exceljs";

const workbook = new exceljs.Workbook();
const pathName = "/Users/kou12345/Downloads/無題のスプレッドシート.xlsx";

const parsedCellValue = (cell: exceljs.Cell): string => {
  const cellValue = cell.value;
  // merged cell
  if (cell.isMerged && cell.master !== cell) {
    return "";
  }
  // null
  if (cellValue === null) {
    return "";
  }
  // number
  if (typeof cellValue === "number") {
    return cellValue.toString();
  }
  // string
  if (typeof cellValue === "string") {
    return cellValue;
  }
  // Date
  if (typeof cellValue === "object" && cellValue instanceof Date) {
    return cellValue.toISOString();
  }
  // exceljs.CellHyperlinkValue
  if (
    typeof cellValue === "object" &&
    cellValue !== null &&
    "hyperlink" in cellValue
  ) {
    return cellValue.text;
  }
  // exceljs.CellFormulaValue
  if (
    typeof cellValue === "object" &&
    cellValue !== null &&
    "formula" in cellValue
  ) {
    return cellValue.result?.toString() || "";
  }
  // exceljs.CellSharedFormulaValue
  if (
    typeof cellValue === "object" &&
    cellValue !== null &&
    "sharedFormula" in cellValue
  ) {
    return cellValue.result?.toString() || "";
  }
  // Array Formula
  if (
    typeof cellValue === "object" &&
    cellValue !== null &&
    "result" in cellValue
  ) {
    return cellValue.result?.toString() || "";
  }
  // exceljs.CellRichTextValue
  if (
    typeof cellValue === "object" &&
    cellValue !== null &&
    "richText" in cellValue
  ) {
    return cellValue.richText.map((rt) => rt.text).join("");
  }
  // boolean
  if (typeof cellValue === "boolean") {
    return cellValue.toString();
  }
  // exceljs.CellErrorValue
  if (
    typeof cellValue === "object" &&
    cellValue !== null &&
    "error" in cellValue
  ) {
    return cellValue.error;
  }
  return "";
};

const getParsedSheetData = (
  workbook: exceljs.Workbook,
  sheetName: string
): string => {
  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) {
    throw new Error("Worksheet not found");
  }
  const sheetData: string[][] = [];

  worksheet.eachRow((row) => {
    const rowData: string[] = [];
    row.eachCell((cell) => {
      rowData.push(parsedCellValue(cell));
    });
    sheetData.push(rowData);
  });

  const output = sheetData
    .map((row) => {
      return row.join(" ") + "\n";
    })
    .join("");

  return output;
};

const highlightCellsWithKeyword = (
  workbook: exceljs.Workbook,
  sheetName: string,
  keywords: string[]
): void => {
  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) {
    throw new Error("Worksheet not found");
  }

  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      const cellValue = cell.value?.toString();
      console.log(typeof cellValue, cellValue);

      // console.log(keywords);
      if (keywords.some((keyword) => cellValue === keyword)) {
        /*
        https://github.com/exceljs/exceljs/issues/2055#issuecomment-1436262550
        lib は Excel からスタイルを読み取り、影響を受けるセルは Excel ファイル内の 1 つのスタイル オブジェクトを共有します。
        そのため、スタイル プロパティを更新すると、他のセルに影響します。
        解決策はクローンスタイルを作成し、再度セルに割り当てることです。
        */
        cell.style = {
          ...(cell.style || {}), // 既存のスタイルを引き継ぐ
          fill: {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "00FF00" },
          },
        };
      }
    });
  });
};

(async () => {
  await workbook.xlsx.readFile(pathName);
  const worksheet = workbook.getWorksheet("シート1");
  if (!worksheet) {
    throw new Error("Worksheet not found");
  }

  // console.log(getParsedSheetData(workbook, "シート1"));

  // ! なぜか、数字を入力すると全ての文字が入力されているセルの背景色が変わる
  // 漢字の場合は文字列のセルが対象になる
  // highlightCellsWithKeyword(workbook, "シート1", ["4"]);

  console.log(getParsedSheetData(workbook, "シート1"));

  // console.log(worksheet.rowCount);
  // console.log(worksheet.columnCount);

  // // worksheet.rowCount, worksheet.columnCountを元に、全てのセルを取得する
  // const allCells = [];
  // for (let i = 1; i <= worksheet.rowCount; i++) {
  //   const row = worksheet.getRow(i);
  //   for (let j = 1; j <= worksheet.columnCount; j++) {
  //     allCells.push(row.getCell(j));
  //   }
  // }

  // // 全てのセルの値を取得する
  // const allValues = allCells.map((cell) => cell.value);
  // console.log(allValues);

  // // 全てのセルの背景色を取得する
  // const allFills = allCells.map((cell) => cell.fill);
  // console.log(allFills);

  // // allValuesをjson形式に変換する
  // const json: { [key: number]: any }[] = [];
  // for (let i = 0; i < allValues.length; i += worksheet.columnCount) {
  //   const row: { [key: number]: any } = {};
  //   for (let j = 0; j < worksheet.columnCount; j++) {
  //     row[j] = allValues[i + j];
  //   }
  //   json.push(row);
  // }

  // console.log(json);

  // console.log(JSON.stringify(json));

  // jsonをoutput.jsonに書き込む
  // fs.writeFileSync("output.json", JSON.stringify(json));

  // console.log(worksheet.getCell("A1").value);

  // const range = worksheet.getRows(1, 10);
  // if (!range) {
  //   throw new Error("Range not found");
  // }
  // // rangeをみやすく表示する
  // range.forEach((row) => {
  //   console.log(row.values);
  // });

  // const row = worksheet.getRow(1);
  // console.log(row.getCell(1).value);

  // // 背景色を取得する
  // const A1 = worksheet.getCell("A1");
  // console.log(A1.fill);

  // const B2 = worksheet.getCell("B2");
  // console.log(B2.fill);

  // const B3 = worksheet.getCell("B3");
  // console.log(B3.fill);
  // // B3の背景色を変更する
  // ! B3以外のセルの背景色も変わってしまう
  // const cell = worksheet.getCell("B3");
  // cell.style = {
  //   ...(cell.style || {}),
  //   fill: {
  //     type: "pattern",
  //     pattern: "solid",
  //     fgColor: { argb: "00FF00" },
  //   },
  // };
  // B3.value = "変更後";

  // worksheet.getCell("B5").value = "10";
  // worksheet.getCell("C5").value = "2";

  // // 書き込み
  // await workbook.xlsx.writeFile(pathName);
  // console.log("Done");
})();
