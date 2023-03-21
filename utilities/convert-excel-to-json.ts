import * as Excel from "exceljs";
import * as path from "path";
import { snakeCase } from "lodash";
import fs from "fs";

let SHEET_LIST = {
  ENGLISH: "en",
  FRENCH: "fr",
  SPANISH: "es",
  GERMAN: "de",
  ITALIAN: "it",
  RUSSIAN: "ru",
  JAPANESE: "ja",
  ARAB: "ar",
  CHINESE: "zh",
  KOREAN: "kr",
  SUOMI: "fi",
  "TIẾNG VIỆT": "vi",
  VIETNAM: "vi",
  INDONESIAN: "id",
  日本語: "ja",
  한국어: "kr",
  FRANÇAIS: "fr",
  ESPAÑOL: "es",
  FILIPINO: "fil",
  हिन्दी: "hi",
  HINDI: "hi",
  TÜRKÇE: "tr",
  فارسی: "fa",
  PORTUGIS: "pt",
  РУCCКИЙ: "ru",
  DEUTSCH: "de",
  ภาษาไทย: "th",
  POLSKI: "pl",
  ITALIANO: "it",
};

let workbook = new Excel.Workbook();
const filePath = path.join(__dirname, process.argv[2]);

if (!fs.existsSync(filePath)) {
  console.log(`File ${filePath} not found on utilities folder!`);
  process.exit(1);
}

let translationJson = {};
let currentPage = "";
workbook.xlsx.readFile(filePath).then(function () {
  workbook.eachSheet(function (worksheet, sheetId) {
    if (!SHEET_LIST[worksheet.name]) {
      console.log(
        `Sheet ${worksheet.name} not found, in the SHEET_LIST variable! Please update and try again!`
      );
      process.exit(1);
    }
    worksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
      if (rowNumber < 3) return;
      // [null,null,"Page","Primary Key","Translation Text"]
      const key = snakeCase(sanitizeCellValue(row.values[3]));

      if (row.values[2] && row.values[2] !== currentPage) {
        if (currentPage) {
          writeJsonFile({
            folder: SHEET_LIST[worksheet.name],
            fileName: snakeCase(sanitizeCellValue(currentPage)) + ".json",
            data: translationJson,
          });
        }
        currentPage = row.values[2];
        translationJson = {};
      } else {
        if (!row.values[3]) return;
        translationJson[key] = sanitizeCellValue(row.values[4]);
      }
    });
  });
});

const writeJsonFile = ({
  folder,
  fileName,
  data,
}: {
  folder: string;
  fileName: string;
  data: any;
}) => {
  const folderPath = path.join(__dirname, "../" + folder);
  if (!fs.existsSync(folderPath)) {
    fs.mkdirSync(folderPath);
  }
  fs.writeFileSync(
    path.join(folderPath, fileName),
    JSON.stringify(data, null, 2)
  );
};

const sanitizeCellValue = (cellValue) => {
  if (typeof cellValue === "string") {
    // do nothing
  } else if (typeof cellValue === "object") {
    if (cellValue.richText) {
      // Rich Text Type
      cellValue = cellValue.richText.map((rt) => rt.text || "").join("");
    } else if (cellValue.result) {
      // Formula Type
      cellValue = cellValue.result;
    }
  }
  return cellValue || "";
};
