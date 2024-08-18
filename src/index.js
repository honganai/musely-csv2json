const XLSX = require("xlsx");
const fs = require("fs");
XLSX.set_fs(fs);

const workbook = XLSX.readFile("./Generator Tools.xlsx");
const WhiteList = [
  "Generator Tools",
  "Standard Inputs (WIP)",
  "技术设计",
  "Prompts",
];

const sheets = workbook.SheetNames.filter(
  (sheetName) => !WhiteList.includes(sheetName)
).map((sheetName) => {
  const worksheet = workbook.Sheets[sheetName];
  const range = XLSX.utils.decode_range(worksheet["!ref"]);
  const sheet = { content: [] };

  for (let C = 0; C < 6; ++C) {
    const keyOfKey = XLSX.utils.encode_cell({ c: 0, r: C });
    const kOfKey = XLSX.utils.encode_cell({ c: 1, r: C });

    const key = worksheet[keyOfKey]?.v || "";
    const value = worksheet[kOfKey]?.v || "";
    sheet[key] = value;
  }

  for (let R = 6; R <= 100; ++R) {
    const row = {};

    for (let C = range.s.c; C <= range.e.c; ++C) {
      const value = worksheet[XLSX.utils.encode_cell({ c: C, r: R })]?.v || "";
      const key = worksheet[XLSX.utils.encode_cell({ c: C, r: 5 })]?.v || "";
      row[key] = value ? String(value).replace(/\n/g, "") : "";
    }

    if (
      Object.keys(row)
        .map((key) => row[key])
        .every((value) => value.toString().trim().length === 0)
    )
      break;
    sheet.content.push(row);
  }

  return sheet;
});

fs.writeFileSync("sheets.json", JSON.stringify(sheets, null, 2));
