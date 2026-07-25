import fs from "node:fs/promises";
import ExcelJS from "exceljs";

const input = new URL("../data/school-master.xlsx", import.meta.url);
const output = new URL("../public/data/schools.json", import.meta.url);

function text(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function normalizeSchoolName(value) {
  return text(value).replace(/\s+/g, "");
}

const workbook = new ExcelJS.Workbook();
await workbook.xlsx.readFile(input.pathname);
let selected;
for (const worksheet of workbook.worksheets) {
  for (let rowNumber = 1; rowNumber <= Math.min(worksheet.rowCount, 20); rowNumber += 1) {
    const headers = new Map();
    worksheet.getRow(rowNumber).eachCell((cell, columnNumber) => {
      headers.set(text(cell.value), columnNumber);
    });
    if (
      ["학교", "학교코드", "학교급", "설립구분", "교육지원청", "우편번호", "학교도로명주소"]
        .every((header) => headers.has(header))
    ) {
      selected = { worksheet, rowNumber, headers };
      break;
    }
  }
  if (selected) break;
}
if (!selected) throw new Error("학교 기본현황 머리글을 찾지 못했습니다.");

const schools = [];
const seen = new Set();
for (
  let rowNumber = selected.rowNumber + 1;
  rowNumber <= selected.worksheet.rowCount;
  rowNumber += 1
) {
  const row = selected.worksheet.getRow(rowNumber);
  const schoolName = text(row.getCell(selected.headers.get("학교")).value);
  if (!schoolName) continue;
  const normalizedName = normalizeSchoolName(schoolName);
  if (seen.has(normalizedName)) {
    throw new Error(`학교명이 중복되어 있습니다: ${schoolName}`);
  }
  seen.add(normalizedName);
  schools.push({
    schoolName,
    normalizedName,
    schoolCode: text(row.getCell(selected.headers.get("학교코드")).value),
    schoolLevel: text(row.getCell(selected.headers.get("학교급")).value),
    establishment: text(row.getCell(selected.headers.get("설립구분")).value),
    educationOffice: text(row.getCell(selected.headers.get("교육지원청")).value),
    postalCode: text(row.getCell(selected.headers.get("우편번호")).value).padStart(5, "0"),
    address: text(row.getCell(selected.headers.get("학교도로명주소")).value),
  });
}

await fs.mkdir(new URL(".", output), { recursive: true });
await fs.writeFile(
  output,
  JSON.stringify(
    {
      schemaVersion: 1,
      sourceFile: "data/school-master.xlsx",
      generatedAt: new Date().toISOString().slice(0, 10),
      count: schools.length,
      schools,
    },
    null,
    2,
  ),
);
console.log(`학교 ${schools.length}개를 ${output.pathname}에 저장했습니다.`);
