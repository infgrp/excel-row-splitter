import type { CellValue, Worksheet } from "exceljs";

export type AwardOutput = {
  headers: string[];
  rows: Array<Array<CellValue | null>>;
  sourceRows: number;
  recipientRows: number;
  maxMembers: number;
};

function cellText(value: CellValue) {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (value instanceof Date) return value.toISOString();
  if ("richText" in value) {
    return value.richText.map((part) => part.text).join("");
  }
  if ("hyperlink" in value) return value.text;
  if ("result" in value && value.result !== undefined) {
    return String(value.result ?? "");
  }
  return "";
}

function splitLines(value: CellValue) {
  return cellText(value)
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .map((part) => part.trim())
    .filter(Boolean);
}

export function extractAwardCategory(value: CellValue) {
  const text = cellText(value).trim();
  const matches = [...text.matchAll(/\(([^()]*)\)/g)];
  const category = matches.at(-1)?.[1]?.trim() ?? text;
  return category.replace(/부문\s*$/, "").replace(/\s+/g, "");
}

function validateHeaders(worksheet: Worksheet) {
  const expected = new Map([
    [1, "호수"],
    [5, "공적내용"],
    [6, "소속"],
    [8, "학년"],
    [9, "성명"],
    [10, "비고"],
  ]);
  const missing = [...expected.entries()]
    .filter(([column, label]) => !cellText(worksheet.getRow(1).getCell(column).value).includes(label))
    .map(([, label]) => label);
  if (missing.length > 0) {
    throw new Error(
      `선택한 시트의 첫 행에서 ${missing.join(", ")} 열을 확인하지 못했습니다.`,
    );
  }
}

export function createAwardOutput(worksheet: Worksheet): AwardOutput {
  validateHeaders(worksheet);
  const groups: Array<{
    serial: CellValue;
    category: string;
    award: CellValue;
    schools: string[];
    grades: string[];
    names: string[];
  }> = [];

  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const names = splitLines(row.getCell(9).value);
    if (names.length === 0) continue;
    let schools = splitLines(row.getCell(6).value);
    if (schools.length === 1 && names.length > 1) {
      schools = Array.from({ length: names.length }, () => schools[0]);
    }
    if (schools.length !== names.length) {
      throw new Error(
        `${rowNumber}행은 소속 수(${schools.length})가 학생 수(${names.length})와 맞지 않습니다. ` +
          "소속이 하나이면 모든 학생에게 반복하고, 둘 이상이면 학생 순서에 맞게 입력해 주세요.",
      );
    }
    let grades = splitLines(row.getCell(8).value);
    if (grades.length === 1 && names.length > 1) {
      grades = Array.from({ length: names.length }, () => grades[0]);
    }
    if (grades.length > names.length) {
      throw new Error(
        `${rowNumber}행은 학년 수(${grades.length})가 학생 수(${names.length})보다 많습니다.`,
      );
    }
    groups.push({
      serial: row.getCell(1).value,
      category: extractAwardCategory(row.getCell(5).value),
      award: row.getCell(10).value,
      schools,
      grades,
      names,
    });
  }

  if (groups.length === 0) {
    throw new Error("선택한 시트에서 학생 명단을 찾지 못했습니다.");
  }

  const maxMembers = Math.max(3, ...groups.map((group) => group.names.length));
  const headers = [
    "1",
    "2",
    "3",
    "sc1",
    ...Array.from({ length: maxMembers }, (_, index) => `s${index + 1}`),
    ...Array.from({ length: maxMembers }, (_, index) => `n${index + 1}`),
  ];
  const rows: Array<Array<CellValue | null>> = [];

  groups.forEach((group) => {
    for (let index = 0; index < group.names.length; index += 1) {
      rows.push([
        group.serial,
        group.category,
        group.award,
        group.schools[index],
        ...Array.from(
          { length: maxMembers },
          (_, memberIndex) => group.grades[memberIndex] ?? null,
        ),
        ...Array.from(
          { length: maxMembers },
          (_, memberIndex) => group.names[memberIndex] ?? null,
        ),
      ]);
    }
  });

  return {
    headers,
    rows,
    sourceRows: groups.length,
    recipientRows: rows.length,
    maxMembers,
  };
}

export function awardOutputFileName(inputName: string, sheetName: string) {
  const safeSheet = sheetName.replace(/[\\/:*?"<>|]/g, "_");
  return `${inputName.replace(/\.xlsx$/i, "")}_${safeSheet}_상장명단.xlsx`;
}
