import type { CellValue, Workbook, Worksheet } from "exceljs";

export type SchoolRecord = {
  schoolName: string;
  normalizedName: string;
  schoolCode: string;
  schoolLevel: string;
  establishment: string;
  educationOffice: string;
  postalCode: string;
  address: string;
};

export type SchoolDirectoryData = {
  schemaVersion: number;
  sourceFile: string;
  generatedAt: string;
  count: number;
  schools: SchoolRecord[];
};

export type SchoolAddressRow = {
  schoolName: string;
  postalCode: string;
  address: string;
  matched: boolean;
};

export type SchoolAddressOutput = {
  rows: SchoolAddressRow[];
  matchedCount: number;
  unmatchedNames: string[];
};

export const SCHOOL_ADDRESS_HEADERS = [
  "학교명",
  "우편번호",
  "주소",
  "수신자 이름",
  "주의사항",
] as const;

export function schoolAddressExcelRow(
  row: SchoolAddressRow,
  recipientName: string,
  deliveryNote: string,
) {
  return [
    row.schoolName,
    row.postalCode,
    row.address,
    recipientName.trim(),
    deliveryNote.trim(),
  ];
}

function cellText(value: CellValue) {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value.trim();
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value).trim();
  }
  if (value instanceof Date) return value.toISOString();
  if ("richText" in value) {
    return value.richText.map((part) => part.text).join("").trim();
  }
  if ("hyperlink" in value) return value.text.trim();
  if ("result" in value && value.result !== undefined) {
    return String(value.result ?? "").trim();
  }
  return "";
}

export function normalizeSchoolName(value: CellValue | string) {
  return cellText(value as CellValue).replace(/\s+/g, "");
}

function findHeader(
  worksheet: Worksheet,
  requiredHeaders: string[],
  maxRows = 20,
) {
  for (
    let rowNumber = 1;
    rowNumber <= Math.min(worksheet.rowCount, maxRows);
    rowNumber += 1
  ) {
    const row = worksheet.getRow(rowNumber);
    const headers = new Map<string, number>();
    row.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
      headers.set(cellText(cell.value), columnNumber);
    });
    if (requiredHeaders.every((header) => headers.has(header))) {
      return { rowNumber, headers };
    }
  }
  return null;
}

export function parseSchoolDirectoryWorkbook(
  workbook: Workbook,
  sourceFile = "업로드한 학교 기본현황.xlsx",
): SchoolDirectoryData {
  const required = [
    "시도교육청",
    "교육지원청",
    "학교급",
    "설립구분",
    "학교",
    "학교코드",
    "우편번호",
    "학교도로명주소",
  ];
  const located = workbook.worksheets
    .map((worksheet) => ({
      worksheet,
      header: findHeader(worksheet, required),
    }))
    .find((item) => item.header);
  if (!located?.header) {
    throw new Error(
      "학교, 우편번호, 학교도로명주소가 있는 학교 기본현황 시트를 찾지 못했습니다.",
    );
  }

  const { worksheet, header } = located;
  const schools: SchoolRecord[] = [];
  const seen = new Set<string>();
  for (
    let rowNumber = header.rowNumber + 1;
    rowNumber <= worksheet.rowCount;
    rowNumber += 1
  ) {
    const row = worksheet.getRow(rowNumber);
    const schoolName = cellText(
      row.getCell(header.headers.get("학교")!).value,
    );
    if (!schoolName) continue;
    const normalizedName = normalizeSchoolName(schoolName);
    if (seen.has(normalizedName)) {
      throw new Error(`학교 기본현황에 같은 학교명이 중복되어 있습니다: ${schoolName}`);
    }
    seen.add(normalizedName);
    schools.push({
      schoolName,
      normalizedName,
      schoolCode: cellText(
        row.getCell(header.headers.get("학교코드")!).value,
      ),
      schoolLevel: cellText(
        row.getCell(header.headers.get("학교급")!).value,
      ),
      establishment: cellText(
        row.getCell(header.headers.get("설립구분")!).value,
      ),
      educationOffice: cellText(
        row.getCell(header.headers.get("교육지원청")!).value,
      ),
      postalCode: cellText(
        row.getCell(header.headers.get("우편번호")!).value,
      ).padStart(5, "0"),
      address: cellText(
        row.getCell(header.headers.get("학교도로명주소")!).value,
      ),
    });
  }
  if (schools.length === 0) {
    throw new Error("학교 기본현황에서 학교 자료를 찾지 못했습니다.");
  }

  return {
    schemaVersion: 1,
    sourceFile,
    generatedAt: new Date().toISOString().slice(0, 10),
    count: schools.length,
    schools,
  };
}

function findAwardSchoolColumn(worksheet: Worksheet) {
  const located = findHeader(worksheet, ["sc1"], 10);
  if (!located) return null;
  return {
    headerRow: located.rowNumber,
    schoolColumn: located.headers.get("sc1")!,
  };
}

export function createSchoolAddressOutput(
  awardWorkbook: Workbook,
  directory: SchoolDirectoryData,
): SchoolAddressOutput {
  const located = awardWorkbook.worksheets
    .map((worksheet) => ({
      worksheet,
      awardHeader: findAwardSchoolColumn(worksheet),
    }))
    .find((item) => item.awardHeader);
  if (!located?.awardHeader) {
    throw new Error(
      "상장 명단에서 sc1 열을 찾지 못했습니다. ‘상장 명단 만들기’에서 생성한 파일인지 확인해 주세요.",
    );
  }

  const uniqueNames: string[] = [];
  const seen = new Set<string>();
  for (
    let rowNumber = located.awardHeader.headerRow + 1;
    rowNumber <= located.worksheet.rowCount;
    rowNumber += 1
  ) {
    const schoolName = cellText(
      located.worksheet
        .getRow(rowNumber)
        .getCell(located.awardHeader.schoolColumn).value,
    );
    const normalized = normalizeSchoolName(schoolName);
    if (!normalized || seen.has(normalized)) continue;
    seen.add(normalized);
    uniqueNames.push(schoolName);
  }
  if (uniqueNames.length === 0) {
    throw new Error("상장 명단에서 학교명을 찾지 못했습니다.");
  }

  const directoryMap = new Map(
    directory.schools.map((school) => [school.normalizedName, school]),
  );
  const unmatchedNames: string[] = [];
  const rows = uniqueNames.map((schoolName) => {
    const school = directoryMap.get(normalizeSchoolName(schoolName));
    if (!school) {
      unmatchedNames.push(schoolName);
      return {
        schoolName,
        postalCode: "",
        address: "",
        matched: false,
      };
    }
    return {
      schoolName: school.schoolName,
      postalCode: school.postalCode,
      address: school.address,
      matched: true,
    };
  });

  return {
    rows,
    matchedCount: rows.length - unmatchedNames.length,
    unmatchedNames,
  };
}

export function schoolAddressOutputFileName(inputName: string) {
  return `${inputName.replace(/\.xlsx$/i, "")}_학교주소명단.xlsx`;
}
