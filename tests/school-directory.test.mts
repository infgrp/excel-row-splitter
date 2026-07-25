import assert from "node:assert/strict";
import test from "node:test";
import ExcelJS from "exceljs";
import {
  createSchoolAddressOutput,
  parseSchoolDirectoryWorkbook,
  SCHOOL_ADDRESS_HEADERS,
  schoolAddressExcelRow,
  schoolAddressOutputFileName,
} from "../app/school-directory.ts";

function createDirectoryWorkbook() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("학교현황");
  sheet.addRow(["(나이스)_학교기본현황"]);
  sheet.addRow([
    "시도교육청",
    "교육지원청",
    "학교급",
    "설립구분",
    "학교",
    "학교코드",
    "우편번호",
    "학교도로명주소",
  ]);
  sheet.addRow([
    "전북특별자치도교육청",
    "군산교육지원청",
    "초등학교",
    "공립",
    "군산금빛초등학교",
    "P100002788",
    54072,
    "전북특별자치도 군산시 궁포3로 48",
  ]);
  sheet.addRow([
    "전북특별자치도교육청",
    "군산교육지원청",
    "초등학교",
    "공립",
    "군산풍문초등학교",
    "P100000593",
    54041,
    "전북특별자치도 군산시 풍마길 22",
  ]);
  return workbook;
}

test("parses a separately managed school directory workbook", () => {
  const directory = parseSchoolDirectoryWorkbook(
    createDirectoryWorkbook(),
    "학교기본현황.xlsx",
  );
  assert.equal(directory.count, 2);
  assert.deepEqual(directory.schools[0], {
    schoolName: "군산금빛초등학교",
    normalizedName: "군산금빛초등학교",
    schoolCode: "P100002788",
    schoolLevel: "초등학교",
    establishment: "공립",
    educationOffice: "군산교육지원청",
    postalCode: "54072",
    address: "전북특별자치도 군산시 궁포3로 48",
  });
});

test("extracts unique award schools and keeps unmatched schools visible", () => {
  const directory = parseSchoolDirectoryWorkbook(createDirectoryWorkbook());
  const award = new ExcelJS.Workbook();
  const sheet = award.addWorksheet("상장명단");
  sheet.addRow(["1", "2", "3", "sc1", "s1", "n1"]);
  sheet.addRow([1, "물리", "금상", "군산금빛초등학교", "6학년", "김학생"]);
  sheet.addRow([1, "물리", "금상", "군산금빛초등학교", "6학년", "김학생"]);
  sheet.addRow([2, "화학", "은상", "없는 학교", "5학년", "이학생"]);

  const output = createSchoolAddressOutput(award, directory);
  assert.equal(output.rows.length, 2);
  assert.equal(output.matchedCount, 1);
  assert.deepEqual(output.unmatchedNames, ["없는 학교"]);
  assert.deepEqual(output.rows[0], {
    schoolName: "군산금빛초등학교",
    postalCode: "54072",
    address: "전북특별자치도 군산시 궁포3로 48",
    matched: true,
  });
  assert.deepEqual([...SCHOOL_ADDRESS_HEADERS], [
    "학교명",
    "우편번호",
    "주소",
    "수신자 이름",
    "주의사항",
  ]);
  assert.deepEqual(
    schoolAddressExcelRow(
      output.rows[0],
      " 과학담당 선생님 ",
      " 상장 재중 · 접지 금지 ",
    ),
    [
      "군산금빛초등학교",
      "54072",
      "전북특별자치도 군산시 궁포3로 48",
      "과학담당 선생님",
      "상장 재중 · 접지 금지",
    ],
  );
  assert.equal(
    schoolAddressOutputFileName("상장명단.xlsx"),
    "상장명단_학교주소명단.xlsx",
  );
});

test("rejects a school directory with duplicate names", () => {
  const workbook = createDirectoryWorkbook();
  workbook.getWorksheet("학교현황")!.addRow([
    "전북특별자치도교육청",
    "군산교육지원청",
    "초등학교",
    "공립",
    "군산금빛초등학교",
    "DIFFERENT",
    54072,
    "다른 주소",
  ]);
  assert.throws(
    () => parseSchoolDirectoryWorkbook(workbook),
    /같은 학교명이 중복/,
  );
});
