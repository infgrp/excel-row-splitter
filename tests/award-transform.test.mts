import assert from "node:assert/strict";
import test from "node:test";
import ExcelJS from "exceljs";
import {
  awardOutputFileName,
  createAwardOutput,
  extractAwardCategory,
} from "../app/award-transform.ts";

test("extracts and normalizes the final parenthesized category", () => {
  assert.equal(
    extractAwardCategory(
      "제72회(2026년) 전북특별자치도과학전람회 본선대회(산업 및 에너지부문)",
    ),
    "산업및에너지",
  );
  assert.equal(extractAwardCategory("행사(설명)(생물부문)"), "생물");
});

test("creates mail-merge rows and repeats each group by student count", () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("학생");
  sheet.addRow([
    "호수",
    "포상년",
    "월일",
    "포상종별",
    "공적내용",
    "소속(주소)",
    "직위",
    "학년",
    "성명",
    "비고\n(상종및등급)",
  ]);
  sheet.addRow([
    3218,
    2026,
    "07.16.",
    "교육감상",
    "본선대회(물리부문)",
    "군산중앙초등학교\n전주서곡초등학교",
    "학생",
    "5학년\n4학년",
    "김하준\n한지우",
    "동상",
  ]);

  const output = createAwardOutput(sheet);

  assert.deepEqual(output.headers, [
    "1",
    "2",
    "3",
    "sc1",
    "s1",
    "s2",
    "s3",
    "n1",
    "n2",
    "n3",
  ]);
  assert.equal(output.sourceRows, 1);
  assert.equal(output.recipientRows, 2);
  assert.deepEqual(output.rows[0], [
    3218,
    "물리",
    "동상",
    "군산중앙초등학교",
    "5학년",
    "4학년",
    null,
    "김하준",
    "한지우",
    null,
  ]);
  assert.deepEqual(output.rows[1], [
    3218,
    "물리",
    "동상",
    "전주서곡초등학교",
    "5학년",
    "4학년",
    null,
    "김하준",
    "한지우",
    null,
  ]);
});

test("repeats one school and one grade for multiple students and sanitizes the filename", () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("학생");
  sheet.addRow([
    "호수",
    null,
    null,
    null,
    "공적내용",
    "소속",
    null,
    "학년",
    "성명",
    "비고",
  ]);
  sheet.addRow([
    1,
    null,
    null,
    null,
    "행사(화학부문)",
    "학교",
    null,
    "2학년",
    "가\n나",
    "은상",
  ]);
  const output = createAwardOutput(sheet);
  assert.deepEqual(output.rows.map((row) => row[3]), ["학교", "학교"]);
  assert.deepEqual(output.rows[0].slice(4, 7), ["2학년", "2학년", null]);
  assert.equal(
    awardOutputFileName("원본.xlsx", "학생/학교순"),
    "원본_학생_학교순_상장명단.xlsx",
  );
});

test("rejects an ambiguous school count", () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("학생");
  sheet.addRow([
    "호수",
    null,
    null,
    null,
    "공적내용",
    "소속",
    null,
    "학년",
    "성명",
    "비고",
  ]);
  sheet.addRow([
    1,
    null,
    null,
    null,
    "행사(물리부문)",
    "학교1\n학교2",
    null,
    "1학년",
    "가\n나\n다",
    "금상",
  ]);
  assert.throws(
    () => createAwardOutput(sheet),
    /소속 수\(2\)가 학생 수\(3\)와 맞지 않습니다/,
  );
});
