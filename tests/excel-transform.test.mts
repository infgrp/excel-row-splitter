import assert from "node:assert/strict";
import test from "node:test";
import ExcelJS from "exceljs";
import { normalizeColumn, processSheet } from "../app/excel-transform.ts";

test("validates Excel column letters", () => {
  assert.equal(normalizeColumn(" d "), "D");
  assert.equal(normalizeColumn("XFD"), "XFD");
  assert.equal(normalizeColumn("XFE"), null);
  assert.equal(normalizeColumn("3"), null);
});

test("splits paired line-break values and copies other cells", () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("자료");
  sheet.addRow(["번호", "분야", "작품명", "학생", "학년", "학교"]);
  sheet.addRow([1, "물리", "작품 A", "김하나\n이두리", "1학년\n2학년", "새봄학교"]);

  const result = processSheet(sheet, [4, 5]);

  assert.deepEqual(result, { name: "자료", splitRows: 1, addedRows: 1 });
  assert.equal(sheet.getCell("D2").value, "김하나");
  assert.equal(sheet.getCell("E2").value, "1학년");
  assert.equal(sheet.getCell("D3").value, "이두리");
  assert.equal(sheet.getCell("E3").value, "2학년");
  assert.equal(sheet.getCell("C3").value, "작품 A");
  assert.equal(sheet.getCell("F3").value, "새봄학교");
});

test("repeats a singleton and leaves a missing paired value blank", () => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("자료");
  sheet.addRow([null, null, null, "한사람", "1학년\n2학년\n3학년"]);
  sheet.addRow([null, null, null, "가\n나", "1학년\n2학년\n3학년"]);

  const result = processSheet(sheet, [4, 5]);

  assert.equal(result.addedRows, 4);
  assert.deepEqual([sheet.getCell("D1").value, sheet.getCell("D2").value, sheet.getCell("D3").value], ["한사람", "한사람", "한사람"]);
  assert.deepEqual([sheet.getCell("D4").value, sheet.getCell("D5").value, sheet.getCell("D6").value], ["가", "나", null]);
});
