import assert from "node:assert/strict";
import test from "node:test";
import { buildLabelPrintHtml, LABEL_SPEC } from "../app/label-print.ts";

test("builds CL425 2-by-5 label sheets with requested typography", () => {
  const rows = Array.from({ length: 11 }, (_, index) => ({
    schoolName: `학교${index + 1}`,
    postalCode: `5400${index}`,
    address: `전북특별자치도 군산시 긴 주소 ${index + 1}`,
    matched: true,
  }));
  const html = buildLabelPrintHtml(
    rows,
    "과학담당 선생님",
    "상장 재중 <접지 금지>",
  );

  assert.equal(LABEL_SPEC.code, "CL425");
  assert.equal((html.match(/class="label-sheet"/g) ?? []).length, 2);
  assert.equal((html.match(/class="label"/g) ?? []).length, 11);
  assert.match(html, /grid-template-columns: repeat\(2, 84\.5mm\)/);
  assert.match(html, /grid-auto-rows: 53\.5mm/);
  assert.match(html, /\.address[\s\S]*font-size: 12pt/);
  assert.match(html, /\.recipient[\s\S]*font-size: 14pt/);
  assert.match(html, /word-break: keep-all/);
  assert.match(html, /상장 재중 &lt;접지 금지&gt;/);
});
