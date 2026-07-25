import type { SchoolAddressRow } from "./school-directory";

export const LABEL_SPEC = {
  code: "CL425",
  columns: 2,
  rows: 5,
  labelsPerSheet: 10,
  labelWidthMm: 84.5,
  labelHeightMm: 53.5,
  leftMarginMm: 16,
  topMarginMm: 8,
  columnGapMm: 9,
  rowGapMm: 2.6,
} as const;

function escapeHtml(value: string) {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function chunk<T>(items: T[], size: number) {
  const chunks: T[][] = [];
  for (let index = 0; index < items.length; index += size) {
    chunks.push(items.slice(index, index + size));
  }
  return chunks;
}

export function buildLabelPrintHtml(
  rows: SchoolAddressRow[],
  recipientName: string,
  deliveryNote: string,
) {
  const sheets = chunk(rows, LABEL_SPEC.labelsPerSheet);
  const recipient = escapeHtml(recipientName.trim());
  const note = escapeHtml(deliveryNote.trim());
  const pages = sheets
    .map((sheetRows) => {
      const labels = sheetRows
        .map(
          (row) => `
          <article class="label">
            <div class="notice">${note || "&nbsp;"}</div>
            <div class="postal">${escapeHtml(row.postalCode)}</div>
            <div class="address">${escapeHtml(row.address)}</div>
            <div class="recipient">${recipient || "&nbsp;"}</div>
          </article>`,
        )
        .join("");
      return `<section class="label-sheet">${labels}</section>`;
    })
    .join("");

  return `<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <title>수상학교 주소 라벨 (${LABEL_SPEC.code})</title>
  <style>
    @page { size: A4 portrait; margin: 0; }
    * { box-sizing: border-box; }
    html, body { margin: 0; padding: 0; background: white; color: #000; }
    body { font-family: "Malgun Gothic", "Apple SD Gothic Neo", sans-serif; }
    .label-sheet {
      width: 210mm;
      height: 297mm;
      padding-top: ${LABEL_SPEC.topMarginMm}mm;
      padding-left: ${LABEL_SPEC.leftMarginMm}mm;
      display: grid;
      grid-template-columns: repeat(${LABEL_SPEC.columns}, ${LABEL_SPEC.labelWidthMm}mm);
      grid-auto-rows: ${LABEL_SPEC.labelHeightMm}mm;
      column-gap: ${LABEL_SPEC.columnGapMm}mm;
      row-gap: ${LABEL_SPEC.rowGapMm}mm;
      align-content: start;
      page-break-after: always;
      break-after: page;
    }
    .label-sheet:last-child { page-break-after: auto; break-after: auto; }
    .label {
      width: ${LABEL_SPEC.labelWidthMm}mm;
      height: ${LABEL_SPEC.labelHeightMm}mm;
      padding: 4.2mm 5mm 3.8mm;
      display: flex;
      flex-direction: column;
      overflow: hidden;
    }
    .notice {
      min-height: 6mm;
      text-align: center;
      font-size: 14pt;
      line-height: 1.15;
      font-weight: 700;
      white-space: nowrap;
    }
    .postal {
      margin-top: 1.2mm;
      min-height: 6mm;
      text-align: left;
      font-size: 14pt;
      line-height: 1.2;
    }
    .address {
      margin-top: 0.8mm;
      text-align: right;
      font-size: 12pt;
      line-height: 1.35;
      word-break: keep-all;
      overflow-wrap: normal;
      white-space: normal;
    }
    .recipient {
      margin-top: auto;
      min-height: 6mm;
      text-align: right;
      font-size: 14pt;
      line-height: 1.2;
      white-space: nowrap;
    }
    @media screen {
      body { background: #d8ddd9; }
      .label-sheet { margin: 8mm auto; background: white; box-shadow: 0 2mm 7mm rgba(0,0,0,.2); }
      .label { outline: .2mm dashed #c6ccc8; }
    }
    @media print {
      .label { outline: none; }
    }
  </style>
</head>
<body>
  ${pages}
  <script>
    window.addEventListener("load", function () {
      window.setTimeout(function () { window.print(); }, 250);
    });
  </script>
</body>
</html>`;
}
