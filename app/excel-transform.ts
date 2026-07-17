import type {
  CellValue,
  Comment,
  DataValidation,
  Row,
  Style,
  Worksheet,
} from "exceljs";

export type SheetResult = {
  name: string;
  splitRows: number;
  addedRows: number;
};

type CellSnapshot = {
  value: CellValue;
  style: Partial<Style>;
  dataValidation: DataValidation;
  note: string | Comment;
};

type RowSnapshot = {
  cells: CellSnapshot[];
  height?: number;
  hidden: boolean;
  outlineLevel: number;
};

function cloneValue<T>(value: T): T {
  if (value === null || value === undefined || typeof value !== "object") {
    return value;
  }
  if (value instanceof Date) return new Date(value.getTime()) as T;
  return structuredClone(value);
}

export function columnToNumber(column: string) {
  let value = 0;
  for (const char of column) value = value * 26 + char.charCodeAt(0) - 64;
  return value;
}

export function normalizeColumn(value: string) {
  const column = value.trim().toUpperCase();
  if (!/^[A-Z]{1,3}$/.test(column)) return null;
  const index = columnToNumber(column);
  return index >= 1 && index <= 16384 ? column : null;
}

function cellText(value: CellValue) {
  if (typeof value === "string") return value;
  if (value && typeof value === "object" && "richText" in value) {
    return value.richText.map((part) => part.text).join("");
  }
  if (value && typeof value === "object" && "hyperlink" in value) {
    return value.text;
  }
  return null;
}

function splitLines(value: CellValue) {
  const text = cellText(value);
  if (text === null) return { raw: "", parts: [] as string[] };
  const raw = text.replace(/\r\n?/g, "\n");
  return {
    raw,
    parts: raw
      .split("\n")
      .map((part) => part.trim())
      .filter(Boolean),
  };
}

function expandSingleValue(raw: string, parts: string[], targetCount: number) {
  if (parts.length !== 1 || targetCount <= 1) return parts;

  // Some school files visually separate grades with repeated spaces rather
  // than Alt+Enter. Recognize that narrow, common case without splitting names.
  const spacedGrades = raw
    .trim()
    .split(/\s{2,}/)
    .map((part) => part.trim())
    .filter(Boolean);
  if (
    spacedGrades.length === targetCount &&
    spacedGrades.every((part) => /학년$/.test(part))
  ) {
    return spacedGrades;
  }

  return Array.from({ length: targetCount }, () => parts[0]);
}

function snapshotRow(row: Row, columnCount: number): RowSnapshot {
  const cells: CellSnapshot[] = [];
  for (let column = 1; column <= columnCount; column += 1) {
    const cell = row.getCell(column);
    cells.push({
      value: cloneValue(cell.value),
      style: cloneValue(cell.style),
      dataValidation: cloneValue(cell.dataValidation),
      note: cloneValue(cell.note),
    });
  }
  return {
    cells,
    height: row.height,
    hidden: row.hidden,
    outlineLevel: row.outlineLevel ?? 0,
  };
}

function restoreRow(row: Row, snapshot: RowSnapshot) {
  if (snapshot.height !== undefined) row.height = snapshot.height;
  row.hidden = snapshot.hidden;
  row.outlineLevel = snapshot.outlineLevel;
  snapshot.cells.forEach((source, index) => {
    const cell = row.getCell(index + 1);
    cell.value = cloneValue(source.value);
    cell.style = cloneValue(source.style);
    if (source.dataValidation && Object.keys(source.dataValidation).length > 0) {
      cell.dataValidation = cloneValue(source.dataValidation);
    }
    if (source.note) cell.note = cloneValue(source.note);
  });
}

export function processSheet(
  worksheet: Worksheet,
  columns: number[],
): SheetResult {
  const lastRow = worksheet.rowCount;
  const lastColumn = Math.max(worksheet.columnCount, ...columns);
  let splitRows = 0;
  let addedRows = 0;

  // Work bottom-up so inserted rows never change the source rows still ahead.
  for (let rowNumber = lastRow; rowNumber >= 1; rowNumber -= 1) {
    const sourceRow = worksheet.getRow(rowNumber);
    const splitValues = new Map<number, { raw: string; parts: string[] }>();
    let outputCount = 1;

    columns.forEach((column) => {
      const split = splitLines(sourceRow.getCell(column).value);
      splitValues.set(column, split);
      outputCount = Math.max(outputCount, split.parts.length);
    });
    if (outputCount <= 1) continue;

    columns.forEach((column) => {
      const split = splitValues.get(column);
      if (!split) return;
      splitValues.set(column, {
        raw: split.raw,
        parts: expandSingleValue(split.raw, split.parts, outputCount),
      });
    });

    const snapshot = snapshotRow(sourceRow, lastColumn);
    const blanks = Array.from({ length: outputCount - 1 }, () =>
      Array.from({ length: lastColumn }, () => null),
    );
    worksheet.spliceRows(rowNumber + 1, 0, ...blanks);

    for (let offset = 0; offset < outputCount; offset += 1) {
      const row = worksheet.getRow(rowNumber + offset);
      if (offset > 0) restoreRow(row, snapshot);
      columns.forEach((column) => {
        const cell = row.getCell(column);
        cell.value = splitValues.get(column)?.parts[offset] ?? null;
        cell.alignment = { ...cell.alignment, wrapText: false };
      });
    }

    splitRows += 1;
    addedRows += outputCount - 1;
  }

  return { name: worksheet.name, splitRows, addedRows };
}

export function outputFileName(inputName: string) {
  return `${inputName.replace(/\.xlsx$/i, "")}_행분리.xlsx`;
}
