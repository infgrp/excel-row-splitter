"use client";

import type { Worksheet } from "exceljs";
import {
  AlertCircle,
  Award,
  Building2,
  Check,
  CheckCircle2,
  Download,
  FileSpreadsheet,
  ListTree,
  LoaderCircle,
  Plus,
  RotateCcw,
  ShieldCheck,
  UploadCloud,
  Upload,
  X,
} from "lucide-react";
import {
  type ChangeEvent,
  type DragEvent,
  type FormEvent,
  type KeyboardEvent,
  useEffect,
  useRef,
  useState,
} from "react";
import {
  awardOutputFileName,
  createAwardOutput,
} from "./award-transform";
import {
  columnToNumber,
  normalizeColumn,
  outputFileName,
  processSheet,
  type SheetResult,
} from "./excel-transform";
import {
  createSchoolAddressOutput,
  parseSchoolDirectoryWorkbook,
  SCHOOL_ADDRESS_HEADERS,
  schoolAddressExcelRow,
  schoolAddressOutputFileName,
  type SchoolDirectoryData,
} from "./school-directory";
import { buildLabelPrintHtml, LABEL_SPEC } from "./label-print";

const MAX_FILE_SIZE = 50 * 1024 * 1024;
const EXCEL_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

type Mode = "split" | "award" | "address";
type Stage = "idle" | "reading" | "ready" | "processing" | "done" | "error";

type SplitResult = {
  kind: "split";
  sheets: SheetResult[];
  splitRows: number;
  addedRows: number;
  outputName: string;
};

type AwardResult = {
  kind: "award";
  sourceSheet: string;
  sourceRows: number;
  recipientRows: number;
  maxMembers: number;
  outputName: string;
};

type AddressResult = {
  kind: "address";
  schoolCount: number;
  matchedCount: number;
  unmatchedNames: string[];
  directorySource: string;
  outputName: string;
};

type ProcessResult = SplitResult | AwardResult | AddressResult;

function formatBytes(bytes: number) {
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function triggerDownload(url: string, name: string) {
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = name;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
}

function safeWorksheetName(name: string) {
  return `상장명단_${name}`.replace(/[\\/:*?"<>|]/g, "_").slice(0, 31);
}

export default function Home() {
  const [mode, setMode] = useState<Mode>("split");
  const [file, setFile] = useState<File | null>(null);
  const [columns, setColumns] = useState(["D", "E"]);
  const [applyAllSheets, setApplyAllSheets] = useState(true);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [stage, setStage] = useState<Stage>("idle");
  const [isDragging, setIsDragging] = useState(false);
  const [showColumnInput, setShowColumnInput] = useState(false);
  const [newColumn, setNewColumn] = useState("");
  const [columnError, setColumnError] = useState("");
  const [error, setError] = useState("");
  const [result, setResult] = useState<ProcessResult | null>(null);
  const [downloadUrl, setDownloadUrl] = useState("");
  const [directory, setDirectory] = useState<SchoolDirectoryData | null>(null);
  const [directorySource, setDirectorySource] = useState("기본 학교 데이터 불러오는 중");
  const [directoryLoading, setDirectoryLoading] = useState(true);
  const [recipientName, setRecipientName] = useState("");
  const [deliveryNote, setDeliveryNote] = useState("");
  const [labelRows, setLabelRows] = useState<
    import("./school-directory").SchoolAddressRow[]
  >([]);
  const inputRef = useRef<HTMLInputElement>(null);
  const directoryInputRef = useRef<HTMLInputElement>(null);
  const bufferRef = useRef<ArrayBuffer | null>(null);

  useEffect(() => {
    let active = true;
    void fetch("./data/schools.json")
      .then((response) => {
        if (!response.ok) throw new Error("학교 데이터 응답 오류");
        return response.json() as Promise<SchoolDirectoryData>;
      })
      .then((data) => {
        if (!active) return;
        if (!Array.isArray(data.schools) || data.schools.length === 0) {
          throw new Error("학교 데이터 없음");
        }
        setDirectory(data);
        setDirectorySource(
          `기본 데이터 · ${data.count.toLocaleString()}개교 · ${data.generatedAt}`,
        );
      })
      .catch(() => {
        if (!active) return;
        setDirectory(null);
        setDirectorySource("기본 학교 데이터를 불러오지 못했습니다");
      })
      .finally(() => {
        if (active) setDirectoryLoading(false);
      });
    return () => {
      active = false;
    };
  }, []);

  useEffect(() => {
    return () => {
      if (downloadUrl) URL.revokeObjectURL(downloadUrl);
    };
  }, [downloadUrl]);

  const clearOutput = () => {
    if (downloadUrl) URL.revokeObjectURL(downloadUrl);
    setDownloadUrl("");
    setResult(null);
    setLabelRows([]);
  };

  const resetFile = () => {
    clearOutput();
    bufferRef.current = null;
    setFile(null);
    setSheetNames([]);
    setSelectedSheet("");
    setError("");
    setStage("idle");
  };

  const changeMode = (nextMode: Mode) => {
    if (nextMode === mode) return;
    resetFile();
    setMode(nextMode);
  };

  const inspectFile = async (nextFile: File) => {
    clearOutput();
    setError("");
    if (!/\.xlsx$/i.test(nextFile.name)) {
      setStage("error");
      setError(".xlsx 형식의 엑셀 파일만 사용할 수 있습니다.");
      return;
    }
    if (nextFile.size > MAX_FILE_SIZE) {
      setStage("error");
      setError("파일 크기는 50MB를 넘을 수 없습니다.");
      return;
    }

    setStage("reading");
    try {
      const ExcelJS = await import("exceljs");
      const buffer = await nextFile.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      const names = workbook.worksheets.map((sheet) => sheet.name);
      if (names.length === 0) throw new Error("시트 없음");
      bufferRef.current = buffer;
      setFile(nextFile);
      setSheetNames(names);
      setSelectedSheet(names[0]);
      setStage("ready");
    } catch {
      bufferRef.current = null;
      setFile(null);
      setSheetNames([]);
      setStage("error");
      setError("파일을 읽지 못했습니다. 손상되지 않은 .xlsx 파일인지 확인해 주세요.");
    }
  };

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const nextFile = event.target.files?.[0];
    if (nextFile) void inspectFile(nextFile);
    event.target.value = "";
  };

  const handleDirectoryFileChange = async (
    event: ChangeEvent<HTMLInputElement>,
  ) => {
    const nextFile = event.target.files?.[0];
    event.target.value = "";
    if (!nextFile) return;
    setError("");
    if (!/\.xlsx$/i.test(nextFile.name)) {
      setError("학교 기본현황은 .xlsx 파일만 사용할 수 있습니다.");
      return;
    }
    try {
      const ExcelJS = await import("exceljs");
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await nextFile.arrayBuffer());
      const nextDirectory = parseSchoolDirectoryWorkbook(
        workbook,
        nextFile.name,
      );
      setDirectory(nextDirectory);
      setDirectorySource(
        `${nextFile.name} · ${nextDirectory.count.toLocaleString()}개교`,
      );
      clearOutput();
      if (file) setStage("ready");
    } catch (caught) {
      setError(
        caught instanceof Error
          ? `학교 기본현황을 적용하지 못했습니다: ${caught.message}`
          : "학교 기본현황을 적용하지 못했습니다.",
      );
    }
  };

  const handleDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setIsDragging(false);
    const nextFile = event.dataTransfer.files?.[0];
    if (nextFile) void inspectFile(nextFile);
  };

  const handleUploadKey = (event: KeyboardEvent<HTMLDivElement>) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      inputRef.current?.click();
    }
  };

  const addColumn = (event: FormEvent) => {
    event.preventDefault();
    const normalized = normalizeColumn(newColumn);
    if (!normalized) {
      setColumnError("A부터 XFD 사이의 열 문자를 입력해 주세요.");
      return;
    }
    if (columns.includes(normalized)) {
      setColumnError("이미 등록된 열입니다.");
      return;
    }
    clearOutput();
    setColumns((current) =>
      [...current, normalized].sort(
        (a, b) => columnToNumber(a) - columnToNumber(b),
      ),
    );
    setNewColumn("");
    setColumnError("");
    setShowColumnInput(false);
    if (file) setStage("ready");
  };

  const removeColumn = (column: string) => {
    if (columns.length === 1) return;
    clearOutput();
    setColumns((current) => current.filter((item) => item !== column));
    if (file) setStage("ready");
  };

  const runSplitProcess = async () => {
    if (!file || !bufferRef.current || columns.length === 0) return;
    const ExcelJS = await import("exceljs");
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(bufferRef.current.slice(0));
    const columnNumbers = columns.map(columnToNumber);
    const targetSheets = applyAllSheets
      ? workbook.worksheets
      : [workbook.getWorksheet(selectedSheet)].filter(
          (sheet): sheet is Worksheet => Boolean(sheet),
        );
    if (targetSheets.length === 0) throw new Error("처리할 시트를 찾지 못했습니다.");

    const sheets = targetSheets.map((sheet) => processSheet(sheet, columnNumbers));
    workbook.calcProperties.fullCalcOnLoad = true;
    const output = await workbook.xlsx.writeBuffer();
    const name = outputFileName(file.name);
    return {
      blob: new Blob([output as BlobPart], { type: EXCEL_MIME }),
      result: {
        kind: "split" as const,
        sheets,
        splitRows: sheets.reduce((sum, sheet) => sum + sheet.splitRows, 0),
        addedRows: sheets.reduce((sum, sheet) => sum + sheet.addedRows, 0),
        outputName: name,
      },
    };
  };

  const runAwardProcess = async () => {
    if (!file || !bufferRef.current || !selectedSheet) return;
    const ExcelJS = await import("exceljs");
    const sourceWorkbook = new ExcelJS.Workbook();
    await sourceWorkbook.xlsx.load(bufferRef.current.slice(0));
    const sourceSheet = sourceWorkbook.getWorksheet(selectedSheet);
    if (!sourceSheet) throw new Error("선택한 시트를 찾지 못했습니다.");
    const outputData = createAwardOutput(sourceSheet);

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "과학전람회 엑셀 도구";
    const sheet = workbook.addWorksheet(safeWorksheetName(selectedSheet));
    sheet.addRow(outputData.headers);
    outputData.rows.forEach((row) => sheet.addRow(row));
    sheet.views = [{ state: "frozen", ySplit: 1 }];
    sheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: outputData.headers.length },
    };
    sheet.columns = outputData.headers.map((header, index) => ({
      key: header,
      width: index === 3 ? 28 : index < 3 ? 16 : 14,
    }));

    sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      row.height = rowNumber === 1 ? 24 : 21;
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.alignment = {
          horizontal: "center",
          vertical: "middle",
          wrapText: indexOfLineBreak(cell.value) >= 0,
        };
        cell.border = {
          top: { style: "thin", color: { argb: "FFB7C0BB" } },
          left: { style: "thin", color: { argb: "FFB7C0BB" } },
          bottom: { style: "thin", color: { argb: "FFB7C0BB" } },
          right: { style: "thin", color: { argb: "FFB7C0BB" } },
        };
        if (rowNumber === 1) {
          cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF0D7048" },
          };
        }
      });
    });

    const output = await workbook.xlsx.writeBuffer();
    const name = awardOutputFileName(file.name, selectedSheet);
    return {
      blob: new Blob([output as BlobPart], { type: EXCEL_MIME }),
      result: {
        kind: "award" as const,
        sourceSheet: selectedSheet,
        sourceRows: outputData.sourceRows,
        recipientRows: outputData.recipientRows,
        maxMembers: outputData.maxMembers,
        outputName: name,
      },
    };
  };

  const runAddressProcess = async () => {
    if (!file || !bufferRef.current || !directory) return;
    const ExcelJS = await import("exceljs");
    const sourceWorkbook = new ExcelJS.Workbook();
    await sourceWorkbook.xlsx.load(bufferRef.current.slice(0));
    const outputData = createSchoolAddressOutput(sourceWorkbook, directory);
    setLabelRows(outputData.rows);

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "과학전람회 엑셀 도구";
    const sheet = workbook.addWorksheet("수상학교 주소명단");
    sheet.addRow([...SCHOOL_ADDRESS_HEADERS]);
    outputData.rows.forEach((row) => {
      sheet.addRow(schoolAddressExcelRow(row, recipientName, deliveryNote));
    });
    sheet.views = [{ state: "frozen", ySplit: 1 }];
    sheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: 5 },
    };
    sheet.columns = [
      { key: "schoolName", width: 32 },
      { key: "postalCode", width: 13 },
      { key: "address", width: 72 },
      { key: "recipientName", width: 22 },
      { key: "deliveryNote", width: 40 },
    ];
    sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      row.height = rowNumber === 1 ? 25 : 22;
      row.eachCell({ includeEmpty: true }, (cell, columnNumber) => {
        cell.alignment = {
          horizontal:
            columnNumber === 3 || columnNumber === 5 ? "left" : "center",
          vertical: "middle",
          wrapText: columnNumber === 5,
        };
        cell.border = {
          top: { style: "thin", color: { argb: "FFC9D2CD" } },
          left: { style: "thin", color: { argb: "FFC9D2CD" } },
          bottom: { style: "thin", color: { argb: "FFC9D2CD" } },
          right: { style: "thin", color: { argb: "FFC9D2CD" } },
        };
        if (columnNumber === 2 && rowNumber > 1) cell.numFmt = "@";
        if (rowNumber === 1) {
          cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FF0D7048" },
          };
        } else if (!outputData.rows[rowNumber - 2]?.matched) {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFF2CC" },
          };
        }
      });
    });

    const output = await workbook.xlsx.writeBuffer();
    const name = schoolAddressOutputFileName(file.name);
    return {
      blob: new Blob([output as BlobPart], { type: EXCEL_MIME }),
      result: {
        kind: "address" as const,
        schoolCount: outputData.rows.length,
        matchedCount: outputData.matchedCount,
        unmatchedNames: outputData.unmatchedNames,
        directorySource,
        outputName: name,
      },
    };
  };

  const runProcess = async () => {
    if (!file || !bufferRef.current || stage === "processing") return;
    clearOutput();
    setError("");
    setStage("processing");
    try {
      const processed =
        mode === "split"
          ? await runSplitProcess()
          : mode === "award"
            ? await runAwardProcess()
            : await runAddressProcess();
      if (!processed) throw new Error("처리할 파일이나 시트를 확인해 주세요.");
      const url = URL.createObjectURL(processed.blob);
      setDownloadUrl(url);
      setResult(processed.result);
      setStage("done");
      triggerDownload(url, processed.result.outputName);
    } catch (caught) {
      setStage("error");
      setError(
        caught instanceof Error && caught.message
          ? `처리하지 못했습니다: ${caught.message}`
          : "파일 처리 중 오류가 발생했습니다.",
      );
    }
  };

  const printAddressLabels = () => {
    if (labelRows.length === 0) return;
    const html = buildLabelPrintHtml(labelRows, recipientName, deliveryNote);
    const blobUrl = URL.createObjectURL(
      new Blob([html], { type: "text/html;charset=utf-8" }),
    );
    const printWindow = window.open(blobUrl, "_blank", "noopener,noreferrer");
    if (!printWindow) {
      URL.revokeObjectURL(blobUrl);
      setError("인쇄 창을 열지 못했습니다. 브라우저의 팝업 차단을 해제해 주세요.");
      return;
    }
    window.setTimeout(() => URL.revokeObjectURL(blobUrl), 60_000);
  };

  const primaryLabel =
    stage === "reading"
      ? "파일 확인 중"
      : stage === "processing"
        ? mode === "split"
          ? "행을 분리하는 중"
          : mode === "award"
            ? "상장 명단을 만드는 중"
            : "학교 주소를 찾는 중"
        : !file
          ? "파일을 먼저 선택하세요"
          : stage === "done"
            ? "현재 설정으로 다시 처리"
            : mode === "split"
              ? "행 분리 및 다운로드"
              : mode === "award"
                ? "상장 명단 생성 및 다운로드"
                : "학교 주소 명단 생성 및 다운로드";
  const step = file ? (stage === "processing" || stage === "done" ? 3 : 2) : 1;
  const stepLabels =
    mode === "split"
      ? ["파일 선택", "열 확인", "분리 및 다운로드"]
      : mode === "award"
        ? ["파일 선택", "시트 선택", "명단 생성 및 다운로드"]
        : ["상장 명단 선택", "학교 데이터 확인", "주소 명단 다운로드"];

  return (
    <div className="app-shell">
      <header className="topbar">
        <div className="topbar-inner">
          <div className="brand">
            <span className="brand-mark" aria-hidden="true">
              <FileSpreadsheet size={34} strokeWidth={1.8} />
            </span>
            <div>
              <h1>과학전람회 엑셀 도구</h1>
              <p>행 분리, 상장 명단, 수상학교 주소 추출을 한곳에서 처리합니다.</p>
            </div>
          </div>
          <div className="privacy-note">
            <ShieldCheck size={22} aria-hidden="true" />
            <span>파일은 브라우저 안에서만 처리됩니다</span>
          </div>
        </div>
      </header>

      <main className="workspace">
        <div className="mode-tabs" role="tablist" aria-label="작업 선택">
          <button
            type="button"
            role="tab"
            aria-selected={mode === "split"}
            className={mode === "split" ? "is-active" : ""}
            onClick={() => changeMode("split")}
          >
            <ListTree size={20} />
            <span><strong>행 분리</strong><small>Alt+Enter 자료 나누기</small></span>
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={mode === "award"}
            className={mode === "award" ? "is-active" : ""}
            onClick={() => changeMode("award")}
          >
            <Award size={20} />
            <span><strong>상장 명단 만들기</strong><small>메일머지용 명단 생성</small></span>
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={mode === "address"}
            className={mode === "address" ? "is-active" : ""}
            onClick={() => changeMode("address")}
          >
            <Building2 size={20} />
            <span><strong>학교 주소 추출</strong><small>우편번호·주소 명단</small></span>
          </button>
        </div>

        <section className="card upload-card" aria-labelledby="upload-title">
          <input
            ref={inputRef}
            id="excel-file"
            className="sr-only"
            type="file"
            accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            onChange={handleFileChange}
          />
          <div
            className={`drop-zone ${isDragging ? "is-dragging" : ""} ${file ? "has-file" : ""}`}
            role="button"
            tabIndex={0}
            onClick={() => inputRef.current?.click()}
            onKeyDown={handleUploadKey}
            onDragEnter={(event) => {
              event.preventDefault();
              setIsDragging(true);
            }}
            onDragOver={(event) => event.preventDefault()}
            onDragLeave={(event) => {
              if (!event.currentTarget.contains(event.relatedTarget as Node | null)) {
                setIsDragging(false);
              }
            }}
            onDrop={handleDrop}
          >
            <span className="upload-icon" aria-hidden="true">
              {file ? <FileSpreadsheet size={42} /> : <Upload size={42} />}
            </span>
            {file ? (
              <>
                <h2 id="upload-title">{file.name}</h2>
                <p>{formatBytes(file.size)} · 시트 {sheetNames.length}개</p>
                <span className="file-action">다른 파일 선택</span>
              </>
            ) : (
              <>
                <h2 id="upload-title">파일을 여기에 끌어 놓으세요</h2>
                <p>
                  {mode === "split"
                    ? "또는 클릭하여 행을 분리할 .xlsx 파일 선택"
                    : mode === "award"
                      ? "또는 클릭하여 포상 대상자 명단 .xlsx 파일 선택"
                      : "또는 클릭하여 생성된 상장 명단 .xlsx 파일 선택"}
                </p>
                <span className="size-badge">최대 50MB</span>
              </>
            )}
          </div>
        </section>

        {mode === "split" ? (
          <section className="card criteria-card" aria-labelledby="criteria-title">
            <div className="section-heading">
              <div>
                <h2 id="criteria-title">분리 기준</h2>
                <p>선택한 열의 Alt+Enter 줄바꿈을 같은 순서로 나눕니다.</p>
              </div>
              <label className="toggle-label">
                <span>모든 시트 적용</span>
                <button
                  type="button"
                  className={`toggle ${applyAllSheets ? "is-on" : ""}`}
                  role="switch"
                  aria-checked={applyAllSheets}
                  onClick={() => {
                    clearOutput();
                    setApplyAllSheets((current) => !current);
                    if (file) setStage("ready");
                  }}
                >
                  <span />
                </button>
              </label>
            </div>

            <div className="column-section">
              <span className="field-label">대상 열</span>
              <div className="column-controls">
                {columns.map((column) => (
                  <span className="column-chip" key={column}>
                    <strong>{column}</strong>
                    <button
                      type="button"
                      aria-label={`${column}열 삭제`}
                      title={columns.length === 1 ? "기준 열은 하나 이상 필요합니다" : `${column}열 삭제`}
                      disabled={columns.length === 1}
                      onClick={() => removeColumn(column)}
                    >
                      <X size={18} />
                    </button>
                  </span>
                ))}

                {showColumnInput ? (
                  <form className="column-form" onSubmit={addColumn}>
                    <input
                      autoFocus
                      value={newColumn}
                      maxLength={3}
                      aria-label="추가할 열 문자"
                      placeholder="예: F"
                      onChange={(event) => {
                        setNewColumn(event.target.value.toUpperCase());
                        setColumnError("");
                      }}
                    />
                    <button type="submit" aria-label="열 추가 확인"><Check size={18} /></button>
                    <button
                      type="button"
                      aria-label="열 추가 취소"
                      onClick={() => {
                        setShowColumnInput(false);
                        setNewColumn("");
                        setColumnError("");
                      }}
                    ><X size={18} /></button>
                  </form>
                ) : (
                  <button type="button" className="add-column" onClick={() => setShowColumnInput(true)}>
                    <Plus size={19} /> 열 추가
                  </button>
                )}
              </div>
              {columnError && <p className="inline-error">{columnError}</p>}
            </div>

            {!applyAllSheets && file && (
              <label className="sheet-select">
                <span>처리할 시트</span>
                <select value={selectedSheet} onChange={(event) => setSelectedSheet(event.target.value)}>
                  {sheetNames.map((name) => <option key={name} value={name}>{name}</option>)}
                </select>
              </label>
            )}

            <div className="rule-row">
              <span>적용 규칙</span>
              <p>한 줄 값은 필요한 만큼 반복 · 항목이 부족한 칸은 비움 · 다른 셀은 그대로 복제</p>
            </div>
          </section>
        ) : mode === "award" ? (
          <section className="card criteria-card award-card" aria-labelledby="award-title">
            <div className="section-heading">
              <div>
                <h2 id="award-title">상장 명단 생성 기준</h2>
                <p>포상 대상자 명단에서 메일머지용 구조를 자동으로 만듭니다.</p>
              </div>
              <span className="mode-badge">예시 형식 적용</span>
            </div>

            <label className="sheet-select always-visible">
              <span>처리할 시트</span>
              <select
                value={selectedSheet}
                disabled={!file}
                onChange={(event) => {
                  clearOutput();
                  setSelectedSheet(event.target.value);
                  if (file) setStage("ready");
                }}
              >
                {!file && <option>파일을 먼저 선택하세요</option>}
                {sheetNames.map((name) => <option key={name} value={name}>{name}</option>)}
              </select>
            </label>

            <div className="mapping-grid" aria-label="열 변환 규칙">
              <span><strong>A</strong><small>호수</small><b>→</b><em>1</em></span>
              <span><strong>E</strong><small>마지막 괄호의 부문</small><b>→</b><em>2</em></span>
              <span><strong>J</strong><small>상종 및 등급</small><b>→</b><em>3</em></span>
              <span><strong>F</strong><small>학생별 소속 학교</small><b>→</b><em>sc1</em></span>
              <span><strong>H</strong><small>학생별 학년</small><b>→</b><em>s1…</em></span>
              <span><strong>I</strong><small>학생별 이름</small><b>→</b><em>n1…</em></span>
            </div>

            <div className="rule-row">
              <span>반복 규칙</span>
              <p>한 행의 학생 수만큼 결과 행을 생성합니다. F열 소속이 여러 개이면 I열 학생과 같은 순서로 D열에 하나씩 넣고, 소속이 하나이면 모든 학생에게 반복합니다. 부문명은 예시처럼 ‘부문’과 공백을 제거합니다.</p>
            </div>
          </section>
        ) : (
          <section className="card criteria-card address-card" aria-labelledby="address-title">
            <div className="section-heading">
              <div>
                <h2 id="address-title">수상학교 주소 명단</h2>
                <p>상장 명단의 sc1 열에서 학교를 모아 우편번호와 주소를 찾습니다.</p>
              </div>
              <span className="mode-badge">중복 학교 자동 제거</span>
            </div>

            <div className="directory-panel">
              <div className="directory-copy">
                <span className="directory-icon" aria-hidden="true">
                  <Building2 size={22} />
                </span>
                <div>
                  <strong>학교 기본현황</strong>
                  <p>{directorySource}</p>
                </div>
              </div>
              <input
                ref={directoryInputRef}
                className="sr-only"
                type="file"
                accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                onChange={(event) => void handleDirectoryFileChange(event)}
              />
              <button
                type="button"
                className="directory-update"
                onClick={() => directoryInputRef.current?.click()}
              >
                <UploadCloud size={18} />
                새 학교현황 적용
              </button>
            </div>

            <div className="address-flow" aria-label="학교 주소 추출 규칙">
              <span><strong>1</strong><small>상장 명단 sc1</small></span>
              <b>→</b>
              <span><strong>2</strong><small>학교명 중복 제거</small></span>
              <b>→</b>
              <span><strong>3</strong><small>주소·수신자·주의사항</small></span>
            </div>

            <div className="address-extra-fields">
              <label>
                <span>수신자 이름</span>
                <input
                  type="text"
                  value={recipientName}
                  placeholder="예: 과학담당 선생님"
                  onChange={(event) => {
                    clearOutput();
                    setRecipientName(event.target.value);
                    if (file) setStage("ready");
                  }}
                />
                <small>모든 학교 행의 주소 다음 열에 동일하게 입력됩니다.</small>
              </label>
              <label>
                <span>주의사항</span>
                <textarea
                  value={deliveryNote}
                  rows={3}
                  placeholder="예: 상장 재중 · 접지 금지"
                  onChange={(event) => {
                    clearOutput();
                    setDeliveryNote(event.target.value);
                    if (file) setStage("ready");
                  }}
                />
                <small>수신자 이름 다음 열에 동일하게 입력됩니다.</small>
              </label>
            </div>

            <div className="rule-row">
              <span>데이터 관리</span>
              <p>앱에는 첨부한 학교 기본현황에서 만든 별도 데이터가 포함되어 있습니다. 최신 현황이 필요하면 위 버튼으로 새 엑셀을 올려 현재 작업에 즉시 적용할 수 있습니다.</p>
            </div>
            <div className="rule-row">
              <span>라벨 인쇄</span>
              <p>아이라벨 {LABEL_SPEC.code} 10칸(84.5×53.5mm) 규격입니다. 인쇄 창에서는 배율을 ‘실제 크기’ 또는 100%로 선택하고 머리글·바닥글을 끄세요.</p>
            </div>
          </section>
        )}

        {error && (
          <div className="message error-message" role="alert">
            <AlertCircle size={20} /><span>{error}</span>
          </div>
        )}

        <button
          type="button"
          className="primary-button"
          disabled={
            !file ||
            stage === "reading" ||
            stage === "processing" ||
            (mode === "address" && (directoryLoading || !directory))
          }
          onClick={() => void runProcess()}
        >
          {stage === "reading" || stage === "processing"
            ? <LoaderCircle className="spinner" size={22} />
            : stage === "done"
              ? <RotateCcw size={21} />
              : <Download size={21} />}
          {primaryLabel}
        </button>

        {result && (
          <section className="card result-card" aria-live="polite">
            <div className="result-icon"><CheckCircle2 size={30} /></div>
            <div className="result-copy">
              <h2>
                {result.kind === "split"
                  ? "행 분리가 완료되었습니다"
                  : result.kind === "award"
                    ? "상장 명단이 완성되었습니다"
                    : "수상학교 주소 명단이 완성되었습니다"}
              </h2>
              {result.kind === "split" ? (
                <>
                  <p>{result.sheets.length}개 시트에서 {result.splitRows}개 행을 나누고, 새 행 {result.addedRows}개를 만들었습니다.</p>
                  <div className="sheet-summary">
                    {result.sheets.map((sheet) => <span key={sheet.name}>{sheet.name} <strong>+{sheet.addedRows}</strong></span>)}
                  </div>
                </>
              ) : result.kind === "award" ? (
                <>
                  <p>‘{result.sourceSheet}’ 시트의 작품 {result.sourceRows}건을 학생 {result.recipientRows}명용 행으로 만들었습니다.</p>
                  <div className="sheet-summary">
                    <span>최대 학생 수 <strong>{result.maxMembers}명</strong></span>
                    <span>생성 행 <strong>{result.recipientRows}개</strong></span>
                  </div>
                </>
              ) : (
                <>
                  <p>수상학교 {result.schoolCount}개 중 {result.matchedCount}개의 우편번호와 주소를 찾았습니다.</p>
                  <div className="sheet-summary">
                    <span>학교 데이터 <strong>{result.directorySource}</strong></span>
                    <span>일치 <strong>{result.matchedCount}개교</strong></span>
                    <span>미확인 <strong>{result.unmatchedNames.length}개교</strong></span>
                  </div>
                  {result.unmatchedNames.length > 0 && (
                    <p className="unmatched-note">
                      미확인 학교: {result.unmatchedNames.join(", ")}
                    </p>
                  )}
                </>
              )}
            </div>
            <div className="result-actions">
              {result.kind === "address" && (
                <button type="button" className="print-labels" onClick={printAddressLabels}>
                  <Building2 size={18} /> 라벨 인쇄
                </button>
              )}
              <button type="button" className="download-again" onClick={() => triggerDownload(downloadUrl, result.outputName)}>
                <Download size={18} /> 다시 다운로드
              </button>
            </div>
          </section>
        )}

        <ol className="steps" aria-label="처리 단계">
          {stepLabels.map((label, index) => (
            <li className={step >= index + 1 ? "is-active" : ""} key={label}>
              <span>{step > index + 1 ? <Check size={16} /> : index + 1}</span>
              <strong>{label}</strong>
            </li>
          ))}
        </ol>

        {file && (
          <button type="button" className="reset-button" onClick={resetFile}>
            <RotateCcw size={16} /> 처음부터 다시
          </button>
        )}
      </main>

      <footer>업로드한 파일은 외부 서버로 전송되거나 저장되지 않습니다.</footer>
    </div>
  );
}

function indexOfLineBreak(value: unknown) {
  return typeof value === "string" ? value.indexOf("\n") : -1;
}
