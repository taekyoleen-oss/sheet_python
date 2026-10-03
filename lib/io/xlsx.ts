// XLSX 열기/내보내기 — SheetJS 0.20.3 (값 + 수식·정의된 이름(부록 O.4), 코드·서식 제외)
// 가져오기: 이 앱 엔진이 해석할 수 있는 수식만 fx로 살리고(값은 엑셀 캐시), 못 하면 값만 남긴다.

import * as XLSX from "xlsx";
import { cellKey, parseCellKey, type Cell, type DefinedName, type Sheet } from "@/types/workbook";
import { parseFormula, parseRefText } from "@/lib/grid/formula";
import { createSheet } from "@/lib/grid/model";

const pad2 = (n: number) => String(n).padStart(2, "0");

const dateToIso = (d: Date): string =>
  `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;

/** SheetJS 워크시트 → Sheet (셀 타입 매핑 n/s/b/d, 날짜 ISO) */
export interface FormulaImportStats {
  kept: number;
  /** 미지원 함수·문법이라 값만 가져온 수식 수 */
  dropped: number;
}

export function wsToSheet(ws: XLSX.WorkSheet, name: string, stats?: FormulaImportStats): Sheet {
  const sheet = createSheet(name);
  const ref = ws["!ref"];
  if (!ref) return sheet;
  const range = XLSX.utils.decode_range(ref);
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr] as XLSX.CellObject | undefined;
      if (!cell) continue;
      if (cell.f) {
        const fx = `=${cell.f}`;
        if (parseFormula(fx).error === undefined) {
          const v = cell.t === "e" ? (cell.w ?? "#VALUE!") : cell.v instanceof Date ? dateToIso(cell.v) : (cell.v ?? null);
          const t = cell.t === "e" ? "e" : cell.t === "b" ? "b" : cell.t === "s" || cell.t === "d" ? "s" : "n";
          sheet.cells[cellKey(r, c)] = { v: v as Cell["v"], t, fx };
          if (stats) stats.kept++;
          continue;
        }
        if (stats) stats.dropped++;
      }
      if (cell.v === undefined || cell.v === null) continue;
      let out: Cell;
      switch (cell.t) {
        case "n":
          out = { v: cell.v as number, t: "n" };
          break;
        case "b":
          out = { v: cell.v as boolean, t: "b" };
          break;
        case "d":
          out = { v: dateToIso(cell.v as Date), t: "d", f: "yyyy-mm-dd" };
          break;
        default: {
          const s = String(cell.v);
          if (s === "") continue;
          out = { v: s, t: "s" };
        }
      }
      sheet.cells[cellKey(r, c)] = out;
    }
  }
  sheet.rowCount = Math.max(sheet.rowCount, range.e.r + 1);
  sheet.colCount = Math.max(sheet.colCount, range.e.c + 1);
  return sheet;
}

/** XLSX 통합 문서 → 시트 + 정의된 이름 + 수식 가져오기 통계 (파일 > 열기) */
export function readXlsxBook(data: ArrayBuffer): {
  sheets: Sheet[];
  names: DefinedName[];
  stats: FormulaImportStats;
} {
  const wb = XLSX.read(data, { type: "array", cellDates: true, cellFormula: true });
  const stats: FormulaImportStats = { kept: 0, dropped: 0 };
  const sheets = wb.SheetNames.map((name) => wsToSheet(wb.Sheets[name], name, stats));
  const names: DefinedName[] = [];
  for (const n of wb.Workbook?.Names ?? []) {
    // 시트 범위 이름(Sheet 지정)·수식 이름·외부 참조는 건너뛴다
    if (n.Sheet !== undefined || !n.Ref) continue;
    const ref = n.Ref.replace(/^=/, "");
    const p = parseRefText(ref);
    if (p?.sheetName !== undefined) names.push({ name: n.Name, ref });
  }
  return { sheets, names, stats };
}

/** CSV/XLSX 파일 데이터 → Sheet[] (시트 이름 유지) */
export function sheetsFromFileData(data: ArrayBuffer): Sheet[] {
  const wb = XLSX.read(data, { type: "array", cellDates: true });
  return wb.SheetNames.map((name) => wsToSheet(wb.Sheets[name], name));
}

/** Sheet → SheetJS 워크시트 (사용 범위만, 값만: 날짜 ISO 문자열·불리언 그대로) */
export function sheetToWs(sheet: Sheet): XLSX.WorkSheet {
  let maxR = 0;
  let maxC = 0;
  const keys = Object.keys(sheet.cells);
  for (const key of keys) {
    const { r, c } = parseCellKey(key);
    if (r > maxR) maxR = r;
    if (c > maxC) maxC = c;
  }
  const aoa: (string | number | boolean | null)[][] = Array.from(
    { length: keys.length === 0 ? 1 : maxR + 1 },
    () => new Array(maxC + 1).fill(null),
  );
  for (const key of keys) {
    const { r, c } = parseCellKey(key);
    aoa[r][c] = sheet.cells[key].v;
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  // 부록 O.4: 수식은 원문을 함께 쓴다(값은 캐시) — 엑셀에서 열면 다시 계산된다
  for (const key of keys) {
    const fx = sheet.cells[key].fx;
    if (!fx) continue;
    const { r, c } = parseCellKey(key);
    const addr = XLSX.utils.encode_cell({ r, c });
    const cell = (ws[addr] ??= { t: "z" }) as XLSX.CellObject;
    cell.f = fx.replace(/^=/, "");
  }
  return ws;
}

/** 전 시트 → .xlsx Blob */
export function sheetsToXlsxBlob(sheets: Sheet[], names: DefinedName[] = []): Blob {
  const wb = XLSX.utils.book_new();
  if (names.length > 0) wb.Workbook = { Names: names.map((n) => ({ Name: n.name, Ref: n.ref })) };
  for (const sheet of sheets) {
    // Excel 시트 이름 제약(31자·금지 문자)은 SheetJS가 오류를 내므로 최소 정리
    const name = sheet.name.replace(/[\\/?*[\]:]/g, "_").slice(0, 31) || "Sheet";
    XLSX.utils.book_append_sheet(wb, sheetToWs(sheet), name);
  }
  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" }) as ArrayBuffer;
  return new Blob([out], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}
