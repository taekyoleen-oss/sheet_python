// 앱 내부 복사·붙여넣기 — 부록 O.3. 시스템 클립보드에는 값(TSV/HTML)만 실리므로,
// 이 앱에서 복사한 범위는 원본 셀(수식 원문 포함)을 메모리에 기억해 두었다가
// 같은 텍스트가 붙여넣어지면 엑셀처럼 수식의 상대 참조를 이동시켜 붙인다.
// 부록 O.4: 잘라내기(cut) 표시가 있으면 붙여넣기가 "이동"이 된다(1회용), 채우기 핸들 연속 패턴.

import { shiftFormula } from "@/lib/grid/formula";
import type { CellEdit } from "@/lib/grid/model";
import { cellKey, type Cell, type CellRange, type Sheet } from "@/types/workbook";

export interface CopiedBlock {
  /** 복사 시 시스템 클립보드에 쓴 text/plain — 붙여넣기 출처 판별용 */
  text: string;
  /** 원본 시트 (잘라내기 이동용) */
  sheetId?: string;
  /** Ctrl+X로 잘라낸 범위 — 붙여넣으면 이동하고 기억을 지운다 */
  cut?: boolean;
  r0: number;
  c0: number;
  cells: (Cell | null)[][];
}

let last: CopiedBlock | null = null;

/** 범위의 원본 셀 스냅샷 (채우기 Ctrl+D/R도 같은 형태를 쓴다) */
export function snapshotRange(sheet: Sheet, range: CellRange, text = ""): CopiedBlock {
  const r0 = Math.min(range.r0, range.r1);
  const c0 = Math.min(range.c0, range.c1);
  const cells: (Cell | null)[][] = [];
  for (let r = r0; r <= Math.max(range.r0, range.r1); r++) {
    const row: (Cell | null)[] = [];
    for (let c = c0; c <= Math.max(range.c0, range.c1); c++) {
      const cell = sheet.cells[cellKey(r, c)];
      if (!cell) {
        row.push(null);
        continue;
      }
      // spill(src) 셀은 Python 출력이라 값으로만 복사한다 (잠금 표시는 따라가지 않는다)
      const { src: _src, ...rest } = cell;
      row.push(rest);
    }
    cells.push(row);
  }
  return { text, r0, c0, cells };
}

export function rememberCopy(sheet: Sheet, range: CellRange, text: string, cut = false): void {
  last = { ...snapshotRange(sheet, range, text), sheetId: sheet.id, cut };
}

/** 잘라내기는 한 번 붙이면 끝 (엑셀 동일) */
export const forgetCopy = (): void => {
  last = null;
};

/** 붙여넣은 text/plain이 이 앱의 마지막 복사본이면 그것을 돌려준다 */
export const internalCopyFor = (text: string | undefined): CopiedBlock | null =>
  last && text !== undefined && normalizeEol(text) === normalizeEol(last.text) ? last : null;

const normalizeEol = (s: string) => s.replace(/\r\n/g, "\n").replace(/\n$/, "");

/**
 * 붙여넣기 편집 목록. 대상 선택이 원본 크기의 정수배면 타일링(엑셀: 한 셀 복사 → 범위 선택 →
 * 붙여넣기로 수식 채우기). 수식은 (대상 - 원본) 만큼 상대 참조 이동, 값은 계산기가 다시 채운다.
 */
export function buildPasteEdits(src: CopiedBlock, dest: CellRange): CellEdit[] {
  const h = src.cells.length;
  const w = src.cells[0]?.length ?? 0;
  const dr0 = Math.min(dest.r0, dest.r1);
  const dc0 = Math.min(dest.c0, dest.c1);
  const dh = Math.abs(dest.r1 - dest.r0) + 1;
  const dw = Math.abs(dest.c1 - dest.c0) + 1;
  const tilesR = dh % h === 0 ? dh / h : 1;
  const tilesC = dw % w === 0 ? dw / w : 1;
  const edits: CellEdit[] = [];
  for (let tr = 0; tr < tilesR; tr++)
    for (let tc = 0; tc < tilesC; tc++)
      src.cells.forEach((row, i) =>
        row.forEach((cell, j) => {
          const r = dr0 + tr * h + i;
          const c = dc0 + tc * w + j;
          const out: Cell | null =
            cell?.fx !== undefined
              ? { ...cell, v: null, fx: shiftFormula(cell.fx, r - (src.r0 + i), c - (src.c0 + j)) }
              : cell && { ...cell };
          edits.push({ r, c, cell: out });
        }),
      );
  return edits;
}

// ── 채우기 핸들 (부록 O.4) ──────────────────────────────

/** 위치 0..n-1의 값에 최소제곱 직선 a + b·k (엑셀 채우기 핸들의 "선형 추세") */
function linearFit(ys: number[]): (k: number) => number {
  const n = ys.length;
  const mx = (n - 1) / 2;
  const my = ys.reduce((a, b) => a + b, 0) / n;
  let sxy = 0;
  let sxx = 0;
  ys.forEach((y, k) => {
    sxy += (k - mx) * (y - my);
    sxx += (k - mx) ** 2;
  });
  const b = sxx === 0 ? 0 : sxy / sxx;
  return (k) => Number((my + b * (k - mx)).toPrecision(15));
}

const TRAILING_NUM = /^(.*?)(\d+)$/;

/**
 * 채우기 핸들로 src를 dest(원본 포함 가능)까지 늘릴 때의 편집 목록.
 * 줄(세로 채우기면 열, 가로면 행)마다:
 * - 수식 → 상대 참조 이동 (반복)
 * - 숫자 2개 이상 → 선형 추세 연장 (1,2 → 3,4 / 위·왼쪽으로 끌면 거꾸로)
 * - 끝이 숫자인 문자 1개("항목1") → 숫자 증가
 * - 그 외(숫자 1개 포함) → 그대로 반복 복사 (엑셀 동일)
 */
export function buildFillEdits(sheet: Sheet, src: CellRange, dest: CellRange): CellEdit[] {
  const s = { r0: Math.min(src.r0, src.r1), c0: Math.min(src.c0, src.c1), r1: Math.max(src.r0, src.r1), c1: Math.max(src.c0, src.c1) };
  const d = { r0: Math.min(dest.r0, dest.r1), c0: Math.min(dest.c0, dest.c1), r1: Math.max(dest.r0, dest.r1), c1: Math.max(dest.c0, dest.c1) };
  const vertical = d.r0 < s.r0 || d.r1 > s.r1;
  const snap = snapshotRange(sheet, s);
  const n = vertical ? s.r1 - s.r0 + 1 : s.c1 - s.c0 + 1;
  const lines = vertical ? s.c1 - s.c0 + 1 : s.r1 - s.r0 + 1;
  const edits: CellEdit[] = [];
  for (let li = 0; li < lines; li++) {
    const line = Array.from({ length: n }, (_, k) => (vertical ? snap.cells[k][li] : snap.cells[li][k]));
    const nums = line.every((c) => c && !c.fx && c.t === "n" && typeof c.v === "number")
      ? linearFit(line.map((c) => c!.v as number))
      : null;
    const lo = vertical ? d.r0 : d.c0;
    const hi = vertical ? d.r1 : d.c1;
    const base = vertical ? s.r0 : s.c0;
    for (let pos = lo; pos <= hi; pos++) {
      const k = pos - base;
      if (k >= 0 && k < n) continue; // 원본 자리
      const idx = ((k % n) + n) % n;
      const from = line[idx];
      const r = vertical ? pos : s.r0 + li;
      const c = vertical ? s.c0 + li : pos;
      let cell: Cell | null;
      if (from?.fx !== undefined) {
        cell = { ...from, v: null, fx: shiftFormula(from.fx, vertical ? k - idx : 0, vertical ? 0 : k - idx) };
      } else if (nums && n >= 2) {
        cell = { ...from!, v: nums(k) };
      } else if (n === 1 && from?.t === "s" && typeof from.v === "string" && TRAILING_NUM.test(from.v)) {
        const [, head, digits] = TRAILING_NUM.exec(from.v)!;
        const next = Number(digits) + k;
        cell = next < 0 ? { ...from } : { ...from, v: head + String(next).padStart(digits.length, "0") };
      } else {
        cell = from && { ...from };
      }
      edits.push({ r, c, cell });
    }
  }
  return edits;
}
