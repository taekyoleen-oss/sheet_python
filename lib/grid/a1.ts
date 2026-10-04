// A1 참조 파서·포매터 — 설계서 §1.7 (sheet() 참조와 상태 바 표기가 공유)

import type { CellRange } from "@/types/workbook";

/** 잘못된 A1 참조 오류 (한국어 메시지) */
export class A1Error extends Error {
  readonly code = "A1_INVALID";
  constructor(ref: string, reason?: string) {
    super(`잘못된 셀 참조입니다: "${ref}"${reason ? ` — ${reason}` : ""}`);
    this.name = "A1Error";
  }
}

/** 0-based 열 번호 → 열 문자 (0="A", 25="Z", 26="AA") */
export function colToLetter(col: number): string {
  if (!Number.isInteger(col) || col < 0) {
    throw new A1Error(String(col), "열 번호는 0 이상의 정수여야 합니다");
  }
  let s = "";
  let n = col;
  for (;;) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
    if (n < 0) break;
  }
  return s;
}

/** 열 문자 → 0-based 열 번호 ("A"=0, "Z"=25, "AA"=26) */
export function letterToCol(letters: string): number {
  if (!/^[A-Za-z]+$/.test(letters)) {
    throw new A1Error(letters, "열 문자가 아닙니다");
  }
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export interface ParsedA1 {
  sheetName?: string;
  /** r0<=r1, c0<=c1로 정규화된 범위 */
  range: CellRange;
  /** 단일 셀 참조 여부 */
  scalar: boolean;
}

const CELL_RE = /^([A-Za-z]{1,3})([1-9][0-9]{0,6})$/;

function parseCell(ref: string, text: string): { r: number; c: number } {
  const m = CELL_RE.exec(text);
  if (!m) throw new A1Error(ref);
  return { r: parseInt(m[2], 10) - 1, c: letterToCol(m[1]) };
}

/** "A1" | "A1:C10" | "Sheet2!A1" | "'시트 이름'!A1:B2" 파싱 */
export function parseA1(ref: string): ParsedA1 {
  if (typeof ref !== "string" || ref.trim() === "") throw new A1Error(String(ref));
  let rest = ref.trim();
  let sheetName: string | undefined;

  if (rest.startsWith("'")) {
    const m = /^'((?:[^']|'')+)'!(.+)$/.exec(rest);
    if (!m) throw new A1Error(ref, "따옴표 시트 이름 형식이 아닙니다");
    sheetName = m[1].replace(/''/g, "'");
    rest = m[2];
  } else if (rest.includes("!")) {
    const i = rest.indexOf("!");
    sheetName = rest.slice(0, i);
    rest = rest.slice(i + 1);
    if (sheetName === "" || sheetName.includes("'")) throw new A1Error(ref, "시트 이름이 비어 있거나 잘못되었습니다");
  }

  const parts = rest.split(":");
  if (parts.length > 2) throw new A1Error(ref);
  const a = parseCell(ref, parts[0]);
  const b = parts.length === 2 ? parseCell(ref, parts[1]) : a;
  return {
    sheetName,
    range: {
      r0: Math.min(a.r, b.r),
      c0: Math.min(a.c, b.c),
      r1: Math.max(a.r, b.r),
      c1: Math.max(a.c, b.c),
    },
    scalar: parts.length === 1,
  };
}

/** 시트 이름에 따옴표가 필요한지 (공백·특수문자·숫자 시작) */
const needsQuote = (name: string): boolean =>
  !/^[A-Za-z가-힣_][A-Za-z0-9가-힣_.]*$/.test(name);

/** CellRange → "A1" | "A1:C10" (+ 시트 접두어, 필요 시 '따옴표') */
export function formatA1(range: CellRange, sheetName?: string): string {
  const { r0, c0, r1, c1 } = range;
  if ([r0, c0, r1, c1].some((n) => !Number.isInteger(n) || n < 0)) {
    throw new A1Error(JSON.stringify(range), "범위 값이 잘못되었습니다");
  }
  const lo = { r: Math.min(r0, r1), c: Math.min(c0, c1) };
  const hi = { r: Math.max(r0, r1), c: Math.max(c0, c1) };
  const cell = (p: { r: number; c: number }) => `${colToLetter(p.c)}${p.r + 1}`;
  const body =
    lo.r === hi.r && lo.c === hi.c ? cell(lo) : `${cell(lo)}:${cell(hi)}`;
  if (sheetName === undefined) return body;
  const prefix = needsQuote(sheetName)
    ? `'${sheetName.replace(/'/g, "''")}'`
    : sheetName;
  return `${prefix}!${body}`;
}
