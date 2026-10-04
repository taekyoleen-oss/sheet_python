// 수식 파서·평가기 — 부록 I(미니 수식) → 부록 O.1(엑셀식 수식 확장). 순수 함수, 스토어 무관.
// 문법: 비교(= <> < > <= >=)·연결(&)·사칙·거듭제곱(^)·단항 ±·백분율(%)·문자열/불리언 리터럴,
// 셀 참조(상대 A1·절대 $A$1·혼합 $A1/A$1, 시트 접두어), 함수 FUNCS 목록.
// 참조 재작성(rewriteRefs·shiftFormula·adjustForStructure)도 같은 토크나이저를 쓴다 — 부록 O.2.

import type { Cell, CellRange } from "@/types/workbook";
import { colToLetter, letterToCol } from "./a1";

export type FormulaErrorCode =
  | "#NAME?"
  | "#REF!"
  | "#VALUE!"
  | "#DIV/0!"
  | "#CIRC!"
  | "#N/A"
  | "#NUM!";

/** 오류 셀 hover 툴팁용 한국어 설명 (부록 I.3) */
export const FORMULA_ERROR_KO: Record<FormulaErrorCode, string> = {
  "#NAME?": "지원하지 않는 함수·이름 또는 문법입니다",
  "#REF!": "잘못된 셀 참조이거나 시트를 찾을 수 없습니다 (삭제된 행·열 참조 포함)",
  "#VALUE!": "값의 형식이 맞지 않습니다 (예: 문자에 사칙연산, 범위를 단일 값 자리에 사용)",
  "#DIV/0!": "0으로 나눌 수 없습니다",
  "#CIRC!": "순환 참조 — 수식들이 서로를 참조합니다",
  "#N/A": "찾는 값이 없습니다 (VLOOKUP·MATCH 등)",
  "#NUM!": "숫자 범위를 벗어났습니다 (예: 음수의 제곱근, 0의 로그)",
};

export const isFormulaError = (v: unknown): v is FormulaErrorCode =>
  typeof v === "string" && v in FORMULA_ERROR_KO;

/** 셀 입력이 수식인지 — `=` 시작(`=` 단독 제외) */
export const isFormula = (input: string): boolean => {
  const t = input.trim();
  return t.startsWith("=") && t.length > 1;
};

export interface FormulaRef {
  /** 시트 접두어 (없으면 수식이 있는 시트) */
  sheetName?: string;
  range: CellRange;
}

/**
 * 평가 시 셀 공급자 — 시트 이름을 해석할 수 없으면 "#REF!"를 반환한다.
 * sheetName === undefined는 수식이 놓인 시트.
 */
export type GetCell = (
  sheetName: string | undefined,
  r: number,
  c: number,
) => Cell | undefined | "#REF!";

/** 계산 결과 — 그대로 Cell의 v·t가 된다. 오류는 t:'e' + 코드 */
export interface FormulaResult {
  v: number | string | boolean;
  t: "n" | "s" | "b" | "e";
}

/** 시트의 사용 크기 — 전체 열/행 참조(A:A)를 실제 크기로 자른다. undefined면 자르지 않는다 */
export type GetBounds = (sheetName: string | undefined) => { rows: number; cols: number } | undefined;

export interface ParsedFormula {
  refs: FormulaRef[];
  /** 파싱 자체가 실패했으면 그 오류 코드 (XLSX 가져오기에서 미지원 수식 판별용) */
  error?: FormulaErrorCode;
  eval(getCell: GetCell, bounds?: GetBounds): FormulaResult;
}

// ── 참조 표기 ($ 절대·상대) ──────────────────────────────

export interface RefCorner {
  r: number;
  c: number;
  /** $행 고정 */
  ar: boolean;
  /** $열 고정 */
  ac: boolean;
}

/**
 * 토큰 하나의 참조. b가 없으면 단일 셀. 파싱 시 a=좌상, b=우하로 정규화된다.
 * kind "cols"(A:C)는 a.c·b.c만, "rows"(2:5)는 a.r·b.r만 의미가 있다(나머지 축은 시트 전체).
 */
export interface RefParts {
  sheetName?: string;
  a: RefCorner;
  b?: RefCorner;
  kind?: "cols" | "rows";
}

export const MAX_ROW = 1_048_575; // 엑셀 한계 (0-based)
export const MAX_COL = 16_383; // XFD

const SHEET_SRC = "'(?:[^']|'')+'|[A-Za-z가-힣_][A-Za-z0-9가-힣_.]*";
const CELL_SRC = "\\$?[A-Za-z]{1,3}\\$?[1-9][0-9]{0,6}";
const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)$/;

function parseCorner(text: string): RefCorner | null {
  const m = CELL_RE.exec(text);
  if (!m) return null;
  const r = Number(m[4]) - 1;
  const c = letterToCol(m[2]);
  if (r > MAX_ROW || c > MAX_COL) return null;
  return { r, c, ar: m[3] === "$", ac: m[1] === "$" };
}

/** 정의된 이름·시트 이름에 쓸 수 있는 형식 */
export const NAME_RE = /^[A-Za-z가-힣_][A-Za-z0-9가-힣_.]*$/;
const needsQuote = (name: string): boolean => !NAME_RE.test(name);

const cornerText = (k: RefCorner): string =>
  `${k.ac ? "$" : ""}${colToLetter(k.c)}${k.ar ? "$" : ""}${k.r + 1}`;

export function formatRef(p: RefParts): string {
  const prefix =
    p.sheetName === undefined
      ? ""
      : `${needsQuote(p.sheetName) ? `'${p.sheetName.replace(/'/g, "''")}'` : p.sheetName}!`;
  const b = p.b ?? p.a;
  if (p.kind === "cols")
    return `${prefix}${p.a.ac ? "$" : ""}${colToLetter(p.a.c)}:${b.ac ? "$" : ""}${colToLetter(b.c)}`;
  if (p.kind === "rows") return `${prefix}${p.a.ar ? "$" : ""}${p.a.r + 1}:${b.ar ? "$" : ""}${b.r + 1}`;
  return prefix + cornerText(p.a) + (p.b ? `:${cornerText(p.b)}` : "");
}

const rangeOf = (p: RefParts): CellRange => {
  const b = p.b ?? p.a;
  return { r0: p.a.r, c0: p.a.c, r1: b.r, c1: b.c };
};

const unquote = (sheet: string | undefined): string | undefined =>
  sheet?.startsWith("'") ? sheet.slice(1, -1).replace(/''/g, "'") : sheet;

/** "Sheet1!$A$1:$B$5" 같은 참조 텍스트 하나 → RefParts (이름 정의·xl() 재작성용). 아니면 null */
export function parseRefText(text: string): RefParts | null {
  try {
    const toks = tokenize(text.trim());
    return toks.length === 1 && toks[0].t === "ref" ? toks[0].ref! : null;
  } catch {
    return null;
  }
}

// ── 토크나이저 ──────────────────────────────────────────

interface Tok {
  t: "num" | "str" | "err" | "ref" | "id" | "op";
  s: string;
  /** 원문 위치 [i, j) — 참조 재작성용 */
  i: number;
  j: number;
  ref?: RefParts;
}

class ParseErr extends Error {
  constructor(readonly code: FormulaErrorCode) {
    super(code);
  }
}

type PatKind = Tok["t"] | "ws" | "colref" | "rowref";
const NOT_AFTER = "(?![A-Za-z0-9_.(!$])";
const PATTERNS: [PatKind, RegExp][] = [
  ["ws", /\s+/y],
  ["str", /"(?:[^"]|"")*"/y],
  ["err", /#(?:REF!|N\/A|DIV\/0!|VALUE!|NAME\?|NUM!|NULL!|CIRC!)/y],
  // 전체 열 A:C · 전체 행 2:5 (숫자·셀 참조보다 먼저)
  ["colref", new RegExp(`(?:(${SHEET_SRC})!)?(\\$?)([A-Za-z]{1,3}):(\\$?)([A-Za-z]{1,3})${NOT_AFTER}`, "y")],
  ["rowref", new RegExp(`(?:(${SHEET_SRC})!)?(\\$?)([1-9][0-9]{0,6}):(\\$?)([1-9][0-9]{0,6})(?![0-9.A-Za-z])`, "y")],
  ["num", /(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/y],
  // 함수명(LOG10( 등)·시트명(Sheet2!)과 겹치지 않게 뒤를 막는다
  ["ref", new RegExp(`(?:(${SHEET_SRC})!)?(${CELL_SRC})(?::(${CELL_SRC}))?${NOT_AFTER}`, "y")],
  ["id", /[A-Za-z가-힣_][A-Za-z0-9가-힣_.]*/y],
  ["op", /<=|>=|<>|[-+*/^&=<>%(),:!]/y],
];

/**
 * names: 정의된 이름(대문자) → 참조. 주면 이름 토큰을 참조 토큰으로 바꾼다(평가용).
 * 참조 재작성(rewriteRefs)은 names 없이 토큰화하므로 수식 속 이름은 그대로 남는다(엑셀 동일).
 */
function tokenize(src: string, names?: ReadonlyMap<string, RefParts>): Tok[] {
  const toks: Tok[] = [];
  let pos = 0;
  outer: while (pos < src.length) {
    for (const [t, re] of PATTERNS) {
      re.lastIndex = pos;
      const m = re.exec(src);
      if (!m) continue;
      const j = pos + m[0].length;
      pos = j;
      if (t === "ws") continue outer;
      const i = j - m[0].length;
      if (t === "colref" || t === "rowref") {
        const lo = t === "colref" ? letterToCol(m[3]) : Number(m[3]) - 1;
        const hi = t === "colref" ? letterToCol(m[5]) : Number(m[5]) - 1;
        if (Math.max(lo, hi) > (t === "colref" ? MAX_COL : MAX_ROW)) throw new ParseErr("#REF!");
        const [x, y] = lo <= hi ? [[lo, m[2]], [hi, m[4]]] : [[hi, m[4]], [lo, m[2]]];
        const corner = (v: number, abs: string, end: boolean): RefCorner =>
          t === "colref"
            ? { r: end ? MAX_ROW : 0, c: v, ar: true, ac: abs === "$" }
            : { r: v, c: end ? MAX_COL : 0, ar: abs === "$", ac: true };
        toks.push({
          t: "ref",
          s: m[0],
          i,
          j,
          ref: {
            sheetName: unquote(m[1]),
            a: corner(x[0] as number, x[1] as string, false),
            b: corner(y[0] as number, y[1] as string, true),
            kind: t === "colref" ? "cols" : "rows",
          },
        });
        continue outer;
      }
      const tok: Tok = { t, s: m[0], i, j };
      if (t === "ref") {
        const a = parseCorner(m[2]);
        const b = m[3] ? parseCorner(m[3]) : undefined;
        if (!a || b === null) throw new ParseErr("#REF!");
        tok.ref = { sheetName: unquote(m[1]), a, b };
        if (b) {
          // 좌상·우하 정규화 (B5:A1 → A1:B5), $ 표시는 해당 행·열을 따라간다
          if (a.r > b.r) [a.r, a.ar, b.r, b.ar] = [b.r, b.ar, a.r, a.ar];
          if (a.c > b.c) [a.c, a.ac, b.c, b.ac] = [b.c, b.ac, a.c, a.ac];
        }
      } else if (t === "id" && names && !/^\s*[(!]/.test(src.slice(j))) {
        const named = names.get(m[0].toUpperCase());
        if (named) {
          toks.push({ t: "ref", s: m[0], i, j, ref: structuredClone(named) });
          continue outer;
        }
      }
      toks.push(tok);
      continue outer;
    }
    throw new ParseErr("#NAME?"); // 알 수 없는 문자
  }
  return toks;
}

// ── AST ─────────────────────────────────────────────────

/** 평가 중 오류 값 (문자열 결과와 구분하기 위한 상자) */
class FErr {
  constructor(readonly code: string) {}
}

/** null = 빈 셀 */
type Scalar = number | string | boolean | null | FErr;
interface RangeV {
  rng: FormulaRef;
}
type Val = Scalar | RangeV;

type BinOp = "+" | "-" | "*" | "/" | "^" | "&" | "=" | "<>" | "<" | ">" | "<=" | ">=";

type Node =
  | { k: "lit"; v: Scalar }
  | { k: "ref"; ref: FormulaRef; scalar: boolean }
  | { k: "un"; op: "-" | "+" | "%"; a: Node }
  | { k: "bin"; op: BinOp; a: Node; b: Node }
  | { k: "fn"; name: string; args: Node[] };

// ── 파서 (재귀 하강, 엑셀 우선순위: 비교 < & < +- < */ < ^ < 단항 < %) ──

function parse(toks: Tok[], refs: FormulaRef[]): Node {
  let i = 0;
  const isOp = (...ss: string[]) => toks[i]?.t === "op" && ss.includes(toks[i].s);
  const expectOp = (s: string) => {
    if (!isOp(s)) throw new ParseErr("#NAME?");
    i++;
  };

  const binLevel =
    (ops: BinOp[], next: () => Node) =>
    (): Node => {
      let a = next();
      while (isOp(...ops)) {
        const op = toks[i++].s as BinOp;
        a = { k: "bin", op, a, b: next() };
      }
      return a;
    };

  const primary = (): Node => {
    const t = toks[i];
    if (!t) throw new ParseErr("#NAME?");
    i++;
    switch (t.t) {
      case "num":
        return { k: "lit", v: Number(t.s) };
      case "str":
        return { k: "lit", v: t.s.slice(1, -1).replace(/""/g, '"') };
      case "err":
        return { k: "lit", v: new FErr(t.s) };
      case "ref": {
        const ref: FormulaRef = { sheetName: t.ref!.sheetName, range: rangeOf(t.ref!) };
        refs.push(ref);
        return { k: "ref", ref, scalar: !t.ref!.b };
      }
      case "id": {
        const name = t.s.toUpperCase();
        if (isOp("(")) {
          i++;
          if (!(name in FUNCS)) throw new ParseErr("#NAME?");
          const args: Node[] = [];
          if (!isOp(")")) {
            args.push(expr());
            while (isOp(",")) {
              i++;
              args.push(expr());
            }
          }
          expectOp(")");
          return { k: "fn", name, args };
        }
        if (name === "TRUE" || name === "FALSE") return { k: "lit", v: name === "TRUE" };
        // 시트!셀 자리에 셀이 아닌 것, 4글자 이상 열 등 = 참조 오류. 그 외 이름(정의된 이름)은 미지원
        if (isOp("!") || /^[A-Za-z]+\d+$/.test(t.s)) throw new ParseErr("#REF!");
        throw new ParseErr("#NAME?");
      }
      case "op":
        if (t.s === "(") {
          const e = expr();
          expectOp(")");
          return e;
        }
        if (t.s === "-" || t.s === "+") return { k: "un", op: t.s, a: unary() };
        throw new ParseErr("#NAME?");
    }
  };

  const postfix = (): Node => {
    let a = primary();
    while (isOp("%")) {
      i++;
      a = { k: "un", op: "%", a };
    }
    return a;
  };
  const unary = (): Node => {
    if (isOp("-", "+")) {
      const op = toks[i++].s as "-" | "+";
      return { k: "un", op, a: unary() };
    }
    return postfix();
  };
  const pow = binLevel(["^"], unary);
  const mul = binLevel(["*", "/"], pow);
  const add = binLevel(["+", "-"], mul);
  const concat = binLevel(["&"], add);
  const expr: () => Node = binLevel(["=", "<>", "<", ">", "<=", ">="], concat);

  const root = expr();
  if (i !== toks.length) throw new ParseErr("#NAME?");
  return root;
}

// ── 값 변환 (엑셀 강제 변환 규칙) ─────────────────────────

const isRange = (v: Val): v is RangeV => typeof v === "object" && v !== null && "rng" in v;

function cellScalar(cell: Cell | undefined | "#REF!"): Scalar {
  if (cell === "#REF!") return new FErr("#REF!");
  if (cell === undefined || cell.v === null || cell.v === "") return null;
  if (cell.t === "e") return new FErr(String(cell.v));
  return cell.v;
}

function toNum(v: Scalar): number | FErr {
  if (v instanceof FErr || typeof v === "number") return v;
  if (v === null) return 0;
  if (typeof v === "boolean") return v ? 1 : 0;
  const t = v.trim();
  const n = Number(t);
  return t === "" || Number.isNaN(n) ? new FErr("#VALUE!") : n;
}

/** 엑셀 일반 서식처럼 15자리 유효숫자로 — 0.1+0.2 → 0.3 */
const round15 = (n: number): number => Number(n.toPrecision(15));

function toStr(v: Scalar): string | FErr {
  if (v instanceof FErr) return v;
  if (v === null) return "";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "number") return String(round15(v));
  return v;
}

function toBool(v: Scalar): boolean | FErr {
  if (v instanceof FErr || typeof v === "boolean") return v;
  if (v === null) return false;
  if (typeof v === "number") return v !== 0;
  const u = v.toUpperCase();
  return u === "TRUE" ? true : u === "FALSE" ? false : new FErr("#VALUE!");
}

/** 엑셀 비교: 숫자 < 문자 < 불리언, 문자는 대소문자 무시. 빈 셀은 상대 유형의 빈 값 */
function compare(a: Scalar, b: Scalar): number {
  const blank = (x: Scalar, other: Scalar): Scalar =>
    x !== null ? x : typeof other === "string" ? "" : typeof other === "boolean" ? false : 0;
  const x = blank(a, b);
  const y = blank(b, a);
  const rank = (v: Scalar) => (typeof v === "number" ? 0 : typeof v === "string" ? 1 : 2);
  if (rank(x) !== rank(y)) return rank(x) - rank(y);
  if (typeof x === "string" && typeof y === "string") {
    const [p, q] = [x.toLowerCase(), y.toLowerCase()];
    return p < q ? -1 : p > q ? 1 : 0;
  }
  const [p, q] = [Number(x), Number(y)];
  return p < q ? -1 : p > q ? 1 : 0;
}

// ── 평가 ────────────────────────────────────────────────

// ponytail: 범위를 좌표 순회 O(area) — 초대형 범위는 #VALUE!로 거부, 필요해지면 존재 셀 순회로 교체
const MAX_AREA = 1_000_000;

interface Ctx {
  getCell: GetCell;
  bounds?: GetBounds;
}

function evalNode(node: Node, ctx: Ctx): Scalar {
  const v = evalArg(node, ctx);
  if (isRange(v)) {
    // 단일 값 자리의 참조: 1×1이면 그 셀, 아니면 #VALUE! (암시적 교차는 미지원)
    const { range, sheetName } = v.rng;
    if (range.r0 !== range.r1 || range.c0 !== range.c1) return new FErr("#VALUE!");
    return cellScalar(ctx.getCell(sheetName, range.r0, range.c0));
  }
  return v;
}

/** 함수 인수 자리: 참조는 범위 그대로(RangeV), 그 외는 스칼라 */
function evalArg(node: Node, ctx: Ctx): Val {
  switch (node.k) {
    case "lit":
      return node.v;
    case "ref":
      return { rng: node.ref };
    case "un": {
      const a = toNum(evalNode(node.a, ctx));
      if (a instanceof FErr) return a;
      return node.op === "-" ? -a : node.op === "%" ? a / 100 : a;
    }
    case "bin": {
      const a = evalNode(node.a, ctx);
      const b = evalNode(node.b, ctx);
      if (a instanceof FErr) return a;
      if (b instanceof FErr) return b;
      switch (node.op) {
        case "&": {
          const [x, y] = [toStr(a), toStr(b)];
          return x instanceof FErr ? x : y instanceof FErr ? y : x + y;
        }
        case "=":
          return compare(a, b) === 0;
        case "<>":
          return compare(a, b) !== 0;
        case "<":
          return compare(a, b) < 0;
        case ">":
          return compare(a, b) > 0;
        case "<=":
          return compare(a, b) <= 0;
        case ">=":
          return compare(a, b) >= 0;
      }
      const x = toNum(a);
      if (x instanceof FErr) return x;
      const y = toNum(b);
      if (y instanceof FErr) return y;
      switch (node.op) {
        case "+":
          return x + y;
        case "-":
          return x - y;
        case "*":
          return x * y;
        case "/":
          return y === 0 ? new FErr("#DIV/0!") : x / y;
        case "^":
          return x ** y;
      }
      return new FErr("#NAME?"); // 도달 불가
    }
    case "fn":
      return FUNCS[node.name](node.args, ctx);
  }
}

/** 범위 → 2D 스칼라 (시트 없음은 #REF!) */
function grid(v: RangeV, ctx: Ctx): Scalar[][] | FErr {
  const { sheetName } = v.rng;
  let range = v.rng.range;
  const b = ctx.bounds?.(sheetName);
  if (b) {
    // 전체 열/행 참조(A:A)는 시트 사용 크기까지만 읽는다 (최소 1×1 유지)
    const r1 = Math.max(range.r0, Math.min(range.r1, b.rows - 1));
    const c1 = Math.max(range.c0, Math.min(range.c1, b.cols - 1));
    range = { ...range, r1, c1 };
  }
  if ((range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1) > MAX_AREA) return new FErr("#VALUE!");
  const out: Scalar[][] = [];
  for (let r = range.r0; r <= range.r1; r++) {
    const row: Scalar[] = [];
    for (let c = range.c0; c <= range.c1; c++) {
      const cell = ctx.getCell(sheetName, r, c);
      if (cell === "#REF!") return new FErr("#REF!");
      row.push(cellScalar(cell));
    }
    out.push(row);
  }
  return out;
}

/** 인수 → 2D (스칼라는 1×1) */
function asGrid(v: Val, ctx: Ctx): Scalar[][] | FErr {
  if (isRange(v)) return grid(v, ctx);
  return v instanceof FErr ? v : [[v]];
}

/**
 * 집계용 숫자 수집 (엑셀 SUM 규칙): 범위 안은 숫자만(빈 칸·문자·불리언 제외, 오류는 전파),
 * 직접 인수는 숫자로 강제 변환. skipErrors는 COUNT용.
 */
function numbers(args: Node[], ctx: Ctx, skipErrors = false): number[] | FErr {
  const out: number[] = [];
  for (const arg of args) {
    const v = evalArg(arg, ctx);
    if (isRange(v)) {
      const g = grid(v, ctx);
      if (g instanceof FErr) return g;
      for (const row of g)
        for (const x of row) {
          if (typeof x === "number") out.push(x);
          else if (x instanceof FErr && !skipErrors) return x;
        }
    } else {
      const n = toNum(v);
      if (n instanceof FErr) {
        if (skipErrors) continue;
        return n;
      }
      out.push(n);
    }
  }
  return out;
}

type Fn = (args: Node[], ctx: Ctx) => Val;

/** 숫자 목록 → 값 함수 */
const agg =
  (f: (xs: number[]) => Scalar, skipErrors = false): Fn =>
  (args, ctx) => {
    const xs = numbers(args, ctx, skipErrors);
    return xs instanceof FErr ? xs : f(xs);
  };

/** 인수 개수 검사 + 각 인수 숫자 변환 후 f */
const num =
  (min: number, max: number, f: (...xs: number[]) => Scalar): Fn =>
  (args, ctx) => {
    if (args.length < min || args.length > max) return new FErr("#VALUE!");
    const xs: number[] = [];
    for (const a of args) {
      const n = toNum(evalNode(a, ctx));
      if (n instanceof FErr) return n;
      xs.push(n);
    }
    return f(...xs);
  };

const text =
  (min: number, max: number, f: (s: string, ...xs: number[]) => Scalar): Fn =>
  (args, ctx) => {
    if (args.length < min || args.length > max) return new FErr("#VALUE!");
    const s = toStr(evalNode(args[0], ctx));
    if (s instanceof FErr) return s;
    const xs: number[] = [];
    for (const a of args.slice(1)) {
      const n = toNum(evalNode(a, ctx));
      if (n instanceof FErr) return n;
      xs.push(n);
    }
    return f(s, ...xs);
  };

const sum = (xs: number[]) => xs.reduce((a, b) => a + b, 0);
const mean = (xs: number[]) => (xs.length === 0 ? new FErr("#DIV/0!") : sum(xs) / xs.length);
const variance = (xs: number[], sample: boolean): number | FErr => {
  const n = xs.length;
  if (n < (sample ? 2 : 1)) return new FErr("#DIV/0!");
  const m = sum(xs) / n;
  return sum(xs.map((x) => (x - m) ** 2)) / (sample ? n - 1 : n);
};
const sqrtOf = (v: number | FErr) => (v instanceof FErr ? v : Math.sqrt(v));

/** 소수 자리 반올림 (엑셀: 0.5는 0에서 먼 쪽) */
const roundTo = (x: number, d: number, mode: "round" | "up" | "down"): number => {
  const p = 10 ** Math.trunc(d);
  const y = Math.abs(round15(x * p));
  const r = mode === "round" ? Math.round(y) : mode === "up" ? Math.ceil(y) : Math.floor(y);
  return (Math.sign(x) * r) / p;
};

/** COUNTIF류 조건: 5, ">=10", "<>a", "a*" (와일드카드 * ?) */
function criterion(c: Scalar): (x: Scalar) => boolean {
  if (c instanceof FErr) return () => false;
  if (typeof c !== "string") return (x) => x !== null && compare(x, c) === 0;
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(c)!;
  const op = m[1] ?? "=";
  const rhsText = m[2];
  const n = rhsText.trim() === "" ? NaN : Number(rhsText);
  if (rhsText === "") {
    // "" 또는 "=" → 빈 칸, "<>" → 비어 있지 않음
    return op === "<>" ? (x) => x !== null && x !== "" : (x) => x === null || x === "";
  }
  if (!Number.isNaN(n)) {
    const test = (d: number) =>
      op === "=" ? d === 0 : op === "<>" ? d !== 0 : op === "<" ? d < 0 : op === ">" ? d > 0 : op === "<=" ? d <= 0 : d >= 0;
    return (x) => (typeof x === "number" ? test(compare(x, n)) : op === "<>");
  }
  if (op === "=" || op === "<>") {
    const re = new RegExp(
      "^" + rhsText.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".") + "$",
      "is",
    );
    return (x) => (typeof x === "string" && re.test(x)) === (op === "=");
  }
  return (x) => {
    if (typeof x !== "string") return false;
    const d = compare(x, rhsText);
    return op === "<" ? d < 0 : op === ">" ? d > 0 : op === "<=" ? d <= 0 : d >= 0;
  };
}

/** SUMIFS(합계범위, 조건범위1, 조건1, …) / COUNTIFS(조건범위1, 조건1, …) 공통 — 일치 셀 위치 마스크 */
function ifsMask(pairs: Node[], ctx: Ctx): { rows: number; cols: number; hit: boolean[][] } | FErr {
  if (pairs.length === 0 || pairs.length % 2 !== 0) return new FErr("#VALUE!");
  let hit: boolean[][] | null = null;
  let rows = 0;
  let cols = 0;
  for (let k = 0; k < pairs.length; k += 2) {
    const g = asGrid(evalArg(pairs[k], ctx), ctx);
    if (g instanceof FErr) return g;
    const crit = evalNode(pairs[k + 1], ctx);
    if (crit instanceof FErr) return crit;
    const test = criterion(crit);
    if (hit === null) {
      rows = g.length;
      cols = g[0].length;
      hit = g.map((row) => row.map(() => true));
    } else if (g.length !== rows || g[0].length !== cols) return new FErr("#VALUE!");
    g.forEach((row, r) => row.forEach((x, c) => (hit![r][c] &&= test(x))));
  }
  return { rows, cols, hit: hit! };
}

/** 합계 범위를 조건 범위 크기로 (엑셀: 합계 범위 좌상단 기준으로 크기를 맞춘다) */
function sized(node: Node, rows: number, cols: number, ctx: Ctx): Scalar[][] | FErr {
  const v = evalArg(node, ctx);
  if (!isRange(v)) return new FErr("#VALUE!");
  const { r0, c0 } = v.rng.range;
  return grid({ rng: { sheetName: v.rng.sheetName, range: { r0, c0, r1: r0 + rows - 1, c1: c0 + cols - 1 } } }, ctx);
}

function sumIfs(sumNode: Node, pairs: Node[], ctx: Ctx, avg: boolean): Scalar {
  const m = ifsMask(pairs, ctx);
  if (m instanceof FErr) return m;
  const g = sized(sumNode, m.rows, m.cols, ctx);
  if (g instanceof FErr) return g;
  const xs: number[] = [];
  for (let r = 0; r < m.rows; r++)
    for (let c = 0; c < m.cols; c++) {
      const x = g[r][c];
      if (!m.hit[r][c]) continue;
      if (x instanceof FErr) return x;
      if (typeof x === "number") xs.push(x);
    }
  return avg ? mean(xs) : sum(xs);
}

/** 1차원으로 펼침 (MATCH용 — 행 또는 열 하나여야 한다) */
function vector(v: Val, ctx: Ctx): Scalar[] | FErr {
  const g = asGrid(v, ctx);
  if (g instanceof FErr) return g;
  if (g.length === 1) return g[0];
  if (g[0].length === 1) return g.map((row) => row[0]);
  return new FErr("#N/A");
}

/** MATCH 위치(0-based) — type 0 정확(와일드카드 지원), 1 이하 최대값, -1 이상 최소값 (정렬 가정) */
function matchIndex(x: Scalar, xs: Scalar[], type: number): number {
  if (type === 0) {
    // 문자는 "=" 접두로 연산자 해석을 막고 와일드카드만 살린다
    const test = typeof x === "string" ? criterion("=" + x) : (y: Scalar) => y !== null && compare(y, x) === 0;
    return xs.findIndex((y) => !(y instanceof FErr) && test(y));
  }
  let found = -1;
  for (let k = 0; k < xs.length; k++) {
    const y = xs[k];
    if (y === null || y instanceof FErr) continue;
    const d = compare(y, x);
    if (type > 0 ? d <= 0 : d >= 0) found = k;
    else break;
  }
  return found;
}

function lookup(args: Node[], ctx: Ctx, horizontal: boolean): Scalar {
  if (args.length < 3 || args.length > 4) return new FErr("#VALUE!");
  const x = evalNode(args[0], ctx);
  if (x instanceof FErr) return x;
  let g = asGrid(evalArg(args[1], ctx), ctx);
  if (g instanceof FErr) return g;
  if (horizontal) g = g[0].map((_, c) => (g as Scalar[][]).map((row) => row[c]));
  const idx = toNum(evalNode(args[2], ctx));
  if (idx instanceof FErr) return idx;
  const approx = args.length === 4 ? toBool(evalNode(args[3], ctx)) : true;
  if (approx instanceof FErr) return approx;
  if (idx < 1) return new FErr("#VALUE!");
  if (idx > g[0].length) return new FErr("#REF!");
  const k = matchIndex(x, g.map((row) => row[0]), approx ? 1 : 0);
  return k < 0 ? new FErr("#N/A") : g[k][Math.trunc(idx) - 1];
}

const FUNCS: Record<string, Fn> = {
  // 집계
  SUM: agg(sum),
  AVERAGE: agg(mean),
  MIN: agg((xs) => (xs.length === 0 ? 0 : Math.min(...xs))),
  MAX: agg((xs) => (xs.length === 0 ? 0 : Math.max(...xs))),
  COUNT: agg((xs) => xs.length, true),
  PRODUCT: agg((xs) => xs.reduce((a, b) => a * b, 1)),
  MEDIAN: agg((xs) => {
    if (xs.length === 0) return new FErr("#NUM!");
    const s = [...xs].sort((a, b) => a - b);
    const h = s.length >> 1;
    return s.length % 2 ? s[h] : (s[h - 1] + s[h]) / 2;
  }),
  VAR: agg((xs) => variance(xs, true)),
  "VAR.S": agg((xs) => variance(xs, true)),
  "VAR.P": agg((xs) => variance(xs, false)),
  STDEV: agg((xs) => sqrtOf(variance(xs, true))),
  "STDEV.S": agg((xs) => sqrtOf(variance(xs, true))),
  "STDEV.P": agg((xs) => sqrtOf(variance(xs, false))),
  COUNTA: (args, ctx) => {
    let n = 0;
    for (const a of args) {
      const g = asGrid(evalArg(a, ctx), ctx);
      if (g instanceof FErr) n++; // 오류 값도 비어 있지 않은 값
      else for (const row of g) for (const x of row) if (x !== null) n++;
    }
    return n;
  },
  SUMPRODUCT: (args, ctx) => {
    let acc: number[][] | null = null;
    for (const a of args) {
      const g = asGrid(evalArg(a, ctx), ctx);
      if (g instanceof FErr) return g;
      if (acc && (g.length !== acc.length || g[0].length !== acc[0].length)) return new FErr("#VALUE!");
      const cur: number[][] = acc ?? g.map((row) => row.map(() => 1));
      for (let r = 0; r < g.length; r++)
        for (let c = 0; c < g[0].length; c++) {
          const x = g[r][c];
          if (x instanceof FErr) return x;
          cur[r][c] *= typeof x === "number" ? x : 0; // 비숫자는 0 (엑셀 동일)
        }
      acc = cur;
    }
    return acc ? sum(acc.flat()) : new FErr("#VALUE!");
  },
  // 조건 집계
  COUNTIF: (args, ctx) => {
    if (args.length !== 2) return new FErr("#VALUE!");
    const m = ifsMask(args, ctx);
    return m instanceof FErr ? m : m.hit.flat().filter(Boolean).length;
  },
  COUNTIFS: (args, ctx) => {
    const m = ifsMask(args, ctx);
    return m instanceof FErr ? m : m.hit.flat().filter(Boolean).length;
  },
  SUMIF: (args, ctx) =>
    args.length < 2 || args.length > 3
      ? new FErr("#VALUE!")
      : sumIfs(args[2] ?? args[0], args.slice(0, 2), ctx, false),
  AVERAGEIF: (args, ctx) =>
    args.length < 2 || args.length > 3
      ? new FErr("#VALUE!")
      : sumIfs(args[2] ?? args[0], args.slice(0, 2), ctx, true),
  SUMIFS: (args, ctx) => (args.length < 3 ? new FErr("#VALUE!") : sumIfs(args[0], args.slice(1), ctx, false)),
  AVERAGEIFS: (args, ctx) => (args.length < 3 ? new FErr("#VALUE!") : sumIfs(args[0], args.slice(1), ctx, true)),
  // 논리 (IF·IFERROR는 고르지 않은 쪽을 평가하지 않는다)
  IF: (args, ctx) => {
    if (args.length < 2 || args.length > 3) return new FErr("#VALUE!");
    const c = toBool(evalNode(args[0], ctx));
    if (c instanceof FErr) return c;
    if (c) return evalArg(args[1], ctx);
    return args.length === 3 ? evalArg(args[2], ctx) : false;
  },
  IFERROR: (args, ctx) => {
    if (args.length !== 2) return new FErr("#VALUE!");
    const v = evalNode(args[0], ctx);
    return v instanceof FErr ? evalArg(args[1], ctx) : v;
  },
  IFNA: (args, ctx) => {
    if (args.length !== 2) return new FErr("#VALUE!");
    const v = evalNode(args[0], ctx);
    return v instanceof FErr && v.code === "#N/A" ? evalArg(args[1], ctx) : v;
  },
  AND: (args, ctx) => logical(args, ctx, (bs) => bs.every(Boolean)),
  OR: (args, ctx) => logical(args, ctx, (bs) => bs.some(Boolean)),
  NOT: (args, ctx) => {
    if (args.length !== 1) return new FErr("#VALUE!");
    const b = toBool(evalNode(args[0], ctx));
    return b instanceof FErr ? b : !b;
  },
  ISBLANK: (args, ctx) => args.length === 1 && evalNode(args[0], ctx) === null,
  ISNUMBER: (args, ctx) => args.length === 1 && typeof evalNode(args[0], ctx) === "number",
  ISTEXT: (args, ctx) => args.length === 1 && typeof evalNode(args[0], ctx) === "string",
  ISERROR: (args, ctx) => args.length === 1 && evalNode(args[0], ctx) instanceof FErr,
  // 수학
  ABS: num(1, 1, Math.abs),
  INT: num(1, 1, Math.floor),
  TRUNC: num(1, 2, (x, d = 0) => roundTo(x, d, "down")),
  ROUND: num(2, 2, (x, d) => roundTo(x, d, "round")),
  ROUNDUP: num(2, 2, (x, d) => roundTo(x, d, "up")),
  ROUNDDOWN: num(2, 2, (x, d) => roundTo(x, d, "down")),
  MOD: num(2, 2, (x, y) => (y === 0 ? new FErr("#DIV/0!") : x - y * Math.floor(x / y))),
  POWER: num(2, 2, (x, y) => x ** y),
  SQRT: num(1, 1, (x) => (x < 0 ? new FErr("#NUM!") : Math.sqrt(x))),
  EXP: num(1, 1, Math.exp),
  LN: num(1, 1, (x) => (x <= 0 ? new FErr("#NUM!") : Math.log(x))),
  LOG10: num(1, 1, (x) => (x <= 0 ? new FErr("#NUM!") : Math.log10(x))),
  LOG: num(1, 2, (x, b = 10) => (x <= 0 || b <= 0 || b === 1 ? new FErr("#NUM!") : Math.log(x) / Math.log(b))),
  SIGN: num(1, 1, Math.sign),
  PI: num(0, 0, () => Math.PI),
  // 문자
  LEN: text(1, 1, (s) => s.length),
  LEFT: text(1, 2, (s, n = 1) => (n < 0 ? new FErr("#VALUE!") : s.slice(0, n))),
  RIGHT: text(1, 2, (s, n = 1) => (n < 0 ? new FErr("#VALUE!") : n === 0 ? "" : s.slice(-n))),
  MID: text(3, 3, (s, start, n) => (start < 1 || n < 0 ? new FErr("#VALUE!") : s.substr(start - 1, n))),
  UPPER: text(1, 1, (s) => s.toUpperCase()),
  LOWER: text(1, 1, (s) => s.toLowerCase()),
  TRIM: text(1, 1, (s) => s.trim().replace(/ +/g, " ")),
  VALUE: text(1, 1, (s) => toNum(s)),
  CONCAT: (args, ctx) => concatAll(args, ctx),
  CONCATENATE: (args, ctx) => concatAll(args, ctx),
  // 찾기
  VLOOKUP: (args, ctx) => lookup(args, ctx, false),
  HLOOKUP: (args, ctx) => lookup(args, ctx, true),
  MATCH: (args, ctx) => {
    if (args.length < 2 || args.length > 3) return new FErr("#VALUE!");
    const x = evalNode(args[0], ctx);
    if (x instanceof FErr) return x;
    const xs = vector(evalArg(args[1], ctx), ctx);
    if (xs instanceof FErr) return xs;
    const type = args.length === 3 ? toNum(evalNode(args[2], ctx)) : 1;
    if (type instanceof FErr) return type;
    const k = matchIndex(x, xs, Math.sign(type));
    return k < 0 ? new FErr("#N/A") : k + 1;
  },
  INDEX: (args, ctx) => {
    if (args.length < 2 || args.length > 3) return new FErr("#VALUE!");
    const g = asGrid(evalArg(args[0], ctx), ctx);
    if (g instanceof FErr) return g;
    const a = toNum(evalNode(args[1], ctx));
    if (a instanceof FErr) return a;
    const b = args.length === 3 ? toNum(evalNode(args[2], ctx)) : undefined;
    if (b instanceof FErr) return b;
    // 한 행짜리 범위에 인수 하나면 열 번호로 본다 (엑셀 동일)
    const [r, c] = b === undefined ? (g.length === 1 ? [1, a] : [a, 1]) : [a, b];
    if (r < 1 || c < 1) return new FErr("#VALUE!"); // 0(행/열 전체)은 미지원
    const row = g[Math.trunc(r) - 1];
    return row && Math.trunc(c) <= row.length ? row[Math.trunc(c) - 1] : new FErr("#REF!");
  },
};

function logical(args: Node[], ctx: Ctx, f: (bs: boolean[]) => boolean): Scalar {
  const bs: boolean[] = [];
  for (const a of args) {
    const v = evalArg(a, ctx);
    const g = asGrid(v, ctx);
    if (g instanceof FErr) return g;
    for (const row of g)
      for (const x of row) {
        if (x === null || (typeof x === "string" && isRange(v))) continue; // 범위 안의 빈칸·문자 무시
        const b = toBool(x);
        if (b instanceof FErr) return b;
        bs.push(b);
      }
  }
  return bs.length === 0 ? new FErr("#VALUE!") : f(bs);
}

function concatAll(args: Node[], ctx: Ctx): Scalar {
  let s = "";
  for (const a of args) {
    const g = asGrid(evalArg(a, ctx), ctx);
    if (g instanceof FErr) return g;
    for (const row of g)
      for (const x of row) {
        const t = toStr(x);
        if (t instanceof FErr) return t;
        s += t;
      }
  }
  return s;
}

function toResult(v: Scalar): FormulaResult {
  if (v instanceof FErr) return { v: v.code, t: "e" };
  if (v === null) return { v: 0, t: "n" }; // =A1 (빈 셀) → 0, 엑셀 동일
  if (typeof v === "number") {
    return Number.isFinite(v) ? { v: round15(v), t: "n" } : { v: "#NUM!", t: "e" };
  }
  return typeof v === "boolean" ? { v, t: "b" } : { v, t: "s" };
}

/**
 * 수식 파싱. 실패해도 throw하지 않는다 — eval이 해당 오류 코드를 돌려주는
 * 상수 수식(refs 없음)이 된다. 함수명은 대소문자 무관.
 */
export function parseFormula(src: string, names?: ReadonlyMap<string, RefParts>): ParsedFormula {
  const body = src.trim().replace(/^=/, "");
  const refs: FormulaRef[] = [];
  try {
    const ast = parse(tokenize(body, names), refs);
    return { refs, eval: (getCell, bounds) => toResult(evalNode(ast, { getCell, bounds })) };
  } catch (e) {
    const code = e instanceof ParseErr ? e.code : "#NAME?";
    return { refs: [], error: code, eval: () => ({ v: code, t: "e" }) };
  }
}

/** 정의된 이름 목록 → 토크나이저용 맵 (대문자 키). 참조 텍스트가 잘못된 이름은 빠진다 */
export function namesMap(names: readonly { name: string; ref: string }[] | undefined): Map<string, RefParts> {
  const m = new Map<string, RefParts>();
  for (const n of names ?? []) {
    const p = parseRefText(n.ref);
    if (p) m.set(n.name.toUpperCase(), p);
  }
  return m;
}

/** 지원 함수 이름 (자동완성·문서용) */
export const FORMULA_FUNCTIONS = Object.keys(FUNCS).sort();

// ── 참조 재작성 (부록 O.2) ────────────────────────────────

/**
 * 수식 안의 참조 토큰마다 map을 적용해 새 수식 문자열을 만든다.
 * map이 null을 돌려주면 그 참조는 `#REF!`가 된다. 바뀌지 않은 참조는 원문 그대로 둔다.
 * 토큰화에 실패한 수식(이미 #NAME? 상태)은 그대로 돌려준다.
 */
export function rewriteRefs(fx: string, map: (p: RefParts) => RefParts | null): string {
  let toks: Tok[];
  try {
    toks = tokenize(fx);
  } catch {
    return fx;
  }
  let out = "";
  let last = 0;
  for (const t of toks) {
    if (t.t !== "ref") continue;
    const before = JSON.stringify(t.ref);
    const next = map(structuredClone(t.ref!));
    const replaced = next === null ? "#REF!" : JSON.stringify(next) === before ? t.s : formatRef(next);
    out += fx.slice(last, t.i) + replaced;
    last = t.j;
  }
  return out + fx.slice(last);
}

const inBounds = (k: RefCorner) => k.r >= 0 && k.c >= 0 && k.r <= MAX_ROW && k.c <= MAX_COL;

/**
 * 복사·붙여넣기/채우기: 상대 참조만 (dr, dc)만큼 이동, `$` 고정 부분은 그대로.
 * 시트 밖으로 밀려나는 참조는 #REF! (엑셀 동일).
 */
export function shiftFormula(fx: string, dr: number, dc: number): string {
  if (dr === 0 && dc === 0) return fx;
  return rewriteRefs(fx, (p) => {
    for (const k of p.b ? [p.a, p.b] : [p.a]) {
      if (!k.ar) k.r += dr;
      if (!k.ac) k.c += dc;
      if (!inBounds(k)) return null;
    }
    return p;
  });
}

export interface StructureEdit {
  /** 행·열이 삽입/삭제된 시트 이름 */
  sheetName: string;
  axis: "row" | "col";
  index: number;
  /** 양수 = index 앞에 삽입, 음수 = index부터 |count|개 삭제 */
  count: number;
}

/**
 * 행/열 삽입·삭제에 따른 참조 조정 (절대·상대 무관 — 구조 변경은 $도 따라 움직인다).
 * ownSheetName은 수식이 놓인 시트(접두어 없는 참조의 시트). 범위 안쪽 삽입은 범위를 넓히고,
 * 범위 일부 삭제는 줄이며, 참조 셀이 통째로 지워지면 #REF!.
 */
export function adjustForStructure(fx: string, ownSheetName: string, ed: StructureEdit): string {
  if (ed.count === 0) return fx;
  const key = ed.axis === "row" ? "r" : "c";
  return rewriteRefs(fx, (p) => {
    if ((p.sheetName ?? ownSheetName) !== ed.sheetName) return p;
    // 전체 열 참조는 행 삽입·삭제와 무관 (전체 행 참조와 열도 마찬가지)
    if ((p.kind === "cols" && key === "r") || (p.kind === "rows" && key === "c")) return p;
    const lo = p.a;
    const hi = p.b ?? p.a;
    if (ed.count > 0) {
      for (const k of p.b ? [lo, hi] : [lo]) if (k[key] >= ed.index) k[key] += ed.count;
      return inBounds(hi) ? p : null;
    }
    const n = -ed.count;
    const end = ed.index + n; // 삭제 구간 [index, end)
    const nLo = lo[key] < ed.index ? lo[key] : lo[key] < end ? ed.index : lo[key] - n;
    const nHi = hi[key] < ed.index ? hi[key] : hi[key] < end ? ed.index - 1 : hi[key] - n;
    if (nHi < nLo) return null;
    lo[key] = nLo;
    hi[key] = nHi;
    return p;
  });
}

/** 시트 이름 변경 시 접두어 참조 갱신 (엑셀 동일 — 이름이 바뀌어도 수식이 끊기지 않는다) */
export const renameSheetInFormula = (fx: string, from: string, to: string): string =>
  rewriteRefs(fx, (p) => (p.sheetName === from ? { ...p, sheetName: to } : p));

/**
 * 잘라내기 이동(부록 O.4): 원본 사각형 **안에 완전히 들어가는** 참조만 (dr, dc)만큼 옮긴다($ 무관).
 * 엑셀처럼 이동한 셀을 가리키던 모든 수식이 새 위치를 따라가고, 이동한 수식 자신의 참조는 그대로다.
 * 같은 시트 안 이동 기준 — srcSheetName은 원본 시트, ownSheetName은 수식이 놓인 시트.
 */
export function moveRefsInFormula(
  fx: string,
  ownSheetName: string,
  srcSheetName: string,
  src: CellRange,
  dr: number,
  dc: number,
): string {
  return rewriteRefs(fx, (p) => {
    if (p.kind || (p.sheetName ?? ownSheetName) !== srcSheetName) return p;
    const rg = rangeOf(p);
    if (rg.r0 < src.r0 || rg.r1 > src.r1 || rg.c0 < src.c0 || rg.c1 > src.c1) return p;
    for (const k of p.b ? [p.a, p.b] : [p.a]) {
      k.r += dr;
      k.c += dc;
      if (!inBounds(k)) return null;
    }
    return p;
  });
}

/** 셀에 직접 입력한 텍스트 → 셀 (`=수식` → fx, 값은 재계산이 채운다 / 숫자 / TRUE·FALSE / 문자열) */
export function cellFromInput(text: string): Cell | null {
  const trimmed = text.trim();
  if (trimmed === "") return null;
  if (isFormula(trimmed)) return { v: null, t: "n", fx: trimmed };
  if (/^(true|false)$/i.test(trimmed)) return { v: trimmed.toLowerCase() === "true", t: "b" };
  const n = Number(trimmed);
  if (!Number.isNaN(n)) return { v: n, t: "n" };
  return { v: text, t: "s" };
}

/** Python 코드의 `sheet("…")`·`xl('…')`(호환 별칭) 문자열 리터럴 참조 (부록 O.4 — 행/열 삽입·시트 이름 변경 재작성) */
const XL_CALL_RE = /\b(?:sheet|xl)\(\s*([rRuU]?)(["'])((?:(?!\2)[^\\n])*)\2/g;

/** 코드 속 xl() 참조 텍스트마다 map 적용. 같은 따옴표가 결과에 들어가면 원문 유지 */
export function rewriteXlRefs(code: string, map: (ref: string) => string): string {
  return code.replace(XL_CALL_RE, (all: string, _pre: string, q: string, ref: string) => {
    const next = map(ref);
    if (next === ref || next.includes(q)) return all;
    return all.slice(0, all.length - ref.length - 1) + next + q;
  });
}

/** 함수 도움말 (수식 입력줄 자동완성 — 부록 O.4). 키는 FUNCS와 같아야 한다(단위 테스트가 확인) */
export const FUNCTION_HELP: Record<string, string> = {
  SUM: "SUM(수1, [수2], …) — 합계",
  AVERAGE: "AVERAGE(수1, [수2], …) — 평균",
  MIN: "MIN(수1, [수2], …) — 최솟값",
  MAX: "MAX(수1, [수2], …) — 최댓값",
  COUNT: "COUNT(값1, …) — 숫자 셀 개수",
  COUNTA: "COUNTA(값1, …) — 비어 있지 않은 셀 개수",
  PRODUCT: "PRODUCT(수1, …) — 곱",
  MEDIAN: "MEDIAN(수1, …) — 중앙값",
  VAR: "VAR(수1, …) — 표본분산",
  "VAR.S": "VAR.S(수1, …) — 표본분산",
  "VAR.P": "VAR.P(수1, …) — 모분산",
  STDEV: "STDEV(수1, …) — 표본표준편차",
  "STDEV.S": "STDEV.S(수1, …) — 표본표준편차",
  "STDEV.P": "STDEV.P(수1, …) — 모표준편차",
  SUMPRODUCT: "SUMPRODUCT(배열1, [배열2], …) — 곱의 합",
  COUNTIF: "COUNTIF(범위, 조건) — 조건에 맞는 셀 개수",
  COUNTIFS: "COUNTIFS(범위1, 조건1, …) — 여러 조건 개수",
  SUMIF: "SUMIF(범위, 조건, [합계범위]) — 조건부 합계",
  SUMIFS: "SUMIFS(합계범위, 범위1, 조건1, …) — 여러 조건 합계",
  AVERAGEIF: "AVERAGEIF(범위, 조건, [평균범위]) — 조건부 평균",
  AVERAGEIFS: "AVERAGEIFS(평균범위, 범위1, 조건1, …) — 여러 조건 평균",
  IF: "IF(조건, 참일 때, [거짓일 때]) — 조건 분기",
  IFERROR: "IFERROR(값, 오류일 때) — 오류 대체",
  IFNA: "IFNA(값, #N/A일 때) — #N/A 대체",
  AND: "AND(논리1, …) — 모두 참",
  OR: "OR(논리1, …) — 하나라도 참",
  NOT: "NOT(논리) — 부정",
  ISBLANK: "ISBLANK(값) — 빈 셀인가",
  ISNUMBER: "ISNUMBER(값) — 숫자인가",
  ISTEXT: "ISTEXT(값) — 문자인가",
  ISERROR: "ISERROR(값) — 오류인가",
  ABS: "ABS(수) — 절댓값",
  INT: "INT(수) — 내림 정수",
  TRUNC: "TRUNC(수, [자릿수]) — 버림",
  ROUND: "ROUND(수, 자릿수) — 반올림",
  ROUNDUP: "ROUNDUP(수, 자릿수) — 올림",
  ROUNDDOWN: "ROUNDDOWN(수, 자릿수) — 내림",
  MOD: "MOD(수, 제수) — 나머지",
  POWER: "POWER(수, 지수) — 거듭제곱",
  SQRT: "SQRT(수) — 제곱근",
  EXP: "EXP(수) — e의 거듭제곱",
  LN: "LN(수) — 자연로그",
  LOG: "LOG(수, [밑=10]) — 로그",
  LOG10: "LOG10(수) — 상용로그",
  SIGN: "SIGN(수) — 부호",
  PI: "PI() — 원주율",
  LEN: "LEN(문자) — 글자 수",
  LEFT: "LEFT(문자, [개수]) — 왼쪽 글자",
  RIGHT: "RIGHT(문자, [개수]) — 오른쪽 글자",
  MID: "MID(문자, 시작, 개수) — 중간 글자",
  UPPER: "UPPER(문자) — 대문자",
  LOWER: "LOWER(문자) — 소문자",
  TRIM: "TRIM(문자) — 공백 정리",
  VALUE: "VALUE(문자) — 숫자로",
  CONCAT: "CONCAT(문자1, …) — 이어 붙이기",
  CONCATENATE: "CONCATENATE(문자1, …) — 이어 붙이기",
  VLOOKUP: "VLOOKUP(찾을값, 표, 열번호, [근사=TRUE]) — 세로 찾기",
  HLOOKUP: "HLOOKUP(찾을값, 표, 행번호, [근사=TRUE]) — 가로 찾기",
  MATCH: "MATCH(찾을값, 범위, [유형=1]) — 위치",
  INDEX: "INDEX(범위, 행, [열]) — 위치의 값",
};
