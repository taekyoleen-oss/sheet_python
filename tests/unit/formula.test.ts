import { describe, expect, it } from "vitest";
import {
  FORMULA_ERROR_KO,
  isFormula,
  isFormulaError,
  parseFormula,
  type GetCell,
} from "@/lib/grid/formula";
import type { Cell } from "@/types/workbook";

/** 키 "시트!r:c" 셀 맵 → GetCell. 시트 목록에 없는 이름은 #REF! */
const grid =
  (cells: Record<string, Cell>, sheets: string[] = ["Sheet1"]): GetCell =>
  (sheetName, r, c) => {
    const name = sheetName ?? "Sheet1";
    if (!sheets.includes(name)) return "#REF!";
    return cells[`${name}!${r}:${c}`];
  };

const empty = grid({});
const ev = (src: string, getCell: GetCell = empty) => parseFormula(src).eval(getCell).v;

describe("isFormula", () => {
  it("= 시작만 수식, = 단독·공백은 제외", () => {
    expect(isFormula("=A1")).toBe(true);
    expect(isFormula("  =A1+1")).toBe(true);
    expect(isFormula("=")).toBe(false);
    expect(isFormula("A1")).toBe(false);
    expect(isFormula("1+1")).toBe(false);
  });
});

describe("파서 — 우선순위·괄호·단항", () => {
  it("4칙연산 우선순위", () => {
    expect(ev("=1+2*3")).toBe(7);
    expect(ev("=(1+2)*3")).toBe(9);
    expect(ev("=2*3-4/2")).toBe(4);
    expect(ev("=10-2-3")).toBe(5); // 좌결합
  });

  it("단항 마이너스", () => {
    expect(ev("=-3+5")).toBe(2);
    expect(ev("=-(2+3)")).toBe(-5);
    expect(ev("=2*-3")).toBe(-6);
  });

  it("소수", () => {
    expect(ev("=1.5*2")).toBe(3);
    expect(ev("=0.5+.5")).toBe(1);
  });
});

describe("셀·시트 참조", () => {
  const g = grid(
    {
      "Sheet1!0:0": { v: 2, t: "n" }, // A1
      "Sheet2!2:1": { v: 10, t: "n" }, // Sheet2!B3
      "시트 이름!1:2": { v: 7, t: "n" }, // '시트 이름'!C2
    },
    ["Sheet1", "Sheet2", "시트 이름"],
  );

  it("A1·Sheet2!B3·'시트 이름'!C2", () => {
    expect(ev("=A1*3", g)).toBe(6);
    expect(ev("=Sheet2!B3+1", g)).toBe(11);
    expect(ev("='시트 이름'!C2*2", g)).toBe(14);
    expect(ev("=A1+Sheet2!B3+'시트 이름'!C2", g)).toBe(19);
  });

  it("빈 셀 참조 = 0", () => {
    expect(ev("=A9+1", g)).toBe(1);
    expect(ev("=-Z99", g)).toBe(0);
  });

  it("refs에 참조가 수집된다 (시트 이름 포함)", () => {
    const { refs } = parseFormula("=A1+SUM(Sheet2!B1:B3)");
    expect(refs).toEqual([
      { sheetName: undefined, range: { r0: 0, c0: 0, r1: 0, c1: 0 } },
      { sheetName: "Sheet2", range: { r0: 0, c0: 1, r1: 2, c1: 1 } },
    ]);
  });
});

describe("집계 함수 5종", () => {
  const g = grid({
    "Sheet1!0:0": { v: 1, t: "n" }, // A1
    "Sheet1!1:0": { v: 2, t: "n" }, // A2
    "Sheet1!2:0": { v: "텍스트", t: "s" }, // A3 — 집계 제외
    "Sheet1!4:0": { v: 4, t: "n" }, // A5 (A4 빈 칸)
    "Sheet1!0:1": { v: 5, t: "n" }, // B1
  });

  it("SUM·AVERAGE·MIN·MAX·COUNT — 빈 칸·비숫자 제외", () => {
    expect(ev("=SUM(A1:A5)", g)).toBe(7);
    expect(ev("=COUNT(A1:A5)", g)).toBe(3); // 숫자만
    expect(ev("=AVERAGE(A1:A5)", g)).toBeCloseTo(7 / 3);
    expect(ev("=MIN(A1:A5)", g)).toBe(1);
    expect(ev("=MAX(A1:A5)", g)).toBe(4);
  });

  it("범위·셀·숫자·식 혼합 인수", () => {
    expect(ev("=SUM(A1:A2, 10, B1, 1+1)", g)).toBe(20);
    expect(ev("=MAX(A1:A2, B1*2)", g)).toBe(10);
    expect(ev("=COUNT(A1:A5, 99)", g)).toBe(4);
  });

  it("대소문자 무관 함수명", () => {
    expect(ev("=sum(a1:a2)", g)).toBe(3);
    expect(ev("=Average(A1:A2)", g)).toBe(1.5);
  });

  it("빈 집계: SUM=0·COUNT=0·MIN/MAX=0·AVERAGE=#DIV/0!", () => {
    expect(ev("=SUM(C1:C5)", g)).toBe(0);
    expect(ev("=COUNT(C1:C5)", g)).toBe(0);
    expect(ev("=MIN(C1:C5)", g)).toBe(0);
    expect(ev("=MAX(C1:C5)", g)).toBe(0);
    expect(ev("=AVERAGE(C1:C5)", g)).toBe("#DIV/0!");
  });
});

describe("오류", () => {
  it("#NAME? — 미지원 함수·문법", () => {
    expect(ev("=FOO(1)")).toBe("#NAME?");
    expect(ev("=1+")).toBe("#NAME?");
    expect(ev("=정의된이름")).toBe("#NAME?");
  });

  it("#REF! — 파싱 불가 참조·시트 없음", () => {
    expect(ev("=ZZZZ1")).toBe("#REF!"); // 열 문자 4자리
    expect(ev("=Sheet2!5A")).toBe("#REF!"); // 셀 자리에 셀이 아닌 토큰
    expect(ev("=Sheet9!A1", empty)).toBe("#REF!"); // getCell이 시트를 못 찾음
    expect(ev("=SUM(Sheet9!A1:A3)", empty)).toBe("#REF!");
  });

  it("#VALUE! — 비숫자 셀 연산·스칼라 자리의 범위", () => {
    const g = grid({ "Sheet1!0:0": { v: "abc", t: "s" } });
    expect(ev("=A1+1", g)).toBe("#VALUE!");
    expect(ev("=A1:B2+1", g)).toBe("#VALUE!");
  });

  it("#DIV/0!", () => {
    expect(ev("=1/0")).toBe("#DIV/0!");
    expect(ev("=1/A1", empty)).toBe("#DIV/0!"); // 빈 셀 = 0
  });

  it("한국어 설명 매핑 + isFormulaError", () => {
    for (const code of ["#NAME?", "#REF!", "#VALUE!", "#DIV/0!", "#CIRC!", "#N/A", "#NUM!"] as const) {
      expect(FORMULA_ERROR_KO[code]).toBeTruthy();
      expect(isFormulaError(code)).toBe(true);
    }
    expect(isFormulaError("#PYTHON!")).toBe(false);
    expect(isFormulaError(42)).toBe(false);
  });
});
