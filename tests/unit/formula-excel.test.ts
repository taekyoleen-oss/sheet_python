// 부록 O.1·O.2 — 엑셀식 수식 확장 + 참조 재작성(복사 이동·행/열 삽입 삭제·시트 이름 변경)
import { describe, expect, it } from "vitest";
import {
  adjustForStructure,
  parseFormula,
  renameSheetInFormula,
  shiftFormula,
  type GetCell,
} from "@/lib/grid/formula";
import { buildPasteEdits, snapshotRange } from "@/lib/grid/clipboard/internal";
import { createWorkbookStore } from "@/lib/grid/model";
import type { Cell } from "@/types/workbook";

/** "r:c" → 셀 (Sheet1만) */
const g =
  (cells: Record<string, Cell>): GetCell =>
  (sheet, r, c) =>
    sheet === undefined || sheet === "Sheet1" ? cells[`${r}:${c}`] : "#REF!";
const ev = (src: string, cells: Record<string, Cell> = {}) => parseFormula(src).eval(g(cells));
const n = (v: number): Cell => ({ v, t: "n" });
const s = (v: string): Cell => ({ v, t: "s" });

describe("연산자 — 엑셀 우선순위·형 변환", () => {
  it("거듭제곱·단항·백분율", () => {
    expect(ev("=2^3^2").v).toBe(64); // 엑셀은 좌결합
    expect(ev("=-2^2").v).toBe(4); // 단항이 ^보다 먼저 (엑셀 동일)
    expect(ev("=50%*4").v).toBe(2);
    expect(ev("=0.1+0.2").v).toBe(0.3); // 15자리 반올림
  });

  it("비교·연결 → 불리언·문자열 결과", () => {
    expect(ev("=1<2")).toEqual({ v: true, t: "b" });
    expect(ev('="a"="A"')).toEqual({ v: true, t: "b" }); // 대소문자 무시
    expect(ev('="합계: "&A1', { "0:0": n(3) })).toEqual({ v: "합계: 3", t: "s" });
    expect(ev("=A1<>0", {}).v).toBe(false); // 빈 셀 = 0
  });

  it("숫자 문자열·불리언 강제 변환", () => {
    expect(ev('="3"+1').v).toBe(4);
    expect(ev("=TRUE+1").v).toBe(2);
    expect(ev('="abc"*2').v).toBe("#VALUE!");
  });

  it("오류 전파와 오류 리터럴", () => {
    expect(ev("=A1+1", { "0:0": { v: "#PYTHON!", t: "e" } }).v).toBe("#PYTHON!");
    expect(ev("=#REF!+1").v).toBe("#REF!");
    expect(ev("=SQRT(-1)").v).toBe("#NUM!");
  });
});

describe("함수", () => {
  const data: Record<string, Cell> = {
    "0:0": s("사과"), "0:1": n(10),
    "1:0": s("배"), "1:1": n(20),
    "2:0": s("사과"), "2:1": n(5),
    "3:0": s("귤"), "3:1": n(7),
  };

  it("IF·IFERROR·AND·OR·NOT (고르지 않은 쪽은 평가하지 않음)", () => {
    expect(ev('=IF(B1>5,"큼","작음")', data).v).toBe("큼");
    expect(ev("=IF(FALSE,1/0,2)").v).toBe(2);
    expect(ev('=IFERROR(1/0,"대체")').v).toBe("대체");
    expect(ev("=AND(B1:B4)", data).v).toBe(true);
    expect(ev("=OR(1>2,FALSE)").v).toBe(false);
    expect(ev("=NOT(0)").v).toBe(true);
  });

  it("ROUND 계열·MOD·LOG", () => {
    expect(ev("=ROUND(2.5,0)").v).toBe(3);
    expect(ev("=ROUND(-2.5,0)").v).toBe(-3); // 0에서 먼 쪽
    expect(ev("=ROUND(1.005,2)").v).toBe(1.01);
    expect(ev("=ROUNDUP(1.21,1)").v).toBe(1.3);
    expect(ev("=ROUNDDOWN(-1.29,1)").v).toBe(-1.2);
    expect(ev("=MOD(-3,2)").v).toBe(1); // 엑셀: 제수 부호
    expect(ev("=LOG(8,2)").v).toBe(3);
  });

  it("COUNTIF·SUMIF·SUMIFS·AVERAGEIF (조건 문자열·와일드카드)", () => {
    expect(ev('=COUNTIF(A1:A4,"사과")', data).v).toBe(2);
    expect(ev('=COUNTIF(B1:B4,">=10")', data).v).toBe(2);
    expect(ev('=SUMIF(A1:A4,"사과",B1:B4)', data).v).toBe(15);
    expect(ev('=SUMIFS(B1:B4,A1:A4,"사*",B1:B4,"<10")', data).v).toBe(5);
    expect(ev('=AVERAGEIF(A1:A4,"<>배",B1:B4)', data).v).toBeCloseTo(22 / 3);
  });

  it("VLOOKUP·INDEX·MATCH·SUMPRODUCT", () => {
    expect(ev('=VLOOKUP("귤",A1:B4,2,FALSE)', data).v).toBe(7);
    expect(ev('=VLOOKUP("없음",A1:B4,2,FALSE)', data).v).toBe("#N/A");
    expect(ev('=MATCH("배",A1:A4,0)', data).v).toBe(2);
    expect(ev("=INDEX(A1:B4,2,2)", data).v).toBe(20);
    expect(ev('=INDEX(B1:B4,MATCH("귤",A1:A4,0))', data).v).toBe(7);
    expect(ev("=SUMPRODUCT(B1:B2,B3:B4)", data).v).toBe(10 * 5 + 20 * 7);
  });

  it("근사 VLOOKUP (정렬된 구간표 — 보험 연령대 조회 등)", () => {
    const bands: Record<string, Cell> = {
      "0:0": n(0), "0:1": s("유아"),
      "1:0": n(20), "1:1": s("성인"),
      "2:0": n(65), "2:1": s("고령"),
    };
    expect(ev("=VLOOKUP(34,A1:B3,2)", bands).v).toBe("성인");
    expect(ev("=VLOOKUP(70,A1:B3,2,TRUE)", bands).v).toBe("고령");
  });

  it("통계·문자 함수", () => {
    expect(ev("=MEDIAN(B1:B4)", data).v).toBe(8.5);
    expect(ev("=STDEV.P(1,3)").v).toBe(1);
    expect(ev("=COUNTA(A1:B4)", data).v).toBe(8);
    expect(ev('=LEFT("PyGrid",2)&MID("PyGrid",3,4)&LEN("ab")').v).toBe("PyGrid2");
  });
});

describe("절대·상대 참조 — 복사 이동 (shiftFormula)", () => {
  it("상대만 움직이고 $는 고정", () => {
    expect(shiftFormula("=A1+$B$1+$C1+D$1", 2, 1)).toBe("=B3+$B$1+$C3+E$1");
    expect(shiftFormula("=SUM(A1:A3)*Sheet2!B2", 1, 0)).toBe("=SUM(A2:A4)*Sheet2!B3");
  });

  it("시트 밖으로 밀려나면 #REF!, 문자열 안의 A1은 건드리지 않음", () => {
    expect(shiftFormula("=A1+1", -1, 0)).toBe("=#REF!+1");
    expect(shiftFormula('="A1"&A1', 0, 1)).toBe('="A1"&B1');
  });

  it("함수명 LOG10 등은 참조로 보지 않음", () => {
    expect(shiftFormula("=LOG10(A1)", 1, 0)).toBe("=LOG10(A2)");
  });
});

describe("행/열 삽입·삭제 (adjustForStructure)", () => {
  const rows = (count: number, index = 2) =>
    ({ sheetName: "Sheet1", axis: "row", index, count }) as const;

  it("삽입: 아래쪽 참조는 밀리고, 범위 안쪽 삽입은 범위를 넓힌다 ($ 무관)", () => {
    expect(adjustForStructure("=A1+A3+$A$5", "Sheet1", rows(2))).toBe("=A1+A5+$A$7");
    expect(adjustForStructure("=SUM(A1:A5)", "Sheet1", rows(1))).toBe("=SUM(A1:A6)");
  });

  it("삭제: 참조 셀이 지워지면 #REF!, 범위는 줄어든다", () => {
    expect(adjustForStructure("=A3+A5", "Sheet1", rows(-1))).toBe("=#REF!+A4");
    expect(adjustForStructure("=SUM(A1:A5)", "Sheet1", rows(-2))).toBe("=SUM(A1:A3)");
    expect(adjustForStructure("=SUM(A3:A4)", "Sheet1", rows(-2))).toBe("=SUM(#REF!)");
  });

  it("다른 시트의 수식은 접두어가 그 시트를 가리킬 때만 조정", () => {
    expect(adjustForStructure("=Sheet1!A5+A5", "Sheet2", rows(1))).toBe("=Sheet1!A6+A5");
  });

  it("열 삽입", () => {
    expect(
      adjustForStructure("=B1+$C$1", "Sheet1", { sheetName: "Sheet1", axis: "col", index: 1, count: 1 }),
    ).toBe("=C1+$D$1");
  });

  it("시트 이름 변경은 접두어 참조를 따라간다", () => {
    expect(renameSheetInFormula("=Sheet2!A1+A1", "Sheet2", "보험료 표")).toBe("='보험료 표'!A1+A1");
  });
});

describe("앱 내부 복사·붙여넣기 (buildPasteEdits)", () => {
  const sheet = {
    id: "s",
    name: "Sheet1",
    rowCount: 10,
    colCount: 5,
    cells: {
      "0:0": n(1),
      "0:1": { v: 2, t: "n", fx: "=A1*$A$1" } as Cell,
      "0:2": { v: 9, t: "n", src: "blk:out" } as Cell,
    },
  };

  it("수식은 상대 참조만 이동, spill 잠금(src)은 따라가지 않음", () => {
    const src = snapshotRange(sheet, { r0: 0, c0: 1, r1: 0, c1: 2 });
    const edits = buildPasteEdits(src, { r0: 3, c0: 1, r1: 3, c1: 1 });
    expect(edits).toEqual([
      { r: 3, c: 1, cell: { v: null, t: "n", fx: "=A4*$A$1" } },
      { r: 3, c: 2, cell: { v: 9, t: "n" } },
    ]);
  });

  it("한 셀 복사 → 범위 선택 붙여넣기는 타일링 (수식 채우기)", () => {
    const src = snapshotRange(sheet, { r0: 0, c0: 1, r1: 0, c1: 1 });
    const fx = buildPasteEdits(src, { r0: 1, c0: 1, r1: 3, c1: 1 }).map((e) => e.cell?.fx);
    expect(fx).toEqual(["=A2*$A$1", "=A3*$A$1", "=A4*$A$1"]);
  });
});

describe("스토어 — 구조 변경이 수식과 값을 함께 옮긴다", () => {
  it("행 삽입 후 수식 원문·값이 엑셀처럼 유지된다", () => {
    const store = createWorkbookStore();
    const sid = store.getState().activeSheetId;
    store.getState().setCells(sid, [
      { r: 0, c: 0, cell: n(1) },
      { r: 1, c: 0, cell: n(2) },
      { r: 2, c: 0, cell: { v: null, t: "n", fx: "=SUM(A1:A2)" } },
    ]);
    expect(store.getState().workbook.sheets[0].cells["2:0"].v).toBe(3);

    store.getState().insertRows(sid, 1, 1); // 2행 앞에 삽입 → 범위 안쪽
    const cells = store.getState().workbook.sheets[0].cells;
    expect(cells["3:0"].fx).toBe("=SUM(A1:A3)");
    expect(cells["3:0"].v).toBe(3);

    store.getState().deleteRows(sid, 0, 1); // 1행 삭제
    const after = store.getState().workbook.sheets[0].cells;
    expect(after["2:0"]).toMatchObject({ fx: "=SUM(A1:A2)", v: 2 });
  });

  it("시트 이름 변경 후에도 다른 시트 참조가 끊기지 않는다", () => {
    const store = createWorkbookStore();
    store.getState().addSheet();
    const [s1, s2] = store.getState().workbook.sheets;
    store.getState().setCells(s2.id, [{ r: 0, c: 0, cell: n(7) }]);
    store.getState().setCells(s1.id, [{ r: 0, c: 0, cell: { v: null, t: "n", fx: `=${s2.name}!A1*2` } }]);
    store.getState().renameSheet(s2.id, "요율");
    expect(store.getState().workbook.sheets[0].cells["0:0"]).toMatchObject({ fx: "=요율!A1*2", v: 14 });
  });
});
