// 부록 O.4 — 전체 열/행 참조·정의된 이름·잘라내기 이동·채우기 핸들·xl() 재작성·블록 앵커 이동·
// XLSX 수식 왕복·함수 도움말·재계산 엔진(긴 사슬·순환)
import { describe, expect, it } from "vitest";
import {
  adjustForStructure,
  FORMULA_FUNCTIONS,
  FUNCTION_HELP,
  namesMap,
  parseFormula,
  rewriteXlRefs,
  shiftFormula,
  type GetCell,
} from "@/lib/grid/formula";
import { buildFillEdits } from "@/lib/grid/clipboard/internal";
import { recalcAfter } from "@/lib/grid/formula-engine";
import { createSheet, createWorkbook, createWorkbookStore } from "@/lib/grid/model";
import { readXlsxBook, sheetsToXlsxBlob } from "@/lib/io/xlsx";
import type { Cell } from "@/types/workbook";

const n = (v: number): Cell => ({ v, t: "n" });
const fx = (f: string): Cell => ({ v: null, t: "n", fx: f });
const g =
  (cells: Record<string, Cell>): GetCell =>
  (sheet, r, c) =>
    sheet === undefined || sheet === "Sheet1" ? cells[`${r}:${c}`] : "#REF!";

describe("전체 열·행 참조 (A:A, 2:3)", () => {
  const cells = { "0:0": n(1), "1:0": n(2), "5:0": n(3), "1:1": n(10), "2:2": n(100) };
  const bounds = () => ({ rows: 200, cols: 26 });

  it("시트 크기로 잘라 평가한다", () => {
    expect(parseFormula("=SUM(A:A)").eval(g(cells), bounds).v).toBe(6);
    expect(parseFormula("=SUM(2:3)").eval(g(cells), bounds).v).toBe(112);
    expect(parseFormula("=SUM(A:B)").eval(g(cells), bounds).v).toBe(16);
  });

  it("복사 이동은 열 문자만, 행 삽입과 무관", () => {
    expect(shiftFormula("=SUM(A:A)", 5, 1)).toBe("=SUM(B:B)");
    expect(shiftFormula("=SUM($A:A)", 0, 2)).toBe("=SUM($A:C)");
    const ed = { sheetName: "Sheet1", axis: "row", index: 0, count: 3 } as const;
    expect(adjustForStructure("=SUM(A:A)+SUM(2:3)", "Sheet1", ed)).toBe("=SUM(A:A)+SUM(5:6)");
  });
});

describe("정의된 이름", () => {
  it("수식에서 이름 → 참조, 함수 호출·시트 접두어와 구분", () => {
    const names = namesMap([{ name: "요율", ref: "Sheet1!$A$1:$A$2" }]);
    const pf = parseFormula("=SUM(요율)*2", names);
    expect(pf.eval(g({ "0:0": n(1), "1:0": n(2) })).v).toBe(6);
    expect(parseFormula("=요율", new Map()).eval(g({})).v).toBe("#NAME?");
  });

  it("스토어: 이름 정의 → 재계산, 행 삽입 시 이름 참조도 이동, 잘못된 이름 거부", () => {
    const store = createWorkbookStore();
    const sid = store.getState().activeSheetId;
    store.getState().setCells(sid, [
      { r: 0, c: 0, cell: n(5) },
      { r: 1, c: 0, cell: n(7) },
      { r: 0, c: 1, cell: fx("=SUM(합계범위)") },
    ]);
    expect(store.getState().workbook.sheets[0].cells["0:1"].v).toBe("#NAME?");
    expect(store.getState().defineName("합계범위", "Sheet1!$A$1:$A$2")).toBeNull();
    expect(store.getState().workbook.sheets[0].cells["0:1"].v).toBe(12);
    store.getState().insertRows(sid, 0, 1);
    expect(store.getState().workbook.names).toEqual([{ name: "합계범위", ref: "Sheet1!$A$2:$A$3" }]);
    expect(store.getState().workbook.sheets[0].cells["1:1"].v).toBe(12);
    expect(store.getState().defineName("A1", "Sheet1!$A$1")).toMatch(/셀 주소/);
    expect(store.getState().defineName("SUM", "Sheet1!$A$1")).toMatch(/함수/);
    expect(store.getState().defineName("x", "$A$1")).toMatch(/시트 이름/);
  });
});

describe("잘라내기 이동 (moveRange)", () => {
  it("옮긴 셀을 가리키던 수식이 따라가고, 옮긴 수식의 바깥 참조는 그대로", () => {
    const store = createWorkbookStore();
    const sid = store.getState().activeSheetId;
    store.getState().setCells(sid, [
      { r: 0, c: 0, cell: n(3) }, // A1
      { r: 0, c: 1, cell: fx("=A1*2") }, // B1
      { r: 0, c: 2, cell: fx("=B1+A1") }, // C1
      { r: 5, c: 5, cell: n(100) }, // F6
      { r: 0, c: 3, cell: fx("=F6") }, // D1 (옮길 대상의 바깥 참조)
    ]);
    // A1:B1 → A10:B10
    expect(store.getState().moveRange(sid, { r0: 0, c0: 0, r1: 0, c1: 1 }, { r: 9, c: 0 })).toBeNull();
    const cells = store.getState().workbook.sheets[0].cells;
    expect(cells["0:0"]).toBeUndefined();
    expect(cells["9:0"].v).toBe(3);
    expect(cells["9:1"]).toMatchObject({ fx: "=A10*2", v: 6 });
    expect(cells["0:2"]).toMatchObject({ fx: "=B10+A10", v: 9 });
    expect(cells["0:3"].fx).toBe("=F6");
  });

  it("spill 셀이 끼면 거부", () => {
    const store = createWorkbookStore();
    const sid = store.getState().activeSheetId;
    store.setState((s) => {
      s.workbook.sheets[0].cells["0:0"] = { v: 1, t: "n", src: "b:o" };
    });
    expect(store.getState().moveRange(sid, { r0: 0, c0: 0, r1: 0, c1: 0 }, { r: 3, c: 3 })).toMatch(/spill/);
  });
});

describe("채우기 핸들 (buildFillEdits)", () => {
  const sheet = createSheet("Sheet1");
  sheet.cells = {
    "0:0": n(1),
    "1:0": n(3),
    "0:1": { v: "항목01", t: "s" },
    "0:2": fx("=A1*10"),
    "0:3": n(7),
  };
  const fill = (src: [number, number, number, number], dest: [number, number, number, number]) =>
    buildFillEdits(
      sheet,
      { r0: src[0], c0: src[1], r1: src[2], c1: src[3] },
      { r0: dest[0], c0: dest[1], r1: dest[2], c1: dest[3] },
    );

  it("숫자 2개 → 선형 추세 (1,3 → 5,7)", () => {
    expect(fill([0, 0, 1, 0], [0, 0, 3, 0]).map((e) => e.cell?.v)).toEqual([5, 7]);
  });
  it("위로 끌면 거꾸로 (1,3 위 → -1)", () => {
    expect(fill([1, 0, 2, 0], [0, 0, 2, 0]).length).toBe(1);
  });
  it("끝 숫자 문자 증가(자릿수 유지)·수식 이동·숫자 1개는 복사", () => {
    expect(fill([0, 1, 0, 1], [0, 1, 2, 1]).map((e) => e.cell?.v)).toEqual(["항목02", "항목03"]);
    expect(fill([0, 2, 0, 2], [0, 2, 2, 2]).map((e) => e.cell?.fx)).toEqual(["=A2*10", "=A3*10"]);
    expect(fill([0, 3, 0, 3], [0, 3, 2, 3]).map((e) => e.cell?.v)).toEqual([7, 7]);
  });
  it("가로 채우기", () => {
    expect(fill([0, 2, 0, 2], [0, 2, 0, 4]).map((e) => [e.c, e.cell?.fx])).toEqual([
      [3, "=B1*10"],
      [4, "=C1*10"],
    ]);
  });
});

describe("xl() 참조 재작성 + 블록 앵커 이동", () => {
  it("rewriteXlRefs는 xl() 문자열만 바꾼다", () => {
    const code = `df = xl("A1:C10", headers=True)\nx = xl('Sheet2!B2')\ns = "A1:C10"`;
    const out = rewriteXlRefs(code, (ref) =>
      adjustForStructure(`=${ref}`, "Sheet1", { sheetName: "Sheet1", axis: "row", index: 0, count: 2 }).slice(1),
    );
    expect(out).toBe(`df = xl("A3:C12", headers=True)\nx = xl('Sheet2!B2')\ns = "A1:C10"`);
  });

  it("행 삽입: 블록 앵커·출력 앵커·코드 xl() 참조가 함께 움직이고 바뀐 블록 id를 돌려준다", () => {
    const store = createWorkbookStore();
    const sid = store.getState().activeSheetId;
    const id = store.getState().addPyBlock(sid, { r: 5, c: 4 })!;
    store.getState().setBlockCode(id, 'df = xl("A1:B4", headers=True)\ndf');
    const changed = store.getState().insertRows(sid, 2, 3);
    expect(changed).toEqual([id]);
    const b = store.getState().workbook.pyBlocks[0];
    expect(b.anchor).toEqual({ r: 8, c: 4 });
    expect(b.outputs![0].anchor).toEqual({ r: 8, c: 4 });
    expect(b.code).toBe('df = xl("A1:B7", headers=True)\ndf');
    expect(store.getState().dirtyBlocks[id]).toBe(true);
  });

  it("시트 이름 변경은 코드의 xl('시트!…')도 갱신", () => {
    const store = createWorkbookStore();
    store.getState().addSheet();
    const [s1, s2] = store.getState().workbook.sheets;
    const id = store.getState().addPyBlock(s1.id, { r: 0, c: 5 })!;
    store.getState().setBlockCode(id, `xl("${s2.name}!A1:A3")`);
    store.getState().renameSheet(s2.id, "요율 표");
    expect(store.getState().workbook.pyBlocks[0].code).toBe(`xl("'요율 표'!A1:A3")`);
  });
});

describe("XLSX 수식·이름 왕복", () => {
  it("지원 수식은 fx로, 이름은 정의로 돌아온다", async () => {
    const sheet = createSheet("Sheet1");
    sheet.cells = { "0:0": n(2), "1:0": n(3), "2:0": { v: 5, t: "n", fx: "=SUM(A1:A2)" } };
    const blob = sheetsToXlsxBlob([sheet], [{ name: "범위", ref: "Sheet1!$A$1:$A$2" }]);
    const book = readXlsxBook(await blob.arrayBuffer());
    expect(book.sheets[0].cells["2:0"]).toEqual({ v: 5, t: "n", fx: "=SUM(A1:A2)" });
    expect(book.names).toEqual([{ name: "범위", ref: "Sheet1!$A$1:$A$2" }]);
    expect(book.stats).toEqual({ kept: 1, dropped: 0 });
  });

  it("미지원 함수 수식은 값만 남기고 센다", async () => {
    const sheet = createSheet("Sheet1");
    sheet.cells = { "0:0": { v: 45000, t: "n", fx: "=TODAY()" } };
    const book = readXlsxBook(await sheetsToXlsxBlob([sheet]).arrayBuffer());
    expect(book.sheets[0].cells["0:0"]).toEqual({ v: 45000, t: "n" });
    expect(book.stats.dropped).toBe(1);
  });
});

describe("재계산 엔진 (색인 + 위상 정렬)", () => {
  it("3000단 누적합 사슬도 빠르게, 정확히", () => {
    const wb = createWorkbook();
    const sheet = wb.sheets[0];
    sheet.cells["0:0"] = n(1);
    sheet.cells["0:1"] = fx("=A1");
    for (let r = 1; r < 3000; r++) {
      sheet.cells[`${r}:0`] = n(1);
      sheet.cells[`${r}:1`] = fx(`=B${r}+A${r + 1}`);
    }
    sheet.rowCount = 3000;
    const t0 = performance.now();
    recalcAfter(wb, null);
    expect(performance.now() - t0).toBeLessThan(3000);
    expect(sheet.cells["2999:1"].v).toBe(3000);
  });

  it("순환 + 하류 + 독립 수식이 섞여도 순환 구성원만 #CIRC!", () => {
    const wb = createWorkbook();
    const c = wb.sheets[0].cells;
    c["0:0"] = fx("=B1+1");
    c["0:1"] = fx("=A1+1"); // A1↔B1 순환
    c["0:2"] = fx("=A1*2"); // 하류
    c["0:3"] = fx("=SUM(E1:E2)"); // 독립
    c["0:4"] = n(4);
    recalcAfter(wb, null);
    expect([c["0:0"].v, c["0:1"].v, c["0:2"].v, c["0:3"].v]).toEqual(["#CIRC!", "#CIRC!", "#CIRC!", 4]);
  });
});

it("함수 도움말이 지원 함수 목록과 1:1", () => {
  expect(Object.keys(FUNCTION_HELP).sort()).toEqual(FORMULA_FUNCTIONS);
});
