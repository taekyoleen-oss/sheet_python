import { describe, expect, it } from "vitest";
import { recalcAfter, type SheetRange } from "@/lib/grid/formula-engine";
import {
  createWorkbook,
  createWorkbookStore,
  setFormulaNotifier,
} from "@/lib/grid/model";
import type { Cell, Workbook } from "@/types/workbook";

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

/** Sheet1 셀 맵으로 워크북 생성 */
function wbWith(cells: Record<string, Cell>): { wb: Workbook; sid: string } {
  const wb = createWorkbook();
  wb.sheets[0].cells = cells;
  return { wb, sid: wb.sheets[0].id };
}

const at = (sid: string, r: number, c: number): SheetRange => ({
  sheetId: sid,
  r0: r,
  c0: c,
  r1: r,
  c1: c,
});

describe("recalcAfter — 연쇄·부분 재계산", () => {
  it("연쇄 재계산: A1 ← B1 ← C1이 의존 순서로 갱신된다", () => {
    const { wb, sid } = wbWith({
      "0:0": { v: 1, t: "n" }, // A1
      "0:1": { v: 0, t: "n", fx: "=A1+1" }, // B1 (캐시 미계산)
      "0:2": { v: 0, t: "n", fx: "=B1*2" }, // C1
    });
    const changed = recalcAfter(wb, [at(sid, 0, 0)]);
    expect(wb.sheets[0].cells["0:1"]).toMatchObject({ v: 2, t: "n", fx: "=A1+1" });
    expect(wb.sheets[0].cells["0:2"]).toMatchObject({ v: 4, t: "n", fx: "=B1*2" });
    expect(changed).toHaveLength(2);

    // 값 편집 → 다시 연쇄
    wb.sheets[0].cells["0:0"] = { v: 5, t: "n" };
    recalcAfter(wb, [at(sid, 0, 0)]);
    expect(wb.sheets[0].cells["0:1"].v).toBe(6);
    expect(wb.sheets[0].cells["0:2"].v).toBe(12);
  });

  it("부분 재계산: 무관한 수식은 계산되지 않는다 (독약 캐시 유지)", () => {
    const { wb, sid } = wbWith({
      "0:0": { v: 1, t: "n" },
      "0:1": { v: 0, t: "n", fx: "=A1+1" },
      "9:9": { v: 999, t: "n", fx: "=Z9+1" }, // 무관 — 캐시가 틀려도 건드리지 않아야 확인 가능
    });
    const changed = recalcAfter(wb, [at(sid, 0, 0)]);
    expect(wb.sheets[0].cells["9:9"].v).toBe(999);
    expect(changed).toEqual([{ sheetId: sid, r0: 0, c0: 1, r1: 0, c1: 1 }]);
  });

  it("변화 없으면 changed 없음 (같은 값 재계산)", () => {
    const { wb, sid } = wbWith({
      "0:0": { v: 1, t: "n" },
      "0:1": { v: 2, t: "n", fx: "=A1+1" }, // 이미 올바른 캐시
    });
    expect(recalcAfter(wb, [at(sid, 0, 0)])).toEqual([]);
  });

  it("순환: 구성원 전부 #CIRC!, 하류 수식은 오류 전파(#CIRC!)", () => {
    const { wb, sid } = wbWith({
      "0:1": { v: 0, t: "n", fx: "=C1+1" }, // B1 → C1
      "0:2": { v: 0, t: "n", fx: "=B1+1" }, // C1 → B1 (순환)
      "0:3": { v: 0, t: "n", fx: "=B1*2" }, // D1 — 순환 하류(구성원 아님)
    });
    recalcAfter(wb, [at(sid, 0, 1)]);
    expect(wb.sheets[0].cells["0:1"]).toMatchObject({ v: "#CIRC!", t: "e" });
    expect(wb.sheets[0].cells["0:2"]).toMatchObject({ v: "#CIRC!", t: "e" });
    expect(wb.sheets[0].cells["0:3"]).toMatchObject({ v: "#CIRC!", t: "e" });
  });

  it("자기 참조도 #CIRC!", () => {
    const { wb, sid } = wbWith({ "0:0": { v: 0, t: "n", fx: "=A1+1" } });
    recalcAfter(wb, [at(sid, 0, 0)]);
    expect(wb.sheets[0].cells["0:0"]).toMatchObject({ v: "#CIRC!", t: "e" });
  });

  it("edited=null은 전체 재계산 — 없던 시트 이름이 해석된다", () => {
    const { wb } = wbWith({ "0:0": { v: 0, t: "n", fx: "=Sheet2!A1+1" } });
    recalcAfter(wb, null);
    expect(wb.sheets[0].cells["0:0"]).toMatchObject({ v: "#REF!", t: "e" });
    wb.sheets.push({ ...wb.sheets[0], id: "s2", name: "Sheet2", cells: { "0:0": { v: 9, t: "n" } } });
    recalcAfter(wb, null);
    expect(wb.sheets[0].cells["0:0"]).toMatchObject({ v: 10, t: "n" });
  });
});

describe("스토어 통합 — 한 트랜잭션·spill 연동·통지", () => {
  it("수식 입력 즉시 계산 + 값 편집 시 같은 트랜잭션 재계산(undo 1회 원복)", async () => {
    const store = createWorkbookStore();
    const sid = store.getState().workbook.sheets[0].id;
    const cells = () => store.getState().workbook.sheets[0].cells;

    store.getState().setCellValue(sid, 0, 0, { v: 1, t: "n" });
    store.getState().setCellValue(sid, 0, 1, { v: null, t: "n", fx: "=A1*10" });
    expect(cells()["0:1"]).toMatchObject({ v: 10, t: "n", fx: "=A1*10" }); // 즉시 계산
    await sleep(350);

    store.getState().setCellValue(sid, 0, 0, { v: 7, t: "n" });
    expect(cells()["0:1"].v).toBe(70); // 편집과 같은 트랜잭션에서 재계산
    await sleep(350);

    store.temporal.getState().undo(); // 한 번의 undo로 편집+재계산 모두 원복
    expect(cells()["0:0"].v).toBe(1);
    expect(cells()["0:1"].v).toBe(10);
  });

  it("spill 반영(applyBlockResult)이 수식 갱신을 유발하고 통지가 나간다", async () => {
    const store = createWorkbookStore();
    const sid = store.getState().workbook.sheets[0].id;
    const notified: SheetRange[][] = [];
    setFormulaNotifier((ranges) => notified.push(ranges));
    try {
      // D1 수식이 A1(블록 spill 위치)을 참조
      store.getState().setCellValue(sid, 0, 3, { v: null, t: "n", fx: "=SUM(A1:A3)" });
      const blockId = store.getState().addPyBlock(sid, { r: 0, c: 0 })!;
      store.getState().applyBlockResult(blockId, [[{ v: 5, t: "n" }], [{ v: 6, t: "n" }]]);
      expect(store.getState().workbook.sheets[0].cells["0:3"].v).toBe(11);

      await Promise.resolve(); // 통지는 microtask
      const flat = notified.flat();
      expect(flat).toContainEqual({ sheetId: sid, r0: 0, c0: 3, r1: 0, c1: 3 });
    } finally {
      setFormulaNotifier(() => {});
    }
  });

  it("clearRange로 입력이 지워지면 수식이 0 기준으로 재계산된다", () => {
    const store = createWorkbookStore();
    const sid = store.getState().workbook.sheets[0].id;
    store.getState().setCellValue(sid, 0, 0, { v: 3, t: "n" });
    store.getState().setCellValue(sid, 1, 0, { v: null, t: "n", fx: "=A1+1" });
    expect(store.getState().workbook.sheets[0].cells["1:0"].v).toBe(4);
    store.getState().clearRange(sid, { r0: 0, c0: 0, r1: 0, c1: 0 });
    expect(store.getState().workbook.sheets[0].cells["1:0"].v).toBe(1); // 빈 셀 = 0
  });

  it("시트 이름 변경으로 이름 참조가 해석/해제된다", () => {
    const store = createWorkbookStore();
    const st = store.getState();
    const sid = st.workbook.sheets[0].id;
    st.addSheet();
    const sid2 = store.getState().workbook.sheets[1].id;
    store.getState().setCellValue(sid2, 0, 0, { v: 42, t: "n" });
    store.getState().setCellValue(sid, 0, 0, { v: null, t: "n", fx: "=결과!A1" });
    expect(store.getState().workbook.sheets[0].cells["0:0"].v).toBe("#REF!");
    store.getState().renameSheet(sid2, "결과");
    expect(store.getState().workbook.sheets[0].cells["0:0"].v).toBe(42);
  });
});
