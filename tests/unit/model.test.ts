import { describe, expect, it } from "vitest";
import { createWorkbook, createWorkbookStore } from "@/lib/grid/model";
import type { PyBlock, Workbook } from "@/types/workbook";

/** 시트로 보내기를 켠 코드 블록 (새 블록 기본은 Python 결과로만 보기) */
const sheetBlock = (
  store: ReturnType<typeof createWorkbookStore>,
  sheetId: string,
  anchor: { r: number; c: number },
): string => {
  const id = store.getState().addPyBlock(sheetId, anchor)!;
  store.getState().setBlockToSheet(id, true);
  return id;
};

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

const fresh = () => {
  const store = createWorkbookStore();
  return { store, sheetId: store.getState().workbook.sheets[0].id };
};

const cells = (store: ReturnType<typeof createWorkbookStore>) =>
  store.getState().workbook.sheets[0].cells;

describe("워크북 스토어", () => {
  it("기본 새 워크북: Sheet1 200×26, timeoutSec 60, calcMode auto", () => {
    const { store } = fresh();
    const wb = store.getState().workbook;
    expect(wb.sheets).toHaveLength(1);
    expect(wb.sheets[0].name).toBe("Sheet1");
    expect(wb.sheets[0].rowCount).toBe(200);
    expect(wb.sheets[0].colCount).toBe(26);
    expect(wb.calcMode).toBe("auto");
    expect(wb.settings).toEqual({ timeoutSec: 60, inferTypesOnPaste: true });
  });

  it("셀 설정·삭제·범위 지우기", () => {
    const { store, sheetId } = fresh();
    expect(store.getState().setCellValue(sheetId, 1, 2, { v: 42, t: "n" })).toBe(true);
    expect(cells(store)["1:2"]).toEqual({ v: 42, t: "n" });
    store.getState().setCellValue(sheetId, 1, 2, null);
    expect(cells(store)["1:2"]).toBeUndefined();

    store.getState().setCells(sheetId, [
      { r: 0, c: 0, cell: { v: "a", t: "s" } },
      { r: 0, c: 1, cell: { v: "b", t: "s" } },
      { r: 5, c: 5, cell: { v: 1, t: "n" } },
    ]);
    store.getState().clearRange(sheetId, { r0: 0, c0: 0, r1: 0, c1: 3 });
    expect(cells(store)["0:0"]).toBeUndefined();
    expect(cells(store)["0:1"]).toBeUndefined();
    expect(cells(store)["5:5"]).toBeDefined();
  });

  it("행 삽입/삭제 시 키 이동", () => {
    const { store, sheetId } = fresh();
    store.getState().setCells(sheetId, [
      { r: 0, c: 0, cell: { v: "머리", t: "s" } },
      { r: 5, c: 3, cell: { v: 7, t: "n" } },
    ]);
    store.getState().insertRows(sheetId, 2, 2);
    expect(store.getState().workbook.sheets[0].rowCount).toBe(202);
    expect(cells(store)["0:0"]).toBeDefined(); // 삽입 지점 이전은 그대로
    expect(cells(store)["5:3"]).toBeUndefined();
    expect(cells(store)["7:3"]).toEqual({ v: 7, t: "n" });

    store.getState().deleteRows(sheetId, 0, 1); // 첫 행 삭제
    expect(cells(store)["0:0"]).toBeUndefined();
    expect(cells(store)["6:3"]).toEqual({ v: 7, t: "n" });
    expect(store.getState().workbook.sheets[0].rowCount).toBe(201);
  });

  it("열 삽입/삭제 시 키·colWidths 이동", () => {
    const { store, sheetId } = fresh();
    store.getState().setCellValue(sheetId, 0, 4, { v: "x", t: "s" });
    store.getState().setColWidth(sheetId, 4, 120);
    store.getState().insertCols(sheetId, 2, 3);
    expect(cells(store)["0:7"]).toEqual({ v: "x", t: "s" });
    expect(store.getState().workbook.sheets[0].colWidths).toEqual({ 7: 120 });
    expect(store.getState().workbook.sheets[0].colCount).toBe(29);

    store.getState().deleteCols(sheetId, 0, 2);
    expect(cells(store)["0:5"]).toEqual({ v: "x", t: "s" });
    expect(store.getState().workbook.sheets[0].colWidths).toEqual({ 5: 120 });
    expect(store.getState().workbook.sheets[0].colCount).toBe(27);

    // frozenCols는 colCount - 1로 클램프 (setFrozenCols와 같은 불변식)
    store.getState().setFrozenCols(sheetId, 26);
    store.getState().deleteCols(sheetId, 0, 25);
    const sheet = store.getState().workbook.sheets[0];
    expect(sheet.colCount).toBe(2);
    expect(sheet.frozenCols).toBe(1);
  });

  it("undo/redo — temporal (300ms 스로틀 간격)", async () => {
    const { store, sheetId } = fresh();
    store.getState().setCellValue(sheetId, 0, 0, { v: "a", t: "s" });
    await sleep(350);
    store.getState().setCellValue(sheetId, 0, 0, { v: "b", t: "s" });
    await sleep(350);

    const temporal = store.temporal.getState();
    temporal.undo();
    expect(cells(store)["0:0"]).toEqual({ v: "a", t: "s" });
    temporal.undo();
    expect(cells(store)["0:0"]).toBeUndefined();
    temporal.redo();
    expect(cells(store)["0:0"]).toEqual({ v: "a", t: "s" });
    temporal.redo();
    expect(cells(store)["0:0"]).toEqual({ v: "b", t: "s" });
  });

  it("selection·activeSheet 변경은 이력에 남지 않는다", async () => {
    const { store, sheetId } = fresh();
    store.getState().setCellValue(sheetId, 0, 0, { v: 1, t: "n" });
    await sleep(350);
    const depth = store.temporal.getState().pastStates.length;
    store.getState().setSelection({ r0: 0, c0: 0, r1: 3, c1: 3 });
    store.getState().setSelection(null);
    await sleep(350);
    expect(store.temporal.getState().pastStates.length).toBe(depth);
  });

  it("src 잠긴 셀 편집 거부", () => {
    const { store, sheetId } = fresh();
    store.getState().setCells(sheetId, [
      { r: 0, c: 0, cell: { v: 1, t: "n", src: "blk1" } },
    ]);
    expect(store.getState().setCellValue(sheetId, 0, 0, { v: 2, t: "n" })).toBe(false);
    expect(cells(store)["0:0"].v).toBe(1);
    // clearRange도 잠긴 셀은 남긴다
    store.getState().clearRange(sheetId, { r0: 0, c0: 0, r1: 0, c1: 0 });
    expect(cells(store)["0:0"].v).toBe(1);
  });

  it("시트 추가/이름 변경/이동/삭제", () => {
    const { store } = fresh();
    store.getState().addSheet();
    const wb = store.getState().workbook;
    expect(wb.sheets).toHaveLength(2);
    expect(wb.sheets[1].name).toBe("Sheet2");
    expect(store.getState().activeSheetId).toBe(wb.sheets[1].id);

    const id2 = wb.sheets[1].id;
    store.getState().renameSheet(id2, "결과");
    expect(store.getState().workbook.sheets[1].name).toBe("결과");

    store.getState().moveSheet(id2, -1);
    expect(store.getState().workbook.sheets[0].id).toBe(id2);

    store.getState().removeSheet(id2);
    expect(store.getState().workbook.sheets).toHaveLength(1);
    // 마지막 시트는 삭제 불가
    store.getState().removeSheet(store.getState().workbook.sheets[0].id);
    expect(store.getState().workbook.sheets).toHaveLength(1);
  });

  it("addSheetWithCells: 시트 생성+채우기가 undo 한 번에 되돌아간다", async () => {
    const { store, sheetId } = fresh();
    store.getState().setCellValue(sheetId, 0, 0, { v: 1, t: "n" });
    await sleep(350);
    store.getState().addSheetWithCells([
      { r: 0, c: 0, cell: { v: "x", t: "s" } },
      { r: 250, c: 30, cell: { v: 2, t: "n" } }, // rowCount/colCount 자동 확장
    ]);
    await sleep(350);
    const wb = store.getState().workbook;
    expect(wb.sheets).toHaveLength(2);
    expect(wb.sheets[1].cells["0:0"]).toEqual({ v: "x", t: "s" });
    expect(wb.sheets[1].rowCount).toBe(251);
    expect(wb.sheets[1].colCount).toBe(31);
    expect(store.getState().activeSheetId).toBe(wb.sheets[1].id);

    store.temporal.getState().undo();
    expect(store.getState().workbook.sheets).toHaveLength(1);
    expect(store.getState().workbook.sheets[0].cells["0:0"]).toEqual({ v: 1, t: "n" });
  });

  it("newWorkbook/loadWorkbook은 이력을 초기화한다", async () => {
    const { store, sheetId } = fresh();
    store.getState().setCellValue(sheetId, 0, 0, { v: 1, t: "n" });
    await sleep(350);
    expect(store.temporal.getState().pastStates.length).toBeGreaterThan(0);
    store.getState().newWorkbook();
    await sleep(350);
    expect(store.temporal.getState().pastStates.length).toBe(0);
    expect(store.getState().workbook.sheets[0].cells).toEqual({});
  });
});

describe("블록 앵커 재지정·접기·출력 선택", () => {
  /** D1 앵커 + 3행 spill(D1:D3)이 기록된 블록 */
  const seed = () => {
    const { store, sheetId } = fresh();
    const id = sheetBlock(store, sheetId, { r: 0, c: 3 });
    store.getState().applyBlockResult(
      id,
      [[{ v: 1, t: "n" }], [{ v: 2, t: "n" }], [{ v: 3, t: "n" }]],
      {
        last: {
          status: "ok",
          stdout: "",
          stderr: "",
          durationMs: 1,
          ranAt: "",
          spillRange: { r0: 0, c0: 3, r1: 2, c1: 3 },
        },
        clearPrevious: true,
      },
    );
    return { store, sheetId, id };
  };

  it("앵커 이동: 이전 spill 제거 + 새 앵커 + dirty, undo 한 단계", async () => {
    const { store, id } = seed();
    await sleep(350);
    expect(Object.keys(cells(store))).toHaveLength(3);
    const depth = store.temporal.getState().pastStates.length;

    expect(store.getState().setBlockAnchor(id, { r: 5, c: 5 })).toBeNull();
    expect(Object.keys(cells(store))).toHaveLength(0); // 옛 spill 제거
    const block = store.getState().workbook.pyBlocks[0];
    expect(block.anchor).toEqual({ r: 5, c: 5 });
    expect(block.last?.spillRange).toBeUndefined();
    expect(store.getState().dirtyBlocks[id]).toBe(true);

    await sleep(350);
    expect(store.temporal.getState().pastStates.length).toBe(depth + 1); // 한 트랜잭션
    store.temporal.getState().undo();
    expect(store.getState().workbook.pyBlocks[0].anchor).toEqual({ r: 0, c: 3 });
    expect(Object.keys(cells(store))).toHaveLength(3);
  });

  it("충돌 거부: 다른 블록 앵커·다른 블록 결과·비어 있지 않은 셀", () => {
    const { store, sheetId, id } = seed();
    const other = sheetBlock(store, sheetId, { r: 9, c: 9 });
    store.getState().applyBlockResult(other, [[{ v: 7, t: "n" }], [{ v: 8, t: "n" }]], {
      clearPrevious: true,
    });
    store.getState().setCellValue(sheetId, 7, 7, { v: "데이터", t: "s" });

    expect(store.getState().setBlockAnchor(id, { r: 9, c: 9 })).toMatch(/블록/);
    expect(store.getState().setBlockAnchor(id, { r: 10, c: 9 })).toMatch(/결과/);
    expect(store.getState().setBlockAnchor(id, { r: 7, c: 7 })).toMatch(/비어 있지 않은/);
    // 거부되면 아무것도 바뀌지 않는다
    expect(store.getState().workbook.pyBlocks[0].anchor).toEqual({ r: 0, c: 3 });
    expect(cells(store)["0:3"]?.v).toBe(1);
    expect(cells(store)["7:7"]?.v).toBe("데이터");
  });

  it("접기 상태는 저장되지만 undo 이력을 만들지 않는다", async () => {
    const { store, sheetId } = fresh();
    const id = sheetBlock(store, sheetId, { r: 0, c: 0 });
    await sleep(350);
    const depth = store.temporal.getState().pastStates.length;

    store.getState().setBlockCollapsed(id, true);
    await sleep(350);
    expect(store.getState().workbook.pyBlocks[0].collapsed).toBe(true);
    expect(store.temporal.getState().pastStates.length).toBe(depth);

    store.getState().setAllCollapsed(false);
    await sleep(350);
    expect(store.getState().workbook.pyBlocks[0].collapsed).toBeUndefined();
    expect(store.temporal.getState().pastStates.length).toBe(depth);

    // 실제 편집은 계속 이력에 남는다
    store.getState().setBlockCode(id, "1+1");
    await sleep(350);
    expect(store.temporal.getState().pastStates.length).toBe(depth + 1);
  });

  it("마크다운 블록: 셀을 쓰지 않고 첫 헤딩이 제목이 된다", () => {
    const { store, sheetId } = fresh();
    const id = store.getState().addPyBlock(sheetId, { r: 2, c: 2 }, "markdown")!;
    store.getState().setBlockMarkdown(id, "# 분석 개요\n본문");
    const block = store.getState().workbook.pyBlocks[0];
    expect(block.kind).toBe("markdown");
    expect(block.title).toBe("분석 개요");
    expect(Object.keys(cells(store))).toHaveLength(0);
  });

  it("setBlockOutput: 병합하고 undefined는 해제", () => {
    const { store, sheetId } = fresh();
    const id = sheetBlock(store, sheetId, { r: 0, c: 0 });
    const out = () => store.getState().workbook.pyBlocks[0].output;

    store.getState().setBlockOutput(id, { variable: "df" });
    expect(out()).toEqual({ variable: "df" });
    store.getState().setBlockOutput(id, { rowLimit: 3, columns: ["a"] });
    expect(out()).toEqual({ variable: "df", rowLimit: 3, columns: ["a"] });
    store.getState().setBlockOutput(id, { variable: undefined, columns: undefined });
    expect(out()).toEqual({ rowLimit: 3 });
    store.getState().setBlockOutput(id, { rowLimit: undefined });
    expect(out()).toBeUndefined();
  });
});

describe("다중 출력 (부록 D.1)", () => {
  const okResult = (spill?: { r0: number; c0: number; r1: number; c1: number }) => ({
    status: "ok" as const,
    stdout: "",
    stderr: "",
    durationMs: 1,
    ranAt: "",
    ...(spill ? { spillRange: spill } : {}),
  });

  /** A1 블록 + 출력 1개(A1) — 정규화로 outputs가 자동 생성된다 */
  const seed = () => {
    const { store, sheetId } = fresh();
    const id = sheetBlock(store, sheetId, { r: 0, c: 0 });
    const block = () => store.getState().workbook.pyBlocks.find((b) => b.id === id)!;
    return { store, sheetId, id, block };
  };

  it("새 블록은 시트로 보내지 않는다 — 결과만 남고 셀은 비어 있다, 끄면 spill 제거", () => {
    const { store, sheetId } = fresh();
    const id = store.getState().addPyBlock(sheetId, { r: 0, c: 0 })!;
    store.getState().applyBlockResult(id, [[{ v: 5, t: "n" }]], {
      last: { status: "ok", stdout: "hi", stderr: "", durationMs: 1, ranAt: "", spillRange: { r0: 0, c0: 0, r1: 0, c1: 0 } },
    });
    expect(cells(store)["0:0"]).toBeUndefined();
    const b = () => store.getState().workbook.pyBlocks[0];
    expect(b().last?.stdout).toBe("hi");
    expect(b().last?.spillRange).toBeUndefined();

    store.getState().setBlockToSheet(id, true);
    store.getState().applyBlockResult(id, [[{ v: 5, t: "n" }]], { clearPrevious: true });
    expect(cells(store)["0:0"]?.v).toBe(5);
    store.getState().setBlockToSheet(id, false);
    expect(cells(store)["0:0"]).toBeUndefined();
  });

  it("addPyBlock은 outputs 1개로 시작하고 레거시 필드와 동기화된다", () => {
    const { block } = seed();
    const outputs = block().outputs!;
    expect(outputs).toHaveLength(1);
    expect(outputs[0]).toMatchObject({
      anchor: { r: 0, c: 0 },
      mode: "values",
      includeIndex: "auto",
    });
    expect(outputs[0].id).toBeTruthy();
    expect(block().anchor).toEqual(outputs[0].anchor);
  });

  it("spill 셀의 src는 '<blockId>:<outputId>'로 기록된다", () => {
    const { store, id, block } = seed();
    const outputId = block().outputs![0].id;
    store.getState().applyBlockResult(id, [[{ v: 1, t: "n" }]], {
      last: okResult({ r0: 0, c0: 0, r1: 0, c1: 0 }),
      clearPrevious: true,
    });
    expect(cells(store)["0:0"].src).toBe(`${id}:${outputId}`);
  });

  it("addOutput: 블록 옆 빈 셀에 추가, 각 출력이 독립된 영역에 기록된다", () => {
    const { store, id, block } = seed();
    const out1 = block().outputs![0].id;
    const out2 = store.getState().addOutput(id)!;
    expect(block().outputs).toHaveLength(2);
    const anchor2 = block().outputs![1].anchor;
    expect(anchor2.r).toBe(0);
    expect(anchor2.c).toBeGreaterThan(0); // 블록 앵커와 겹치지 않는다

    store.getState().applyOutputResults(id, [
      {
        outputId: out1,
        cells: [[{ v: 1, t: "n" }], [{ v: 2, t: "n" }]],
        clearPrevious: true,
        last: okResult({ r0: 0, c0: 0, r1: 1, c1: 0 }),
      },
      {
        outputId: out2,
        cells: [[{ v: "x", t: "s" }]],
        clearPrevious: true,
        last: okResult({ r0: anchor2.r, c0: anchor2.c, r1: anchor2.r, c1: anchor2.c }),
      },
    ]);
    expect(cells(store)["0:0"].src).toBe(`${id}:${out1}`);
    expect(cells(store)["1:0"].src).toBe(`${id}:${out1}`);
    expect(cells(store)[`0:${anchor2.c}`]).toMatchObject({ v: "x", src: `${id}:${out2}` });
    // 레거시 뷰는 outputs[0]만 반영한다
    expect(block().last?.spillRange).toEqual({ r0: 0, c0: 0, r1: 1, c1: 0 });
  });

  it("한 실행의 모든 출력 반영 = undo 한 단계", async () => {
    const { store, id, block } = seed();
    const out1 = block().outputs![0].id;
    const out2 = store.getState().addOutput(id)!;
    await sleep(350);
    const depth = store.temporal.getState().pastStates.length;

    store.getState().applyOutputResults(id, [
      { outputId: out1, cells: [[{ v: 1, t: "n" }]], clearPrevious: true, last: okResult() },
      { outputId: out2, cells: [[{ v: 2, t: "n" }]], clearPrevious: true, last: okResult() },
    ]);
    await sleep(350);
    expect(store.temporal.getState().pastStates.length).toBe(depth + 1);

    store.temporal.getState().undo();
    expect(Object.keys(cells(store))).toHaveLength(0);
  });

  it("removeOutput: 그 출력의 spill만 지운다, 마지막 하나는 거부", () => {
    const { store, id, block } = seed();
    const out1 = block().outputs![0].id;
    const out2 = store.getState().addOutput(id)!;
    const c2 = block().outputs![1].anchor.c;
    store.getState().applyOutputResults(id, [
      { outputId: out1, cells: [[{ v: 1, t: "n" }]], clearPrevious: true, last: okResult() },
      { outputId: out2, cells: [[{ v: 2, t: "n" }]], clearPrevious: true, last: okResult() },
    ]);
    expect(Object.keys(cells(store))).toHaveLength(2);

    store.getState().removeOutput(id, out2);
    expect(block().outputs).toHaveLength(1);
    expect(cells(store)[`0:${c2}`]).toBeUndefined(); // 지운 출력의 셀만 사라진다
    expect(cells(store)["0:0"]?.v).toBe(1);

    store.getState().removeOutput(id, out1); // 마지막 출력은 남는다
    expect(block().outputs).toHaveLength(1);
  });

  it("setOutputAnchor: 그 출력의 옛 spill만 제거하고 이동, 충돌은 거부", () => {
    const { store, id, block } = seed();
    const out1 = block().outputs![0].id;
    const out2 = store.getState().addOutput(id)!;
    const c2 = block().outputs![1].anchor.c;
    store.getState().applyOutputResults(id, [
      {
        outputId: out1,
        cells: [[{ v: 1, t: "n" }]],
        clearPrevious: true,
        last: okResult({ r0: 0, c0: 0, r1: 0, c1: 0 }),
      },
      {
        outputId: out2,
        cells: [[{ v: 2, t: "n" }]],
        clearPrevious: true,
        last: okResult({ r0: 0, c0: c2, r1: 0, c1: c2 }),
      },
    ]);

    expect(store.getState().setOutputAnchor(id, out2, { r: 8, c: 8 })).toBeNull();
    expect(block().outputs![1].anchor).toEqual({ r: 8, c: 8 });
    expect(cells(store)[`0:${c2}`]).toBeUndefined(); // 옮긴 출력의 옛 셀만 제거
    expect(cells(store)["0:0"]?.v).toBe(1); // 다른 출력은 그대로
    expect(block().anchor).toEqual({ r: 0, c: 0 }); // 블록 앵커(=outputs[0])는 그대로

    // 같은 블록의 다른 출력 위로는 옮길 수 없다
    expect(store.getState().setOutputAnchor(id, out2, { r: 0, c: 0 })).toMatch(/다른 출력/);
  });

  it("setBlockAnchor는 outputs[0]을 옮긴다 (레거시 경로)", () => {
    const { store, id, block } = seed();
    expect(store.getState().setBlockAnchor(id, { r: 4, c: 4 })).toBeNull();
    expect(block().anchor).toEqual({ r: 4, c: 4 });
    expect(block().outputs![0].anchor).toEqual({ r: 4, c: 4 });
  });

  it("구 워크북 로드: outputs 생성 + src === blockId 태그 이관", () => {
    const { store } = fresh();
    const wb: Workbook = createWorkbook();
    const sheetId = wb.sheets[0].id;
    const legacy: PyBlock = {
      id: "old-block",
      sheetId,
      anchor: { r: 2, c: 1 },
      code: "1+1",
      outputMode: "values",
      includeIndex: "auto",
      output: { rowLimit: 5 },
    };
    wb.pyBlocks = [legacy];
    wb.sheets[0].cells = {
      "2:1": { v: 1, t: "n", src: "old-block" },
      "3:1": { v: 2, t: "n", src: "old-block" },
      "0:5": { v: "사용자", t: "s" },
    };
    store.getState().loadWorkbook(wb);

    const block = store.getState().workbook.pyBlocks[0];
    const outputId = block.outputs![0].id;
    expect(block.outputs).toHaveLength(1);
    expect(block.outputs![0]).toMatchObject({ anchor: { r: 2, c: 1 }, selection: { rowLimit: 5 } });
    expect(cells(store)["2:1"].src).toBe(`old-block:${outputId}`);
    expect(cells(store)["3:1"].src).toBe(`old-block:${outputId}`);
    expect(cells(store)["0:5"].src).toBeUndefined();

    // 이관된 태그로 spill 잠금·제거가 계속 동작한다
    expect(store.getState().setCellValue(sheetId, 2, 1, { v: 9, t: "n" })).toBe(false);
    store.getState().removePyBlock("old-block");
    expect(cells(store)["2:1"]).toBeUndefined();
    expect(cells(store)["3:1"]).toBeUndefined();
    expect(cells(store)["0:5"]?.v).toBe("사용자");
  });
});

describe("블록 자리 교환 (↑↓)", () => {
  const ok = (spill: { r0: number; c0: number; r1: number; c1: number }) => ({
    status: "ok" as const,
    stdout: "",
    stderr: "",
    durationMs: 1,
    ranAt: "",
    spillRange: spill,
  });

  it("계산 순서 이웃과 앵커 교환 + 양쪽 spill 제거·dirty, undo 한 단계", async () => {
    const { store, sheetId } = fresh();
    const a = sheetBlock(store, sheetId, { r: 0, c: 0 }); // A1 (첫째)
    const b = sheetBlock(store, sheetId, { r: 5, c: 0 }); // A6 (둘째)
    store.getState().applyBlockResult(a, [[{ v: 1, t: "n" }, { v: 2, t: "n" }]], {
      last: ok({ r0: 0, c0: 0, r1: 0, c1: 1 }),
      clearPrevious: true,
    });
    store.getState().applyBlockResult(b, [[{ v: 9, t: "n" }]], {
      last: ok({ r0: 5, c0: 0, r1: 5, c1: 0 }),
      clearPrevious: true,
    });
    await sleep(350);
    expect(Object.keys(cells(store))).toHaveLength(3);
    const depth = store.temporal.getState().pastStates.length;

    // 아래 블록을 위로 → 첫 블록과 자리 교환
    expect(store.getState().swapBlockOrder(b, "up")).toBe(a);
    const byId = (id: string) => store.getState().workbook.pyBlocks.find((x) => x.id === id)!;
    expect(byId(b).anchor).toEqual({ r: 0, c: 0 });
    expect(byId(a).anchor).toEqual({ r: 5, c: 0 });
    expect(Object.keys(cells(store))).toHaveLength(0); // 양쪽 spill 제거
    expect(byId(a).last?.spillRange).toBeUndefined();
    expect(byId(b).last?.spillRange).toBeUndefined();
    expect(store.getState().dirtyBlocks[a]).toBe(true);
    expect(store.getState().dirtyBlocks[b]).toBe(true);

    await sleep(350);
    expect(store.temporal.getState().pastStates.length).toBe(depth + 1); // 한 트랜잭션
    store.temporal.getState().undo();
    expect(byId(a).anchor).toEqual({ r: 0, c: 0 });
    expect(byId(b).anchor).toEqual({ r: 5, c: 0 });
    expect(Object.keys(cells(store))).toHaveLength(3);
  });

  it("경계에서는 교환하지 않는다", () => {
    const { store, sheetId } = fresh();
    const a = sheetBlock(store, sheetId, { r: 0, c: 0 });
    const b = sheetBlock(store, sheetId, { r: 5, c: 0 });
    expect(store.getState().swapBlockOrder(a, "up")).toBeNull(); // 첫 블록 ↑
    expect(store.getState().swapBlockOrder(b, "down")).toBeNull(); // 마지막 블록 ↓
    expect(store.getState().workbook.pyBlocks.map((x) => x.anchor)).toEqual([
      { r: 0, c: 0 },
      { r: 5, c: 0 },
    ]);
    expect(store.getState().dirtyBlocks).toEqual({});
  });

  it("이웃이 다른 시트면 시트까지 함께 옮긴다", () => {
    const { store, sheetId } = fresh();
    store.getState().addSheet();
    const sheet2 = store.getState().workbook.sheets[1].id;
    const a = sheetBlock(store, sheetId, { r: 9, c: 0 }); // Sheet1!A10
    const b = sheetBlock(store, sheet2, { r: 0, c: 0 }); // Sheet2!A1 (시트 순으로 뒤)
    expect(store.getState().swapBlockOrder(a, "down")).toBe(b);
    const byId = (id: string) => store.getState().workbook.pyBlocks.find((x) => x.id === id)!;
    expect(byId(a)).toMatchObject({ sheetId: sheet2, anchor: { r: 0, c: 0 } });
    expect(byId(b)).toMatchObject({ sheetId, anchor: { r: 9, c: 0 } });
  });
});
