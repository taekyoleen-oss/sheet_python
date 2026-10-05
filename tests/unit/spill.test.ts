import { describe, expect, it } from "vitest";
import { checkSpillConflict } from "@/lib/grid/spill";
import type { PyBlock, Sheet } from "@/types/workbook";

const sheet = (cells: Sheet["cells"]): Sheet => ({
  id: "s1",
  name: "Sheet1",
  rowCount: 10,
  colCount: 5,
  cells,
});

const block = (id: string, r: number, c: number): PyBlock => ({
  id,
  sheetId: "s1",
  anchor: { r, c },
  code: "",
  outputMode: "values",
  includeIndex: "auto",
});

describe("checkSpillConflict", () => {
  it("빈 범위 → 충돌 없음, 시트 경계 초과도 성장이므로 충돌 아님", () => {
    const sh = sheet({});
    expect(checkSpillConflict(sh, [block("b1", 0, 0)], "b1", { r: 0, c: 0 }, [3, 2])).toBeNull();
    // rowCount 10·colCount 5를 넘는 spill도 허용
    expect(checkSpillConflict(sh, [block("b1", 8, 3)], "b1", { r: 8, c: 3 }, [10, 10])).toBeNull();
  });

  it("비어 있지 않은 사용자 셀과 겹치면 충돌", () => {
    const sh = sheet({ "2:1": { v: "데이터", t: "s" } });
    const reason = checkSpillConflict(sh, [block("b1", 0, 0)], "b1", { r: 0, c: 0 }, [4, 3]);
    expect(reason).toMatch(/비어 있지 않은 셀\(B3\)/);
  });

  it("자기 spill 교체는 허용, 앵커 셀의 사용자 값도 블록 소유로 본다", () => {
    const sh = sheet({
      "0:0": { v: 1, t: "n", src: "b1" },
      "1:0": { v: 2, t: "n", src: "b1" },
    });
    expect(checkSpillConflict(sh, [block("b1", 0, 0)], "b1", { r: 0, c: 0 }, [5, 2])).toBeNull();
    // 앵커에 src 없는 값이 있어도 (블록 생성 후 입력된 경우) 앵커는 소유
    const sh2 = sheet({ "0:0": { v: "x", t: "s" } });
    expect(checkSpillConflict(sh2, [block("b1", 0, 0)], "b1", { r: 0, c: 0 }, [2, 1])).toBeNull();
  });

  it("다른 블록의 spill 셀과 겹치면 충돌", () => {
    const sh = sheet({ "1:1": { v: 9, t: "n", src: "b2" } });
    const reason = checkSpillConflict(
      sh,
      [block("b1", 0, 0), block("b2", 5, 5)],
      "b1",
      { r: 0, c: 0 },
      [3, 3],
    );
    expect(reason).toMatch(/다른 블록의 결과\(B2\)/);
  });

  it("시트에 추가된 다른 블록의 앵커(셀 없음)와 겹치면 충돌", () => {
    const sh = sheet({});
    const reason = checkSpillConflict(
      sh,
      [block("b1", 0, 0), { ...block("b2", 2, 1), sheetOut: true }],
      "b1",
      { r: 0, c: 0 },
      [4, 3],
    );
    expect(reason).toMatch(/다른 Python 블록\(B3\)/);
  });

  it("시트에 없는 블록·위치 미정 출력의 앵커는 자리를 차지하지 않는다", () => {
    const sh = sheet({});
    const unplaced: PyBlock = {
      ...block("b3", 1, 1),
      sheetOut: true,
      outputs: [{ id: "o", anchor: { r: 1, c: 1 }, mode: "values", includeIndex: "auto", unplaced: true }],
    };
    expect(
      checkSpillConflict(sh, [block("b1", 0, 0), block("b2", 2, 1), unplaced], "b1", { r: 0, c: 0 }, [4, 3]),
    ).toBeNull();
  });

  it("빈 값(v:null)의 소유 없는 셀은 충돌 아님", () => {
    const sh = sheet({ "1:0": { v: null, t: "s" } });
    expect(checkSpillConflict(sh, [block("b1", 0, 0)], "b1", { r: 0, c: 0 }, [3, 1])).toBeNull();
  });
});
