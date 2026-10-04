// 부록 P — 속성 창 탐색기 도우미(경로·작업 폴더 코드·불러오기 코드)와 변수를 만든 블록 찾기
import { describe, expect, it } from "vitest";
import {
  breadcrumbs,
  fmtSize,
  joinPath,
  loadCode,
  normPath,
  parentPath,
  relPath,
  varNameFor,
  workDirCode,
} from "@/lib/grid/files";
import { createWorkbookStore } from "@/lib/grid/model";
import { definingBlock } from "@/lib/grid/model-output";

describe("경로", () => {
  it("정규화·이동·빵부스러기", () => {
    expect(normPath("/home/pyodide/../x/./y/")).toBe("/home/x/y");
    expect(joinPath("/home/pyodide", "data")).toBe("/home/pyodide/data");
    expect(parentPath("/home/pyodide")).toBe("/home");
    expect(parentPath("/")).toBe("/");
    expect(breadcrumbs("/home/pyodide").map((b) => b.path)).toEqual(["/", "/home", "/home/pyodide"]);
  });

  it("작업 폴더 안이면 상대 경로, 밖이면 절대 경로", () => {
    expect(relPath("/home/pyodide/data/a.csv", "/home/pyodide")).toBe("data/a.csv");
    expect(relPath("/tmp/a.csv", "/home/pyodide")).toBe("/tmp/a.csv");
    expect(relPath("/tmp/a.csv", "/")).toBe("tmp/a.csv");
  });

  it("크기 표시", () => {
    expect(fmtSize(512)).toBe("512 B");
    expect(fmtSize(2048)).toBe("2.0 KB");
  });
});

describe("코드 생성", () => {
  it("작업 폴더 지정은 os.makedirs + os.chdir 코드", () => {
    expect(workDirCode("/home/pyodide/data")).toBe(
      'import os\nos.makedirs("/home/pyodide/data", exist_ok=True)\nos.chdir("/home/pyodide/data")',
    );
  });

  it("변수 이름: 확장자 제거·특수문자 정리·숫자 시작 보정", () => {
    expect(varNameFor("wage.xlsx")).toBe("wage");
    expect(varNameFor("보험 데이터-2025.csv")).toBe("보험_데이터_2025");
    expect(varNameFor("2025.csv")).toBe("df_2025");
  });

  it("확장자별 불러오기 코드 (마지막 줄은 미리보기 식)", () => {
    const csv = loadCode("/home/pyodide/data/claims.csv", "/home/pyodide")!;
    expect(csv.varName).toBe("claims");
    expect(csv.code).toContain('claims = pd.read_csv("data/claims.csv")');
    expect(csv.code).toContain('encoding="cp949"');
    expect(csv.code.split("\n").pop()).toBe("claims.head(20)");
    expect(loadCode("/home/pyodide/wage.xlsx", "/home/pyodide")!.code).toContain('wage = pd.read_excel("wage.xlsx")');
    expect(loadCode("/a/b.tsv", "/a")!.code).toContain('pd.read_csv("b.tsv", sep="\\t")');
    expect(loadCode("/a/notes.txt", "/a")!.code).toContain('with open("notes.txt", encoding="utf-8") as f:');
    expect(loadCode("/a/image.png", "/a")).toBeNull();
  });
});

describe("변수를 만든 블록", () => {
  it("계산 순서상 그 이름에 대입한 마지막 블록, 없으면 마지막 블록", () => {
    const store = createWorkbookStore();
    const sid = store.getState().activeSheetId;
    const a = store.getState().addPyBlock(sid, { r: 0, c: 5 })!;
    const b = store.getState().addPyBlock(sid, { r: 10, c: 5 })!;
    const c = store.getState().addPyBlock(sid, { r: 20, c: 5 })!;
    store.getState().setBlockCode(a, "df = load()\nmodel = None");
    store.getState().setBlockCode(b, "X, y = split(df)\nmodel = fit(X, y)\nmodel == 1");
    store.getState().setBlockCode(c, "print(model)");
    const wb = store.getState().workbook;
    expect(definingBlock(wb, "model")?.id).toBe(b);
    expect(definingBlock(wb, "df")?.id).toBe(a);
    expect(definingBlock(wb, "X")?.id).toBe(b);
    expect(definingBlock(wb, "없는것")?.id).toBe(c);
  });
});
