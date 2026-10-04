import { describe, expect, it } from "vitest";
import {
  checkSizeGuard,
  MAX_BYTES,
  parseWorkbookJson,
  serializeWorkbook,
  WARN_BYTES,
} from "@/lib/io/workbook-json";
import { createWorkbook } from "@/lib/grid/model";
import type { Workbook } from "@/types/workbook";

function sampleWorkbook(): Workbook {
  const wb = createWorkbook();
  wb.title = "왕복 테스트";
  wb.calcMode = "manual";
  wb.settings = { timeoutSec: 30, inferTypesOnPaste: false };
  wb.initScript = "import numpy as np";
  wb.sheets[0].cells = {
    "0:0": { v: "이름", t: "s" },
    "1:0": { v: 12.5, t: "n", f: "0.0%" },
    "2:0": { v: true, t: "b" },
    "3:0": { v: "2026-09-03", t: "d", f: "yyyy-mm-dd" },
    "4:0": { v: "#PYTHON!", t: "e", src: "blk1:out1" },
    "0:6": { v: 7, t: "n", src: "blk1:out2" },
    "5:0": { v: 13.5, t: "n", fx: "=B2*2+SUM(A1:A3)" }, // 미니 수식 (부록 I)
    "6:0": { v: "#DIV/0!", t: "e", fx: "=1/0" },
  };
  wb.sheets[0].colWidths = { 0: 120 };
  wb.sheets[0].frozenCols = 1;
  wb.pyBlocks = [
    {
      id: "blk1",
      sheetId: wb.sheets[0].id,
      anchor: { r: 4, c: 0 },
      code: 'xl("A1")',
      toSheet: true, // 시트에 추가된 블록 — spill·앵커 카드가 왕복 보존돼야 한다
      outputMode: "object",
      includeIndex: "always",
      title: "생명표 요약",
      note: "## 배경\n설명 텍스트 **굵게**", // 부록 J.4 — 코드 블록 설명
      collapsed: true,
      output: { variable: "df", columns: ["a", "b"], rowLimit: 5 },
      last: {
        status: "ok",
        kind: "image",
        stdout: "",
        stderr: "",
        durationMs: 12,
        ranAt: "2026-09-03T00:00:00Z",
        imageBlobId: "blob-1",
      },
      // v1.2 다중 출력: outputs가 정본, 위 레거시 필드는 outputs[0]의 뷰
      outputs: [
        {
          id: "out1",
          anchor: { r: 4, c: 0 },
          mode: "object",
          includeIndex: "always",
          selection: { variable: "df", columns: ["a", "b"], rowLimit: 5 },
          last: {
            status: "ok",
            kind: "image",
            stdout: "",
            stderr: "",
            durationMs: 12,
            ranAt: "2026-09-03T00:00:00Z",
            imageBlobId: "blob-1",
          },
        },
        {
          id: "out2",
          anchor: { r: 0, c: 6 },
          mode: "values",
          includeIndex: "auto",
          selection: { variable: "wide" },
          label: "요약표",
          last: {
            status: "ok",
            kind: "table",
            stdout: "",
            stderr: "",
            durationMs: 3,
            ranAt: "2026-09-03T00:00:01Z",
            spillRange: { r0: 0, c0: 6, r1: 0, c1: 6 },
          },
        },
      ],
    },
    {
      id: "blk2",
      sheetId: wb.sheets[0].id,
      anchor: { r: 0, c: 5 },
      code: "",
      outputMode: "values",
      includeIndex: "auto",
      kind: "markdown",
      title: "분석 개요",
      markdown: "# 분석 개요\n\n- 첫째\n- 둘째",
      collapsed: false,
    },
  ];
  return wb;
}

describe("workbook-json", () => {
  it("왕복 diff 0 — imageBlobId 제거만 예외", () => {
    const wb = sampleWorkbook();
    const restored = parseWorkbookJson(serializeWorkbook(wb));
    const expected = structuredClone(wb);
    delete expected.pyBlocks[0].last!.imageBlobId; // 유일하게 허용되는 차이
    delete expected.pyBlocks[0].outputs![0].last!.imageBlobId;
    expect(restored).toEqual(expected);
  });

  it("미니 수식 fx 왕복 보존 — 값 캐시·오류 셀 포함 (부록 I)", () => {
    const restored = parseWorkbookJson(serializeWorkbook(sampleWorkbook()));
    expect(restored.sheets[0].cells["5:0"]).toEqual({
      v: 13.5,
      t: "n",
      fx: "=B2*2+SUM(A1:A3)",
    });
    expect(restored.sheets[0].cells["6:0"]).toEqual({ v: "#DIV/0!", t: "e", fx: "=1/0" });
  });

  it("v1.2 다중 출력 왕복 보존 — outputs가 정본, 레거시 필드는 outputs[0]와 동기", () => {
    const restored = parseWorkbookJson(serializeWorkbook(sampleWorkbook()));
    const outputs = restored.pyBlocks[0].outputs!;
    expect(outputs).toHaveLength(2);
    expect(outputs[1]).toMatchObject({
      id: "out2",
      anchor: { r: 0, c: 6 },
      mode: "values",
      selection: { variable: "wide" },
      label: "요약표",
    });
    expect(outputs[1].last?.spillRange).toEqual({ r0: 0, c0: 6, r1: 0, c1: 6 });
    // 레거시 뷰 = outputs[0]
    expect(restored.pyBlocks[0].anchor).toEqual(outputs[0].anchor);
    expect(restored.pyBlocks[0].outputMode).toBe(outputs[0].mode);
    expect(restored.pyBlocks[0].output).toEqual(outputs[0].selection);
    // 마크다운 블록은 출력이 없다
    expect(restored.pyBlocks[1].outputs).toBeUndefined();
  });

  it("다중 출력 이전 파일: 로드 시 outputs 1개로 정규화 + src 태그 이관", () => {
    const old = JSON.stringify({
      version: 1,
      id: "x",
      title: "t",
      sheets: [
        {
          id: "s",
          name: "S",
          rowCount: 10,
          colCount: 5,
          cells: {
            "0:0": { v: 1, t: "n", src: "b" }, // 구 표기: blockId 단독
            "1:0": { v: 2, t: "n" },
          },
        },
      ],
      pyBlocks: [
        {
          id: "b",
          sheetId: "s",
          anchor: { r: 0, c: 0 },
          code: "1+1",
          toSheet: true,
          outputMode: "values",
          includeIndex: "always",
          output: { variable: "df" },
          last: { status: "ok", stdout: "", stderr: "", durationMs: 1, ranAt: "" },
        },
      ],
    });
    const wb = parseWorkbookJson(old);
    const outputs = wb.pyBlocks[0].outputs!;
    expect(outputs).toHaveLength(1);
    expect(outputs[0]).toMatchObject({
      anchor: { r: 0, c: 0 },
      mode: "values",
      includeIndex: "always",
      selection: { variable: "df" },
    });
    expect(outputs[0].id).toBeTruthy();
    expect(outputs[0].last?.status).toBe("ok");
    // spill 셀의 소유 표시가 "<blockId>:<outputId>"로 옮겨진다
    expect(wb.sheets[0].cells["0:0"].src).toBe(`b:${outputs[0].id}`);
    expect(wb.sheets[0].cells["1:0"].src).toBeUndefined();
  });

  it("v1.1 필드(kind·title·markdown·collapsed·output) 왕복 보존", () => {
    const restored = parseWorkbookJson(serializeWorkbook(sampleWorkbook()));
    expect(restored.version).toBe(1);
    expect(restored.pyBlocks[0]).toMatchObject({
      title: "생명표 요약",
      note: "## 배경\n설명 텍스트 **굵게**", // 부록 J.4 왕복 보존
      collapsed: true,
      output: { variable: "df", columns: ["a", "b"], rowLimit: 5 },
    });
    expect(restored.pyBlocks[1]).toMatchObject({
      kind: "markdown",
      markdown: "# 분석 개요\n\n- 첫째\n- 둘째",
      title: "분석 개요",
    });
  });

  it("v1.1 이전 파일(선택 필드 없음)도 그대로 열린다", () => {
    const old = JSON.stringify({
      version: 1,
      id: "x",
      title: "t",
      sheets: [{ id: "s", name: "S", rowCount: 10, colCount: 5, cells: {} }],
      pyBlocks: [
        {
          id: "b",
          sheetId: "s",
          anchor: { r: 0, c: 0 },
          code: "1+1",
          outputMode: "values",
          includeIndex: "auto",
        },
      ],
    });
    const wb = parseWorkbookJson(old);
    expect(wb.pyBlocks[0].kind).toBeUndefined(); // 기본 = 코드 블록
    expect(wb.pyBlocks[0].output).toBeUndefined();
    expect(wb.pyBlocks[0].collapsed).toBeUndefined();
  });

  it("손상 파일 거부 (한국어 메시지)", () => {
    expect(() => parseWorkbookJson("{잘못된 json")).toThrow(/손상/);
    expect(() => parseWorkbookJson('{"version":1}')).toThrow(/손상/);
    expect(() =>
      parseWorkbookJson('{"version":1,"id":"x","title":"t","sheets":[{"id":"s"}]}'),
    ).toThrow(/손상/);
  });

  it("상위 버전 파일 안내", () => {
    expect(() => parseWorkbookJson('{"version":2,"id":"x","sheets":[]}')).toThrow(
      /새 버전/,
    );
  });

  it("크기 가드: 50MB 경고·100MB 거부", () => {
    expect(checkSizeGuard(1024)).toBe("ok");
    expect(checkSizeGuard(WARN_BYTES)).toBe("ok");
    expect(checkSizeGuard(WARN_BYTES + 1)).toBe("warn");
    expect(checkSizeGuard(MAX_BYTES)).toBe("warn");
    expect(checkSizeGuard(MAX_BYTES + 1)).toBe("block");
  });

  it("선택 필드 기본값 보정", () => {
    const min = JSON.stringify({
      version: 1,
      id: "x",
      title: "t",
      sheets: [{ id: "s", name: "S", rowCount: 10, colCount: 5, cells: {} }],
    });
    const wb = parseWorkbookJson(min);
    expect(wb.pyBlocks).toEqual([]);
    expect(wb.calcMode).toBe("auto");
    expect(wb.settings.timeoutSec).toBe(60);
  });
});
