// 부록 K — 데이터 내장 샘플 워크북 정합성.
// 스키마·sheet() 참조 범위·문서 구조·파일 크기를 파일 단위로 검증한다(실행은 e2e).

import { readFileSync, readdirSync } from "node:fs";
import path from "node:path";

import { describe, expect, test } from "vitest";

import { parseA1 } from "@/lib/grid/a1";
import { blocksInOrder } from "@/lib/grid/model";
import { parseWorkbookJson } from "@/lib/io/workbook-json";
import { cellKey, type Sheet } from "@/types/workbook";

const DIR = path.resolve("data/sample-workbooks");
const FILES = readdirSync(DIR).filter((f) => f.endsWith(".pygrid.json"));

/** 각 워크북 상한 2MB (부록 K.3 — 초과하면 행 샘플링이 필요하다) */
const MAX_BYTES = 2 * 1024 * 1024;

/** 코드에서 sheet("참조", headers=…) 호출 추출 */
const XL_CALL = /\bsheet\(\s*"([^"]+)"\s*(?:,\s*headers\s*=\s*(True|False)\s*)?\)/g;

/** 값이 있는 셀의 최대 행·열 (사용 범위) */
function usedRange(sheet: Sheet): { maxR: number; maxC: number } {
  let maxR = -1;
  let maxC = -1;
  for (const key of Object.keys(sheet.cells)) {
    const [r, c] = key.split(":").map(Number);
    if (r > maxR) maxR = r;
    if (c > maxC) maxC = c;
  }
  return { maxR, maxC };
}

test("샘플 워크북 파일이 존재한다", () => {
  expect(FILES.length).toBeGreaterThanOrEqual(16);
  for (const f of ["life-table", "premium-glm", "freq-severity", "survival-retention", "chain-ladder", "premium-term", "cancer-multi", "risk-rate", "nonsurrender", "term-variants", "whole-life-multi", "accident-class", "cancer-annuity", "ci-whole-life"]) {
    expect(FILES).toContain(`${f}.pygrid.json`);
  }
});

describe.each(FILES)("%s", (file) => {
  const raw = readFileSync(path.join(DIR, file), "utf8");
  const wb = parseWorkbookJson(raw); // 스키마 유효 — 실패하면 여기서 throw

  test("파일 크기 상한", () => {
    expect(Buffer.byteLength(raw, "utf8")).toBeLessThanOrEqual(MAX_BYTES);
  });

  test("시트 크기가 사용 범위를 담는다", () => {
    for (const sheet of wb.sheets) {
      const { maxR, maxC } = usedRange(sheet);
      expect(sheet.rowCount).toBeGreaterThan(maxR);
      expect(sheet.colCount).toBeGreaterThan(maxC);
    }
  });

  test("sheet() 참조가 실제 시트 사용 범위 안이고 헤더 행은 문자열", () => {
    let refCount = 0;
    for (const b of wb.pyBlocks) {
      if (b.kind === "markdown") continue;
      for (const m of b.code.matchAll(XL_CALL)) {
        refCount++;
        const p = parseA1(m[1]);
        const sheet =
          p.sheetName === undefined
            ? wb.sheets.find((s) => s.id === b.sheetId)
            : wb.sheets.find((s) => s.name === p.sheetName);
        expect(sheet, `${b.id}: 시트 없음 ${m[1]}`).toBeDefined();
        const { maxR, maxC } = usedRange(sheet!);
        expect(p.range.r1, `${b.id}: ${m[1]} 행 초과`).toBeLessThanOrEqual(maxR);
        expect(p.range.c1, `${b.id}: ${m[1]} 열 초과`).toBeLessThanOrEqual(maxC);
        if (m[2] !== "True") continue;
        for (let c = p.range.c0; c <= p.range.c1; c++) {
          const cell = sheet!.cells[cellKey(p.range.r0, c)];
          expect(cell?.t, `${b.id}: ${m[1]} 헤더 행이 문자열이 아님 (열 ${c})`).toBe("s");
        }
      }
    }
    expect(refCount).toBeGreaterThan(0);
  });

  test("코드 블록 앵커는 빈 셀 (데이터 위로 spill 금지)", () => {
    for (const b of wb.pyBlocks) {
      if (b.kind === "markdown") continue;
      const sheet = wb.sheets.find((s) => s.id === b.sheetId)!;
      expect(sheet.cells[cellKey(b.anchor.r, b.anchor.c)], `${b.id} 앵커 충돌`).toBeUndefined();
    }
  });

  test("문서 구조 — 제목 마크다운 + [설명 · 코드] 교차", () => {
    const blocks = blocksInOrder(wb);
    if (blocks.length < 2) return; // 단일 블록 예제(손해율)는 문서 구조를 갖지 않는다
    expect(blocks[0].kind).toBe("markdown");
    expect(blocks[0].markdown?.startsWith("# ")).toBe(true);
    blocks.forEach((b, i) => {
      if (i === 0) return;
      const expected = i % 2 === 1 ? "markdown" : "code";
      expect(b.kind ?? "code", `블록 ${i} (${b.id})`).toBe(expected);
      if (expected === "markdown") expect(b.markdown?.startsWith("## ")).toBe(true);
      else expect(b.code.trim().length).toBeGreaterThan(0);
      expect(b.title, `블록 ${i} 제목 없음`).toBeTruthy();
    });
  });
});
