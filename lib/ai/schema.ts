// AI 시트 스키마 (부록 E R6 → L.5) — 값은 전송하지 않는다: 이름·사용 범위·헤더 행만.
// list_sheets 도구(lib/ai/tools.ts)와 첫 메시지의 시트 개요 한 줄이 함께 쓴다.

import { cellKey, type Workbook } from "@/types/workbook";
import { formatA1 } from "@/lib/grid/a1";
import { usedRange } from "@/lib/io/data-import";

export interface SheetSchema {
  name: string;
  /** 사용 범위 A1 (예: "A1:F601") — sheet() 참조에 그대로 쓸 수 있다 */
  range: string;
  /** 첫 행 문자열(헤더 후보, 최대 60열). 데이터 값은 전송하지 않는다 */
  headers: string[];
  rows: number;
}

/** 워크북 → 시트 스키마 (순수 함수, 셀 값은 헤더 행 외 미포함) */
export function sheetSchemas(workbook: Workbook): SheetSchema[] {
  return workbook.sheets.map((sheet) => {
    const r = usedRange(sheet);
    const headers: string[] = [];
    for (let c = r.c0; c <= Math.min(r.c1, r.c0 + 59); c++) {
      const cell = sheet.cells[cellKey(r.r0, c)];
      headers.push(cell == null || cell.v == null ? "" : String(cell.v));
    }
    return {
      name: sheet.name,
      range: formatA1(r),
      headers,
      rows: r.r1 - r.r0 + 1,
    };
  });
}

/**
 * 전송 메시지에 붙이는 시트 개요 한 줄 — 사소한 질문에 list_sheets 왕복을 줄인다.
 * 값은 포함되지 않는다(이름·범위·블록 수만). 자세한 내용은 도구로 확인하게 한다.
 */
export function sheetOverview(workbook: Workbook): string {
  const sheets = workbook.sheets
    .map((s) => `${s.name}!${formatA1(usedRange(s))}`)
    .join(", ");
  return `[워크북 개요] 시트: ${sheets || "(없음)"} · 블록 ${workbook.pyBlocks.length}개 (값·코드는 도구로 확인하세요)`;
}
