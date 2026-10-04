// 현재 그리드 선택 → sheet() 참조 문자열 (참조 삽입 바·스니펫 {{range}} 공용, §2.3.2)

import { cellKey } from "@/types/workbook";
import { formatA1 } from "./a1";
import { useWorkbookStore } from "./model";

/**
 * 선택 범위를 sheet() 호출 문자열로. targetSheetId(블록의 시트)와 다르면 시트 접두어.
 * headers 휴리스틱: 첫 행이 전부 문자열(비어 있지 않음) + 아래 행에 숫자 존재 → headers=True.
 */
export function xlRefForSelection(targetSheetId?: string): string | null {
  const st = useWorkbookStore.getState();
  const sel = st.selection;
  if (!sel) return null;
  const sheet = st.workbook.sheets.find((s) => s.id === st.activeSheetId);
  if (!sheet) return null;

  let headers = false;
  if (sel.r1 > sel.r0) {
    let firstRowAllStrings = true;
    let firstRowHasAny = false;
    let bodyHasNumber = false;
    for (let c = sel.c0; c <= sel.c1; c++) {
      const cell = sheet.cells[cellKey(sel.r0, c)];
      if (cell && cell.v !== null && cell.v !== "") {
        firstRowHasAny = true;
        if (cell.t !== "s") firstRowAllStrings = false;
      }
    }
    outer: for (let r = sel.r0 + 1; r <= sel.r1; r++) {
      for (let c = sel.c0; c <= sel.c1; c++) {
        if (sheet.cells[cellKey(r, c)]?.t === "n") {
          bodyHasNumber = true;
          break outer;
        }
      }
    }
    headers = firstRowHasAny && firstRowAllStrings && bodyHasNumber;
  }

  const crossSheet = targetSheetId !== undefined && targetSheetId !== sheet.id;
  const ref = formatA1(sel, crossSheet ? sheet.name : undefined);
  return `sheet("${ref}"${headers ? ", headers=True" : ""})`;
}
