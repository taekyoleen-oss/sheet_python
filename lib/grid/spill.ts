// spill 충돌 검사 — 설계서 §2.3.6·CLAUDE.md §4, 다중 출력은 부록 D.1.
// 충돌: 이 출력 소유가 아닌 비어 있지 않은 셀, 다른 출력의 spill 셀·앵커와 겹침.
// 시트 경계 초과는 충돌이 아니다(applyOutputResults가 시트를 늘린다).

import { cellKey, type PyBlock, type Sheet } from "@/types/workbook";
import { colToLetter } from "./a1";
import { bindingTag, onSheet, outputsOf, srcBlockId } from "./outputs";

export function checkSpillConflict(
  sheet: Sheet,
  blocks: PyBlock[],
  /** 소유 태그 `"<blockId>:<outputId>"` (구 워크북·단일 출력은 blockId 단독) */
  owner: string,
  anchor: { r: number; c: number },
  shape: [rows: number, cols: number],
): string | null {
  const [rows, cols] = shape;
  const r1 = anchor.r + rows - 1;
  const c1 = anchor.c + cols - 1;
  const ownerBlock = srcBlockId(owner);

  // 다른 출력의 앵커가 범위 안에 있으면 충돌 (아직 셀이 없어도)
  for (const other of blocks) {
    for (const o of outputsOf(other)) {
      const tag = bindingTag(other.id, o);
      if (tag === owner) continue;
      if (!onSheet(other, o)) continue; // 시트에 없는(미적용·위치 미정) 출력은 자리를 차지하지 않는다
      const sheetId = o.sheetId ?? other.sheetId;
      if (sheetId !== sheet.id) continue;
      if (o.anchor.r >= anchor.r && o.anchor.r <= r1 && o.anchor.c >= anchor.c && o.anchor.c <= c1) {
        const at = `${colToLetter(o.anchor.c)}${o.anchor.r + 1}`;
        return other.id === ownerBlock
          ? `이 블록의 다른 출력(${at})과 겹칩니다`
          : `다른 Python 블록(${at})과 겹칩니다`;
      }
    }
    // 마크다운 블록은 출력이 없지만 앵커는 자리를 차지한다
    if (other.kind === "markdown" && other.sheetId === sheet.id) {
      if (
        other.anchor.r >= anchor.r &&
        other.anchor.r <= r1 &&
        other.anchor.c >= anchor.c &&
        other.anchor.c <= c1
      ) {
        return `다른 Python 블록(${colToLetter(other.anchor.c)}${other.anchor.r + 1})과 겹칩니다`;
      }
    }
  }

  for (let r = anchor.r; r <= r1; r++) {
    for (let c = anchor.c; c <= c1; c++) {
      const cell = sheet.cells[cellKey(r, c)];
      if (!cell) continue;
      if (cell.src === owner) continue; // 자기 spill 교체는 허용
      if (r === anchor.r && c === anchor.c && !cell.src) continue; // 앵커 셀은 출력 소유
      if (cell.src) {
        return srcBlockId(cell.src) === ownerBlock
          ? `이 블록의 다른 출력(${colToLetter(c)}${r + 1})과 겹칩니다`
          : `다른 블록의 결과(${colToLetter(c)}${r + 1})와 겹칩니다`;
      }
      if (cell.v !== null && cell.v !== "") {
        return `비어 있지 않은 셀(${colToLetter(c)}${r + 1})과 겹칩니다`;
      }
    }
  }
  return null;
}
