// 다중 출력(OutputBinding) 공용 유틸 — 설계서 부록 D.1.
// outputs가 정본이고 anchor/outputMode/includeIndex/output/last는 outputs[0]의 동기화된 뷰다.
// spill 셀의 소유 표시는 "<blockId>:<outputId>" — 구 워크북의 "<blockId>"는 로드 시 정규화한다.

import type { OutputBinding, PyBlock, Workbook } from "@/types/workbook";

export const newId = (): string =>
  typeof crypto !== "undefined" && "randomUUID" in crypto
    ? crypto.randomUUID()
    : Math.random().toString(36).slice(2);

/** spill 셀 소유 태그 */
export const srcTag = (blockId: string, outputId: string): string =>
  `${blockId}:${outputId}`;

/** 소유 태그 → 블록 id (구 워크북의 blockId 단독 표기도 그대로 인식) */
export const srcBlockId = (src: string): string => {
  const i = src.indexOf(":");
  return i < 0 ? src : src.slice(0, i);
};

/** 정규화 전 블록도 안전하게 다루기 위한 폴백: 출력 id가 없으면 블록 id가 곧 태그 */
export const bindingTag = (blockId: string, o: OutputBinding): string =>
  o.id ? srcTag(blockId, o.id) : blockId;

/** 이 출력을 시트 셀에 쓰는가 — 블록이 '시트에 추가'된 상태이고 출력 자체가 꺼지지 않았을 때만 */
export const onSheet = (b: PyBlock, o: OutputBinding): boolean => b.toSheet === true && !o.off;

/** 코드 블록의 출력 목록. 마크다운은 빈 배열, 정규화 전 블록은 레거시 필드로 합성 */
export function outputsOf(b: PyBlock): OutputBinding[] {
  if (b.outputs && b.outputs.length > 0) return b.outputs;
  if (b.kind === "markdown") return [];
  return [
    {
      id: "",
      anchor: b.anchor,
      mode: b.outputMode,
      includeIndex: b.includeIndex,
      ...(b.output ? { selection: b.output } : {}),
      ...(b.last ? { last: b.last } : {}),
    },
  ];
}

/** outputs[0] → 레거시 뷰 필드. 모든 출력 변경 뒤에 호출한다 (immer draft 안전) */
export function syncLegacy(b: PyBlock): void {
  const o = b.outputs?.[0];
  if (!o) return;
  // outputs[0]은 언제나 블록 시트에 놓인다 — 시트가 다르면 블록이 따라간다
  if (o.sheetId && o.sheetId !== b.sheetId) b.sheetId = o.sheetId;
  if (o.sheetId) delete o.sheetId;
  b.anchor = { r: o.anchor.r, c: o.anchor.c };
  b.outputMode = o.mode;
  b.includeIndex = o.includeIndex;
  b.output = o.selection ? { ...o.selection } : undefined;
  b.last = o.last;
}

/** 코드 블록에 출력 바인딩 최소 1개를 보장한다 (레거시 필드에서 유도) */
export function normalizeBlock(b: PyBlock): void {
  // e2e 전용 스위치(__pygridStore와 같은 테스트 훅): spill을 검증하는 스펙은 블록을 시트에 추가된 상태로 시작
  if (
    b.kind !== "markdown" &&
    b.toSheet === undefined &&
    (globalThis as { __pygridToSheetDefault?: boolean }).__pygridToSheetDefault === true
  )
    b.toSheet = true;
  if (b.kind === "markdown") {
    if (b.outputs) delete b.outputs; // 마크다운은 출력이 없다
    return;
  }
  if (!b.outputs || b.outputs.length === 0) {
    b.outputs = [
      {
        id: newId(),
        anchor: { r: b.anchor.r, c: b.anchor.c },
        mode: b.outputMode,
        includeIndex: b.includeIndex,
        ...(b.output ? { selection: b.output } : {}),
        ...(b.last ? { last: b.last } : {}),
      },
    ];
    return;
  }
  for (const o of b.outputs) if (!o.id) o.id = newId();
  syncLegacy(b);
}

/**
 * 워크북 전체 정규화 (열기·복원 진입점). 블록마다 outputs를 보장하고
 * 구 워크북의 `src === blockId` spill 셀을 `"<blockId>:<outputs[0].id>"`로 옮긴다.
 */
export function normalizeWorkbook(wb: Workbook): Workbook {
  const firstTag = new Map<string, string>();
  // 시트에 쓰지 않는 블록·출력 — 저장된 옛 spill 셀은 지운다
  const offIds = new Set<string>();
  const offTags = new Set<string>();
  for (const b of wb.pyBlocks) {
    normalizeBlock(b);
    if (b.kind === "markdown") continue;
    if (b.toSheet !== true) offIds.add(b.id);
    for (const o of b.outputs ?? []) {
      if (onSheet(b, o)) continue;
      offTags.add(srcTag(b.id, o.id));
      if (o.last?.spillRange) delete o.last.spillRange;
    }
    syncLegacy(b);
    const first = b.outputs?.[0];
    if (first) firstTag.set(b.id, srcTag(b.id, first.id));
  }
  // ponytail: 시트 셀 전수 스캔 O(cells) — 열기 1회뿐이라 무해
  for (const sheet of wb.sheets) {
    for (const key of Object.keys(sheet.cells)) {
      const src = sheet.cells[key].src;
      if (src && (offTags.has(src) || (!src.includes(":") && offIds.has(src)))) {
        delete sheet.cells[key];
        continue;
      }
      if (!src || src.includes(":")) continue;
      const tag = firstTag.get(src);
      if (tag) sheet.cells[key] = { ...sheet.cells[key], src: tag };
    }
  }
  return wb;
}
