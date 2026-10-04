// 블록 실행 공개 API — M5: RunCoordinator(calc-engine) 배선. 반영 로직은 calc-host.ts.

import { toast } from "sonner";
import type { BlockKind } from "@/types/workbook";
import { formatA1 } from "./a1";
import { getCoordinator, makeView, notifyWorkbookEdit, toastIfQueued } from "./calc-host";
import { useWorkbookStore, type AnchorPickTarget } from "./model";

/** ▶ 단일 실행: 해당 블록 + 하위 의존 블록을 위상 순서로 (§2.4) */
export async function runBlock(blockId: string): Promise<void> {
  toastIfQueued();
  await getCoordinator().runBlocks([blockId], makeView());
}

/** ▶ 여러 블록(목차 절 실행 등) — 시드 + 하위 의존을 위상 순서로 */
export async function runBlocks(blockIds: string[]): Promise<void> {
  if (blockIds.length === 0) return;
  toastIfQueued();
  await getCoordinator().runBlocks(blockIds, makeView());
}

/** 전체 실행 — 변수 공유만으로 이어진 블록의 순서는 여기서만 보장된다(§2.4) */
export async function runAllBlocks(): Promise<void> {
  toastIfQueued();
  await getCoordinator().runAll(makeView());
}

// 계산 순서 헬퍼는 스토어(model.ts)가 ↑↓ 교환에 직접 쓰므로 그곳에 있다
export { blocksInOrder } from "./model";

/** ↑↓ 자리 교환 — 성공하면 자동 모드는 두 블록 재실행, 수동 모드는 dirty 배지 */
export function moveBlock(blockId: string, direction: "up" | "down"): void {
  const otherId = useWorkbookStore.getState().swapBlockOrder(blockId, direction);
  if (!otherId) return;
  const blocks = useWorkbookStore.getState().workbook.pyBlocks;
  const runnable = [blockId, otherId].filter(
    (id) => blocks.find((b) => b.id === id)?.kind !== "markdown",
  );
  if (runnable.length > 0) notifyWorkbookEdit([], runnable);
}

function addAtSelection(kind: BlockKind): void {
  const st = useWorkbookStore.getState();
  const anchor = st.selection
    ? { r: st.selection.r0, c: st.selection.c0 }
    : { r: 0, c: 0 };
  // 앵커는 계산 순서용 위치일 뿐(시트에 추가 전엔 셀을 선점하지 않음) — 이미 다른 블록 자리면 아래 빈 행으로
  let id: string | null = null;
  for (let r = anchor.r; id === null && r < anchor.r + 500; r++) {
    id = st.addPyBlock(st.activeSheetId, { r, c: anchor.c }, kind);
  }
  if (id === null) {
    toast.error("이 셀에는 블록을 만들 수 없습니다 (이미 블록 앵커이거나 spill 셀입니다)");
    return;
  }
  st.setFocusBlock(id);
}

/** 툴바 ＋ Python 블록 / Ctrl+Shift+P */
export const addBlockAtSelection = (): void => addAtSelection("code");

/** 툴바 ＋ 마크다운 — 실행되지 않는 문서 블록 (앵커는 위치·목차 대상일 뿐) */
export const addMarkdownAtSelection = (): void => addAtSelection("markdown");

/** 출력 위치 지정 확정 — 실패는 사유 toast, 성공은 자동 모드에서 재실행(수동은 dirty 배지) */
export function applyAnchorPick(
  target: AnchorPickTarget,
  sheetId: string,
  anchor: { r: number; c: number },
): void {
  const { blockId, outputId } = target;
  const err = useWorkbookStore
    .getState()
    .setOutputAnchor(blockId, outputId, { sheetId, r: anchor.r, c: anchor.c });
  if (err) {
    toast.error(err);
    return;
  }
  const st = useWorkbookStore.getState();
  st.setAnchorPicking(null);
  const block = st.workbook.pyBlocks.find((b) => b.id === blockId);
  toast(
    `출력 위치를 ${formatA1({ r0: anchor.r, c0: anchor.c, r1: anchor.r, c1: anchor.c })}(으)로 옮겼습니다`,
  );
  if (block && block.kind !== "markdown") notifyWorkbookEdit([], [blockId]);
}
