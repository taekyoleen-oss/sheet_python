// CalcHost 구현 + RunCoordinator 배선 — 계약: /output/calc-engine-api.md
// 스토어 반영(트랜잭션·undo 단계)은 전부 여기서. calc-engine.ts는 읽기 전용 계약.

import { toast } from "sonner";
import {
  buildGraph,
  dirtyPropagation,
  resolveRefs,
  RunCoordinator,
  type CalcHost,
  type OutputArea,
  type SheetRange,
  type WorkbookView,
} from "@/lib/runtime/calc-engine";
import { getRuntimeClient } from "@/lib/runtime/client";
import type { RunPayload } from "@/lib/runtime/protocol";
import { summarizeErrorKo } from "@/lib/runtime/errors-ko";
import { putBlob } from "@/lib/storage/db";
import {
  cellKey,
  type CalcMode,
  type Cell,
  type CellRange,
  type OutputBinding,
  type PyBlock,
  type RunResult,
  type Sheet,
} from "@/types/workbook";
import { setFormulaNotifier, useWorkbookStore, type OutputApply } from "./model";
import { onSheet, outputsOf, srcBlockId, srcTag } from "./outputs";
import { checkSpillConflict } from "./spill";

const getState = () => useWorkbookStore.getState();

/** 블록이 점유한 출력 영역 — 값 모드는 마지막 spill, 그 밖(객체·미실행)은 앵커 1×1 (부록 D.1) */
const areasOf = (b: PyBlock): OutputArea[] =>
  outputsOf(b)
    .filter((o) => onSheet(b, o)) // 시트에 쓰지 않는 출력은 셀을 차지하지 않는다
    .map((o) => {
    const rg =
      o.mode === "values" && o.last?.status === "ok" && o.last.spillRange
        ? o.last.spillRange
        : { r0: o.anchor.r, c0: o.anchor.c, r1: o.anchor.r, c1: o.anchor.c };
    return { ...rg, sheetId: o.sheetId ?? b.sheetId };
  });

/** 스토어 현재 상태의 평면 사본 (엔진 계약) */
export function makeView(): WorkbookView {
  const wb = getState().workbook;
  return {
    // 시트에 쓰지 않는 출력은 객체 모드로 실행 — 셀 대신 미리보기(표·이미지·repr)를 받는다
    blocks: wb.pyBlocks.map((b) => {
      const { id, sheetId, anchor, code, outputMode, includeIndex, output, kind, outputs } = b;
      const first = outputsOf(b)[0];
      return {
        id,
        sheetId,
        anchor,
        code,
        outputMode: first && !onSheet(b, first) ? ("object" as const) : outputMode,
        includeIndex,
        output,
        kind,
        outputs: outputs?.map((o) => (onSheet(b, o) ? o : { ...o, mode: "object" as const })),
      };
    }),
    sheetOrder: wb.sheets.map((s) => s.id),
    spills: new Map(
      wb.pyBlocks.map((b) => [
        b.id,
        b.last?.status === "ok" ? b.last.spillRange : undefined,
      ]),
    ),
    areas: new Map(wb.pyBlocks.map((b) => [b.id, areasOf(b)])),
  };
}

const resolveSheetName = (name: string): string | undefined =>
  getState().workbook.sheets.find((s) => s.name === name)?.id;

/** 합성 errorType(엔진 발신)은 message가 이미 한국어 — errors-ko 매핑을 거치지 않는다 */
const ENGINE_ERRORS = new Set(["PyGridCycleError", "PyGridAnalyzeError", "PyGridRefError", "WorkerError"]);

// 객체 모드 blob 저장(await)이 직렬 큐 밖에서 끝나므로, 늦게 도착한 이전 실행의
// 반영이 다음 실행 결과를 덮지 않도록 블록별 실행 시퀀스로 가드한다.
const runSeq = new Map<string, number>();

/** 블록 전체 실패 → 모든 출력 앵커에 #PYTHON! (한 트랜잭션). 마크다운·출력 없음이면 무시 */
function applyBlockFailure(
  blockId: string,
  payload: Extract<RunPayload, { ok: false }>,
): void {
  const st = getState();
  const block = st.workbook.pyBlocks.find((b) => b.id === blockId);
  if (!block) return;
  const last: RunResult = {
    status: "error",
    stdout: payload.stdout,
    stderr: payload.stderr,
    traceback: payload.traceback || undefined,
    summaryKo:
      payload.errorType === "PyGridCycleError"
        ? "순환 참조"
        : ENGINE_ERRORS.has(payload.errorType)
          ? payload.message
          : summarizeErrorKo(payload.errorType, payload.message),
    durationMs: payload.durationMs,
    ranAt: new Date().toISOString(),
  };
  st.applyOutputResults(
    blockId,
    (block.outputs ?? []).map((o) => ({
      outputId: o.id,
      cells: [[{ v: "#PYTHON!", t: "e" as const }]],
      last,
    })),
  );
  st.setExecutedRefs(blockId, null); // 부록 J.3: 실패한 실행의 참조 표시는 지운다
}

export const calcHost: CalcHost = {
  getCell(sheetId, r, c) {
    const cell = getState()
      .workbook.sheets.find((s) => s.id === sheetId)
      ?.cells[cellKey(r, c)];
    return cell ? { v: cell.v, t: cell.t } : undefined;
  },

  resolveSheet: resolveSheetName,

  onBusy(blockId) {
    getState().setBlockRunning(blockId, true);
  },

  onResult(blockId, payload, cells, spill) {
    const seq = (runSeq.get(blockId) ?? 0) + 1;
    runSeq.set(blockId, seq);
    const st = getState();
    st.setBlockRunning(blockId, false);
    st.markExecuted(blockId);
    st.clearDirty(blockId);
    const block = st.workbook.pyBlocks.find((b) => b.id === blockId);
    if (!block) return;
    const ranAt = new Date().toISOString();

    if (!payload.ok) {
      applyBlockFailure(blockId, payload);
      return;
    }

    const base: RunResult = {
      status: "ok",
      kind: payload.kind,
      shape: payload.shape,
      preview: payload.preview,
      stdout: payload.stdout,
      stderr: payload.stderr,
      durationMs: payload.durationMs,
      ranAt,
    };

    // 값 모드 성공 — payload가 셀을 돌려줬는지로 판정 (실행 중 모드 토글에 안전)
    if (cells && spill) {
      const sheet = st.workbook.sheets.find((s) => s.id === block.sheetId);
      const firstId = block.outputs?.[0]?.id;
      if (!sheet || cells.length === 0 || !firstId) return;
      const conflict = checkSpillConflict(
        sheet,
        st.workbook.pyBlocks,
        srcTag(blockId, firstId),
        block.anchor,
        [cells.length, cells[0].length],
      );
      const first = block.outputs?.[0];
      if (conflict && first && onSheet(block, first)) {
        st.applyBlockResult(blockId, [[{ v: "#SPILL!", t: "e" }]], {
          last: { ...base, status: "spill", summaryKo: conflict },
          clearPrevious: true, // 엑셀처럼: 펼칠 수 없으면 이전 결과 영역도 비우고 #SPILL!만 남긴다
        });
        return;
      }
      st.applyBlockResult(blockId, cells, {
        last: { ...base, spillRange: spill },
        clearPrevious: true,
      });
      if (first && onSheet(block, first)) {
        st.setFlash({ sheetId: block.sheetId, range: spill });
        setTimeout(() => getState().setFlash(null), 400);
      }
      return;
    }

    // 객체 모드 성공 — 이미지 blob 저장 후 카드 반영 (저장 실패 시 imageBlobId 생략)
    void (async () => {
      let imageBlobId: string | undefined;
      if (payload.imagePng) {
        try {
          imageBlobId = crypto.randomUUID();
          await putBlob(imageBlobId, new Blob([payload.imagePng], { type: "image/png" }));
        } catch {
          imageBlobId = undefined;
        }
      }
      if (runSeq.get(blockId) !== seq) return; // 그 사이 새 실행 결과가 도착함
      const label = `[${payload.typeName ?? payload.kind}${
        payload.shape ? ` ${payload.shape[0]}×${payload.shape[1]}` : ""
      }]`;
      const stNow = getState();
      if (!stNow.workbook.pyBlocks.some((b) => b.id === blockId)) return;
      stNow.applyBlockResult(blockId, [[{ v: label, t: "s" }]], {
        last: { ...base, imageBlobId },
        clearPrevious: true,
      });
      const range = {
        r0: block.anchor.r,
        c0: block.anchor.c,
        r1: block.anchor.r,
        c1: block.anchor.c,
      };
      stNow.setFlash({ sheetId: block.sheetId, range });
      setTimeout(() => getState().setFlash(null), 400);
    })();
  },

  /**
   * 다중 출력 반영 (부록 D.1) — 코드 1회 실행의 모든 출력을 **한 트랜잭션**으로 적용한다.
   * 이미지 blob 저장은 트랜잭션 전에 끝내 undo 단계가 쪼개지지 않게 한다.
   */
  onOutputs(blockId, payload, items) {
    const seq = (runSeq.get(blockId) ?? 0) + 1;
    runSeq.set(blockId, seq);
    const st0 = getState();
    st0.setBlockRunning(blockId, false);
    st0.markExecuted(blockId);
    st0.clearDirty(blockId);
    if (!st0.workbook.pyBlocks.some((b) => b.id === blockId)) return;
    // 실행 자체가 실패하면 items는 빈 배열 — 모든 출력에 오류를 표시한다 (계약 문서)
    if (!payload.ok) {
      applyBlockFailure(blockId, payload);
      return;
    }

    {
      // 셀 반영은 반드시 **동기**여야 한다 — 엔진은 onOutputs(void) 직후 다음 블록의
      // 스냅샷을 뜨므로, 여기서 미루면 의존 블록이 빈 spill을 읽는다(G4 회귀).
      // 이미지 blob 저장만 반영 후 비동기로 하고 imageBlobId를 사후 패치한다.
      const st = getState();
      const block = st.workbook.pyBlocks.find((b) => b.id === blockId);
      if (!block) return;
      const ranAt = new Date().toISOString();
      const common = {
        stdout: payload.stdout,
        stderr: payload.stderr,
        durationMs: payload.durationMs,
        ranAt,
      };

      // 충돌 검사용 시뮬레이션 시트: 이 블록의 옛 spill은 전부 다시 쓰이므로 비우고,
      // 앞선 출력이 확정한 셀은 채워 넣어 출력끼리의 충돌도 잡는다.
      // ponytail: 실패한 출력의 옛 spill은 남지만(성공 시에만 교체) 시뮬레이션에서는 지운 것으로 본다 —
      // 다음 실행에서 #SPILL!로 드러난다.
      const sim = new Map<string, Sheet>();
      const simSheet = (o: OutputBinding): Sheet | undefined => {
        const id = o.sheetId ?? block.sheetId;
        const hit = sim.get(id);
        if (hit) return hit;
        const orig = st.workbook.sheets.find((s) => s.id === id);
        if (!orig) return undefined;
        const cells: Record<string, Cell> = {};
        for (const key of Object.keys(orig.cells)) {
          const src = orig.cells[key].src;
          if (src && srcBlockId(src) === blockId) continue;
          cells[key] = orig.cells[key];
        }
        const copy = { ...orig, cells };
        sim.set(id, copy);
        return copy;
      };

      const applies: OutputApply[] = [];
      let flash: { sheetId: string; range: CellRange } | null = null;
      for (const { outputId, item, cells, spill } of items) {
        const binding = block.outputs?.find((o) => o.id === outputId);
        if (!binding) continue;
        if (!item.ok) {
          // 출력 단위 실패 격리: 이 출력만 #PYTHON!, 나머지는 정상 반영
          applies.push({
            outputId,
            cells: [[{ v: "#PYTHON!", t: "e" }]],
            last: {
              ...common,
              status: "error",
              traceback: item.traceback || undefined,
              summaryKo: summarizeErrorKo(item.errorType, item.message),
            },
          });
          continue;
        }
        const base: RunResult = {
          ...common,
          status: "ok",
          kind: item.kind,
          shape: item.shape,
          preview: item.preview,
        };
        const sheet = simSheet(binding);
        if (!sheet) continue;
        const write = (rows: Cell[][], last: RunResult) => {
          applies.push({ outputId, cells: rows, clearPrevious: true, last });
          const tag = srcTag(blockId, outputId);
          rows.forEach((row, i) =>
            row.forEach((cell, j) => {
              sheet.cells[cellKey(binding.anchor.r + i, binding.anchor.c + j)] = {
                ...cell,
                src: tag,
              };
            }),
          );
        };
        if (cells && spill && cells.length > 0) {
          const conflict = checkSpillConflict(
            sheet,
            st.workbook.pyBlocks,
            srcTag(blockId, outputId),
            binding.anchor,
            [cells.length, cells[0].length],
          );
          if (conflict && onSheet(block, binding)) {
            applies.push({
              outputId,
              cells: [[{ v: "#SPILL!", t: "e" }]],
              clearPrevious: true, // 이전 결과 영역도 함께 비운다 (남은 셀이 새 자리를 막지 않게)
              last: { ...base, status: "spill", summaryKo: conflict },
            });
            continue;
          }
          write(cells, { ...base, spillRange: spill });
          if (onSheet(block, binding)) flash ??= { sheetId: sheet.id, range: spill };
        } else {
          // 객체 모드(또는 셀이 없는 결과) — 앵커 1칸 카드 라벨
          const label = `[${item.typeName ?? item.kind}${
            item.shape ? ` ${item.shape[0]}×${item.shape[1]}` : ""
          }]`;
          write([[{ v: label, t: "s" }]], base);
          if (onSheet(block, binding)) flash ??= {
            sheetId: sheet.id,
            range: {
              r0: binding.anchor.r,
              c0: binding.anchor.c,
              r1: binding.anchor.r,
              c1: binding.anchor.c,
            },
          };
        }
      }
      st.applyOutputResults(blockId, applies);
      // 부록 J.3: 이 실행이 읽은 sheet() 참조를 그리드 표시용으로 기록 (분석 캐시 재사용 — 값싸다)
      void analyzedRefs(block.code).then((refs) => {
        const now = getState();
        if (!now.workbook.pyBlocks.some((b) => b.id === blockId)) return;
        try {
          now.setExecutedRefs(blockId, resolveRefs(refs, block.sheetId, resolveSheetName));
        } catch {
          now.setExecutedRefs(blockId, null);
        }
      });
      if (flash) {
        st.setFlash(flash);
        setTimeout(() => getState().setFlash(null), 400);
      }
      // 이미지 blob 사후 저장 — 실패하면 미리보기만 없음. 새 실행이 오면 폐기(seq 가드).
      for (const { outputId, item } of items) {
        if (!item.ok || !item.imagePng) continue;
        const png = item.imagePng;
        void (async () => {
          try {
            const id = crypto.randomUUID();
            await putBlob(id, new Blob([png], { type: "image/png" }));
            if (runSeq.get(blockId) === seq) {
              getState().patchOutputImage(blockId, outputId, id);
            }
          } catch {
            /* blob 저장 실패 → 미리보기만 없음 */
          }
        })();
      }
    }
  },
};

// ── 조율기 싱글턴 ────────────────────────────────────────

let coordinator: RunCoordinator | null = null;

export function getCoordinator(): RunCoordinator {
  coordinator ??= new RunCoordinator(getRuntimeClient(), calcHost);
  // 모드·타임아웃은 호출 시점 스토어에서 항상 갱신 (워크북 복원·설정 변경 대응)
  coordinator.mode = getState().workbook.calcMode;
  coordinator.timeoutSec = getState().workbook.settings.timeoutSec;
  return coordinator;
}

/** 계산 모드 전환 — 스토어 + 조율기 동기화 (툴바·상태 바 공용) */
export function setCalcModeEverywhere(mode: CalcMode): void {
  getState().setCalcMode(mode);
  if (coordinator) coordinator.mode = mode;
}

// ── 편집 통지 (자동: 엔진 디바운스 실행 / 수동: dirty 배지) ──

// ponytail: analyze refs 로컬 캐시(코드 문자열 키, 무한 보관) — 블록 수 규모라 무해
const refsCache = new Map<string, string[]>();

async function analyzedRefs(code: string): Promise<string[]> {
  const hit = refsCache.get(code);
  if (hit) return hit;
  try {
    const refs = await getRuntimeClient().analyze(code);
    refsCache.set(code, refs);
    return refs;
  } catch {
    refsCache.set(code, []);
    return [];
  }
}

async function markDirtyManual(ranges: SheetRange[], editedBlockIds: string[]): Promise<void> {
  const view = makeView();
  const resolved = new Map<string, SheetRange[]>();
  for (const b of view.blocks) {
    resolved.set(b.id, resolveRefs(await analyzedRefs(b.code), b.sheetId, resolveSheetName));
  }
  const graph = buildGraph(view.blocks, resolved, view.spills);
  const dirty = dirtyPropagation(resolved, graph, ranges, editedBlockIds);
  if (dirty.size > 0) getState().markDirty([...dirty]);
}

/** 셀 편집·코드 수정 통지 진입점 (SheetGrid·PyBlockCard·붙여넣기에서 호출) */
export function notifyWorkbookEdit(ranges: SheetRange[], editedBlockIds: string[] = []): void {
  if (getState().workbook.pyBlocks.length === 0) return;
  if (getState().workbook.calcMode === "auto") {
    getCoordinator().notifyEdit(ranges, editedBlockIds, makeView());
  } else {
    void markDirtyManual(ranges, editedBlockIds);
  }
}

/** 코드 수정 통지 — 계산 모드와 상관없이 실행하지 않고 dirty 표시만(이전 결과 유지). 실행은 ▶·전체 실행이 한다 */
export function notifyCodeEdit(blockId: string): void {
  if (getState().workbook.pyBlocks.length === 0) return;
  void markDirtyManual([], [blockId]);
}

// 미니 수식 재계산으로 v가 바뀐 셀 → 의존 Python 블록 dirty/재실행 (부록 I.2).
// 스토어가 직접 import하면 순환이라 콜백으로 등록한다.
setFormulaNotifier((ranges) => notifyWorkbookEdit(ranges));

/** §4.8: 준비 전 실행 요청은 큐잉된다 — 문구만 안내 */
export function toastIfQueued(): void {
  if (getRuntimeClient().getStatus() !== "ready") {
    toast("런타임 준비 후 실행됩니다", { id: "rt-queue" });
  }
}

// e2e·디버깅용 노출
if (typeof window !== "undefined") {
  (window as unknown as { __pygridCalc: unknown }).__pygridCalc = {
    notifyEdit: notifyWorkbookEdit,
    whenIdle: () => getCoordinator().whenIdle(),
  };
}
