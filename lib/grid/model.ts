// 워크북 Zustand 스토어 — immer(불변 편집) + zundo(undo/redo, workbook만 이력에 포함)

import { create } from "zustand";
import { immer } from "zustand/middleware/immer";
import { temporal } from "zundo";
import { setAutoFreeze } from "immer";
import throttle from "lodash/throttle";
import {
  cellKey,
  parseCellKey,
  type BlockKind,
  type CalcMode,
  type Cell,
  type CellRange,
  type IncludeIndex,
  type OutputMode,
  type OutputSelection,
  type PyBlock,
  type RunResult,
  type Sheet,
  type Workbook,
} from "@/types/workbook";
import { colToLetter } from "./a1";
import {
  adjustForStructure,
  FORMULA_FUNCTIONS,
  moveRefsInFormula,
  NAME_RE,
  parseRefText,
  renameSheetInFormula,
  rewriteXlRefs,
  type StructureEdit,
} from "./formula";
import { recalcAfter, rewriteAllFormulas, type SheetRange } from "./formula-engine";
import { markdownTitle } from "./markdown";
import {
  newId,
  normalizeBlock,
  normalizeWorkbook,
  outputsOf,
  srcBlockId,
  srcTag,
  syncLegacy,
} from "./outputs";
import { checkSpillConflict } from "./spill";

// 10k×50 셀 워크북을 deep-freeze하면 로드가 수 초 걸린다. 모든 변경은 스토어 액션 경유.
setAutoFreeze(false);

export const createSheet = (name: string): Sheet => ({
  id: newId(),
  name,
  rowCount: 200,
  colCount: 26,
  cells: {},
});

export const createWorkbook = (): Workbook => {
  const now = new Date().toISOString();
  return {
    id: newId(),
    version: 1,
    title: "새 워크북",
    sheets: [createSheet("Sheet1")],
    pyBlocks: [],
    initScript: "",
    calcMode: "auto",
    settings: { timeoutSec: 60, inferTypesOnPaste: true },
    createdAt: now,
    updatedAt: now,
  };
};

/** 기본 계산 순서: 시트 순 → 앵커 행 → 열 (패널 나열·↑↓ 자리 교환 공용) */
export function blocksInOrder(workbook: Workbook): PyBlock[] {
  const sheetIndex = new Map(workbook.sheets.map((s, i) => [s.id, i]));
  return [...workbook.pyBlocks].sort(
    (a, b) =>
      (sheetIndex.get(a.sheetId) ?? 0) - (sheetIndex.get(b.sheetId) ?? 0) ||
      a.anchor.r - b.anchor.r ||
      a.anchor.c - b.anchor.c,
  );
}

/** 부록 J.2: 범위 전체가 굵은지 (굵게 토글 판정). src 셀은 판정 제외, 서식 없는 셀이 있으면 false */
export function isRangeBold(sheet: Sheet, range: CellRange): boolean {
  const r0 = Math.min(range.r0, range.r1);
  const r1 = Math.max(range.r0, range.r1);
  const c0 = Math.min(range.c0, range.c1);
  const c1 = Math.max(range.c0, range.c1);
  for (let r = r0; r <= r1; r++) {
    for (let c = c0; c <= c1; c++) {
      const cell = sheet.cells[cellKey(r, c)];
      if (cell?.src) continue;
      if (!cell?.st?.b) return false;
    }
  }
  return true;
}

export interface CellEdit {
  r: number;
  c: number;
  cell: Cell | null;
}

/** 한 출력의 실행 반영 단위 (applyOutputResults) */
export interface OutputApply {
  outputId: string;
  /** 출력 앵커부터 기록할 셀. 실패·충돌은 1×1 오류 셀 */
  cells: Cell[][];
  /** 이전 spill 제거 (성공 시에만 교체 — 설계서 §4) */
  clearPrevious?: boolean;
  last?: RunResult;
}

/** 출력 위치 지정 중인 대상 (다음 그리드 클릭이 이 출력의 앵커가 된다) */
export interface AnchorPickTarget {
  blockId: string;
  outputId: string;
}

// ── 미니 수식 재계산 훅 (부록 I.2) ─────────────────────────
// 셀을 쓰는 모든 액션이 같은 트랜잭션 안에서 recalcFormulas를 부른다.
// v가 바뀐 수식 셀은 편집 통지(notifyWorkbookEdit)로 이어진다 — calc-host가
// 등록한다(모듈 import 순환 방지). 통지는 set 트랜잭션 밖(microtask)에서.
let formulaNotify: ((ranges: SheetRange[]) => void) | null = null;
export const setFormulaNotifier = (fn: (ranges: SheetRange[]) => void): void => {
  formulaNotify = fn;
};

/** edited === "all"은 전체 재계산(시트 추가/삭제/이름 변경 등 구조 변경) */
function recalcFormulas(wb: Workbook, edited: SheetRange[] | "all"): void {
  const changed = recalcAfter(wb, edited === "all" ? null : edited);
  if (changed.length > 0 && formulaNotify) {
    const fn = formulaNotify;
    queueMicrotask(() => fn(changed));
  }
}

/** 행/열 삽입·삭제에 따른 한 점(앵커)의 새 위치 — 지워진 행/열 위의 점은 삭제 위치로 당겨진다 */
function shiftPoint(p: { r: number; c: number }, ed: StructureEdit): { r: number; c: number } {
  const key = ed.axis === "row" ? "r" : "c";
  const v = p[key];
  let nv = v;
  if (ed.count > 0) nv = v >= ed.index ? v + ed.count : v;
  else {
    const n = -ed.count;
    nv = v < ed.index ? v : v < ed.index + n ? ed.index : v - n;
  }
  return { ...p, [key]: nv };
}

/**
 * 부록 O.2·O.4: 행/열 삽입·삭제 시 (셀 이동 전에 호출)
 * - 모든 시트 수식 참조·정의된 이름을 엑셀처럼 조정
 * - Python 블록·출력 앵커를 함께 이동, 코드 안 `xl("…")` 참조도 같은 규칙으로 재작성
 * 반환: 바뀐 블록 id (dirty 표시됨 — 호출부가 재실행 통지)
 */
function applyStructure(state: WorkbookState, ed: StructureEdit): string[] {
  const wb = state.workbook;
  rewriteAllFormulas(wb, (fx, own) => adjustForStructure(fx, own, ed));
  for (const n of wb.names ?? []) n.ref = adjustForStructure(`=${n.ref}`, ed.sheetName, ed).slice(1);
  const sheetId = wb.sheets.find((s) => s.name === ed.sheetName)?.id;
  const nameOf = new Map(wb.sheets.map((s) => [s.id, s.name]));
  const changed: string[] = [];
  for (const b of wb.pyBlocks) {
    let hit = false;
    const move = (p: { r: number; c: number }) => {
      const q = shiftPoint(p, ed);
      if (q.r !== p.r || q.c !== p.c) {
        p.r = q.r;
        p.c = q.c;
        hit = true;
      }
    };
    if (b.sheetId === sheetId) {
      move(b.anchor);
      if (b.last?.spillRange) delete b.last.spillRange;
    }
    for (const o of b.outputs ?? []) {
      if ((o.sheetId ?? b.sheetId) !== sheetId) continue;
      move(o.anchor);
      if (o.last?.spillRange) {
        delete o.last.spillRange; // spill 셀은 함께 이동했다 — 테두리는 재실행 때 다시 잡힌다
        hit = true;
      }
    }
    if (b.kind !== "markdown") {
      const own = nameOf.get(b.sheetId) ?? ed.sheetName;
      const code = rewriteXlRefs(b.code, (ref) => adjustForStructure(`=${ref}`, own, ed).slice(1));
      if (code !== b.code) {
        b.code = code;
        hit = true;
      }
    }
    if (hit) {
      changed.push(b.id);
      state.dirtyBlocks[b.id] = true;
    }
  }
  return changed;
}

/** 정의된 이름 검증 — 문제가 있으면 한국어 사유 */
export function nameProblem(name: string, ref: string): string | null {
  if (!NAME_RE.test(name)) return "이름은 글자·밑줄로 시작하고 글자·숫자·밑줄·마침표만 쓸 수 있습니다";
  if (/^[A-Za-z]{1,3}\d+$/.test(name)) return "셀 주소처럼 보이는 이름은 쓸 수 없습니다";
  if (/^(TRUE|FALSE)$/i.test(name) || FORMULA_FUNCTIONS.includes(name.toUpperCase()))
    return "예약어·함수 이름은 쓸 수 없습니다";
  const p = parseRefText(ref);
  if (!p || p.sheetName === undefined) return "참조는 시트 이름을 포함해야 합니다 (예: Sheet1!$A$1:$A$10)";
  return null;
}

/**
 * 지정 출력(outputId 생략 시 블록 전체)의 spill 셀 제거 — 다른 시트 출력까지 훑는다.
 * 반환: 지워진 셀의 시트별 경계 상자 (수식 재계산 대상)
 */
function clearSpillCells(wb: Workbook, blockId: string, outputId?: string): SheetRange[] {
  const tag = outputId === undefined ? undefined : srcTag(blockId, outputId);
  const boxes: SheetRange[] = [];
  for (const sheet of wb.sheets) {
    let box: SheetRange | null = null;
    for (const key of Object.keys(sheet.cells)) {
      const src = sheet.cells[key].src;
      if (!src) continue;
      if (tag ? src === tag : srcBlockId(src) === blockId) {
        delete sheet.cells[key];
        const { r, c } = parseCellKey(key);
        box = box
          ? {
              sheetId: sheet.id,
              r0: Math.min(box.r0, r),
              c0: Math.min(box.c0, c),
              r1: Math.max(box.r1, r),
              c1: Math.max(box.c1, c),
            }
          : { sheetId: sheet.id, r0: r, c0: c, r1: r, c1: c };
      }
    }
    if (box) boxes.push(box);
  }
  return boxes;
}

/** 새 출력·블록 자리 찾기용: 값·spill·블록 앵커·다른 출력 앵커가 있으면 쓸 수 없다 */
export function cellTaken(sheet: Sheet, blocks: PyBlock[], r: number, c: number): boolean {
  const cell = sheet.cells[cellKey(r, c)];
  if (cell && (cell.src || (cell.v !== null && cell.v !== ""))) return true;
  return blocks.some(
    (b) =>
      (b.sheetId === sheet.id && b.anchor.r === r && b.anchor.c === c) ||
      outputsOf(b).some(
        (o) => (o.sheetId ?? b.sheetId) === sheet.id && o.anchor.r === r && o.anchor.c === c,
      ),
  );
}

/** 블록 출력 영역(기존 출력·spill) 오른쪽 두 칸부터 처음 비어 있는 열 */
function freeOutputCol(sheet: Sheet, blocks: PyBlock[], block: PyBlock): number {
  let c = block.anchor.c;
  for (const o of block.outputs ?? []) {
    if ((o.sheetId ?? block.sheetId) !== sheet.id) continue;
    c = Math.max(c, Math.max(o.anchor.c, o.last?.spillRange?.c1 ?? o.anchor.c) + 2);
  }
  const limit = c + 200;
  while (c < limit && cellTaken(sheet, blocks, block.anchor.r, c)) c++;
  return c;
}

export interface WorkbookState {
  workbook: Workbook;
  activeSheetId: string;
  selection: CellRange | null;
  /** 실행 중 블록 표시(#BUSY! 렌더) — workbook 밖이라 undo 이력에 안 남는다 */
  runningBlocks: Record<string, true>;
  /** 실행 성공 400ms 플래시 범위 (렌더 전용) */
  flash: { sheetId: string; range: CellRange } | null;
  /** Python 패널에서 포커스할 블록 (블록 추가 직후) */
  focusBlockId: string | null;
  /** 재실행 필요 블록(수동 모드 배지) — workbook 밖 transient */
  dirtyBlocks: Record<string, true>;
  /** 편집기 xl() 커서 → 그리드 점선 하이라이트 (§4.8) */
  hoverRange: { sheetId: string; range: CellRange } | null;
  /** 그리드 spill hover → 블록 카드 강조 (§4.8 역방향) */
  hoverBlockId: string | null;
  /** 출력 미리보기 탭 대상 블록 */
  selectedBlockId: string | null;
  /** 하단 패널 활성 탭 */
  /** 부록 P.7: 속성 창 위쪽 탭 (하단 패널을 여기로 옮겼다) */
  propsTopTab: "variables" | "console" | "diagnostics";
  /** 부록 P.7: 속성 창 아래쪽 탭 */
  propsBottomTab: "files" | "preview";
  /** 마지막으로 포커스된 블록 편집기 (참조 삽입·스니펫 대상) */
  lastEditorBlockId: string | null;
  /** 출력 위치 지정 중인 출력 — 다음 그리드 클릭이 앵커가 된다 (transient) */
  anchorPicking: AnchorPickTarget | null;
  /** 목차 패널 열림 (설정에 저장, undo 대상 아님) */
  tocOpen: boolean;
  /** AI 채팅 패널 열림 (부록 G.2, 설정에 저장, undo 대상 아님) */
  aiChatOpen: boolean;
  /** 부록 P: 속성 창(변수·파일) 열림 — 설정에 저장, undo 대상 아님 */
  propsOpen: boolean;
  /** 부록 P: 속성 창 고정(true = 옆에 붙어 화면을 나눔, false = 화면 위에 겹침) */
  propsPinned: boolean;
  /** 부록 P.6: 스프레드시트 작게 보기(보조 화면 — 기본 켬, 설정에 저장) */
  gridCompact: boolean;
  /** 부록 P.6: 스프레드시트 전체 화면 (세션 한정) */
  gridMaximized: boolean;
  /** 스프레드시트 패널 접힘 (설정에 저장, undo 대상 아님) */
  gridCollapsed: boolean;
  /** Python 패널 접힘 (설정에 저장, undo 대상 아님) */
  pyCollapsed: boolean;
  /** 상단 뷰 전환 — 워크북 | 데이터 예제/분석 (부록 E, 설정에 저장, undo 대상 아님) */
  view: "workbook" | "reference";
  /** 부록 J.3: 블록별 마지막 성공 실행이 읽은 xl() 참조 범위 (transient — 이력·저장 무관) */
  executedRefs: Record<string, SheetRange[]>;
  /** 부록 J.3: 실행 참조 표시 토글 (기본 켬, 설정에 저장, undo 대상 아님) */
  showRefs: boolean;
  /** 부록 L.2: 채팅 에이전트가 read_range로 읽은 근거 범위 (transient — 이력·저장 무관) */
  chatRefs: SheetRange[];
  /** spill 잠김(src) 셀이면 false를 반환하고 아무것도 바꾸지 않는다 */
  setCellValue: (sheetId: string, r: number, c: number, cell: Cell | null) => boolean;
  /** 일괄 편집 = 한 트랜잭션 = 한 undo 단계 */
  setCells: (sheetId: string, edits: CellEdit[]) => void;
  clearRange: (sheetId: string, range: CellRange) => void;
  /** 행/열 삽입·삭제 — 수식·이름·블록 앵커·xl() 참조 조정까지 한 트랜잭션. 반환: 바뀐 블록 id */
  insertRows: (sheetId: string, index: number, count: number) => string[];
  insertCols: (sheetId: string, index: number, count: number) => string[];
  deleteRows: (sheetId: string, index: number, count: number) => string[];
  deleteCols: (sheetId: string, index: number, count: number) => string[];
  /**
   * 부록 O.4 잘라내기 이동: src 범위를 dest로 옮긴다. 옮긴 셀을 가리키던 모든 수식·이름이 따라간다.
   * spill 셀이 끼면 아무것도 바꾸지 않고 한국어 사유를 반환한다.
   */
  moveRange: (sheetId: string, src: CellRange, dest: { r: number; c: number }) => string | null;
  /** 부록 O.4 이름 정의 (같은 이름이면 교체) — 문제가 있으면 한국어 사유 */
  defineName: (name: string, ref: string) => string | null;
  removeName: (name: string) => void;
  addSheet: () => void;
  /** 새 시트 생성 + 셀 채우기를 한 트랜잭션(= 한 undo 단계)으로 */
  addSheetWithCells: (edits: CellEdit[]) => void;
  renameSheet: (sheetId: string, name: string) => void;
  removeSheet: (sheetId: string) => void;
  moveSheet: (sheetId: string, offset: number) => void;
  setColWidth: (sheetId: string, col: number, width: number) => void;
  setFrozenCols: (sheetId: string, n: number) => void;
  setTitle: (title: string) => void;
  setSelection: (range: CellRange | null) => void;
  setActiveSheet: (id: string) => void;
  newWorkbook: () => void;
  loadWorkbook: (wb: Workbook) => void;
  /** 블록 생성. 앵커에 이미 블록·spill 셀이 있으면 null */
  addPyBlock: (
    sheetId: string,
    anchor: { r: number; c: number },
    kind?: BlockKind,
  ) => string | null;
  /**
   * 앵커 재지정 — 한 트랜잭션(= 한 undo 단계).
   * 이전 spill(src===id) 제거 + 앵커(·시트) 이동 + dirty 표시.
   * 충돌하면 아무것도 바꾸지 않고 한국어 사유를 반환한다.
   */
  setBlockAnchor: (
    id: string,
    anchor: { r: number; c: number },
    sheetId?: string,
  ) => string | null;
  /**
   * ↑↓ 자리 교환 — 계산 순서상 이웃 블록과 {sheetId, anchor}를 맞바꾼다(= 실행 순서 변경).
   * 한 트랜잭션: 두 블록의 spill 제거 + 자리 교환 + 양쪽 dirty.
   * 반환값은 교환한 상대 블록 id (경계에서 교환하지 않으면 null).
   */
  swapBlockOrder: (id: string, direction: "up" | "down") => string | null;
  /** 블록 + 그 spill 셀 제거 (한 트랜잭션) */
  removePyBlock: (id: string) => void;
  setBlockCode: (id: string, code: string) => void;
  /** outputs[0]의 출력 모드 (레거시 뷰) */
  setBlockOutputMode: (id: string, mode: OutputMode) => void;
  /** outputs[0]의 출력 선택 병합. 값이 undefined인 키는 제거된다 (레거시 뷰) */
  setBlockOutput: (id: string, patch: OutputSelection) => void;
  /** 출력 추가 — 블록 근처 빈 셀에 기본 바인딩(마지막 표현식·값 모드). 반환: 새 출력 id */
  addOutput: (blockId: string) => string | null;
  /**
   * 부록 O.5: 출력 여러 개를 한 트랜잭션(= 한 undo 단계)으로 추가 — 모델 결과 보내기.
   * start가 없으면 블록 출력 영역 오른쪽 빈 칸부터, 각 출력은 width+1칸 간격으로 가로 배치.
   */
  addOutputs: (
    blockId: string,
    specs: { selection: OutputSelection; label: string; width: number; includeIndex?: IncludeIndex }[],
    start?: { sheetId: string; r: number; c: number },
  ) => string[];
  /** 출력 삭제 — 그 출력의 spill 셀을 같은 트랜잭션에서 지운다. 마지막 하나는 거부 */
  removeOutput: (blockId: string, outputId: string) => void;
  /**
   * 출력 앵커 재지정 — 한 트랜잭션(= 한 undo 단계).
   * 이 출력의 이전 spill 제거 + 앵커(·시트) 이동 + dirty 표시. 충돌하면 한국어 사유 반환.
   */
  setOutputAnchor: (
    blockId: string,
    outputId: string,
    target: { sheetId?: string; r: number; c: number },
  ) => string | null;
  setOutputSelection: (blockId: string, outputId: string, patch: OutputSelection) => void;
  setOutputMode: (blockId: string, outputId: string, mode: OutputMode) => void;
  setOutputIncludeIndex: (blockId: string, outputId: string, value: IncludeIndex) => void;
  setOutputLabel: (blockId: string, outputId: string, label: string) => void;
  /** 한 실행의 모든 출력 반영 — 한 트랜잭션(= 한 undo 단계) */
  applyOutputResults: (blockId: string, results: OutputApply[]) => void;
  /** 이미지 blob 사후 패치 — 비동기 저장 완료 후 last.imageBlobId만 갱신 (히스토리 무관) */
  patchOutputImage: (blockId: string, outputId: string, imageBlobId: string) => void;
  setBlockMarkdown: (id: string, markdown: string) => void;
  /** 부록 J.4: 코드 블록 설명. null이면 필드 삭제(영역 소멸). 마크다운 블록엔 불가 */
  setBlockNote: (id: string, note: string | null) => void;
  setBlockTitle: (id: string, title: string) => void;
  setBlockCollapsed: (id: string, collapsed: boolean) => void;
  /** 패널 헤더 '모두 접기/펼치기' */
  setAllCollapsed: (collapsed: boolean) => void;
  /**
   * outputs[0] 결과 반영 (레거시 단일 출력 경로) — applyOutputResults의 얇은 래퍼.
   * 실패(#PYTHON!)·충돌(#SPILL!)은 앵커 1셀만 쓰고 이전 spill은 유지한다(설계서 §4: 성공 시에만 교체).
   */
  applyBlockResult: (
    blockId: string,
    cells: Cell[][],
    opts?: { last?: RunResult; clearPrevious?: boolean },
  ) => void;
  setBlockRunning: (id: string, running: boolean) => void;
  setFlash: (flash: { sheetId: string; range: CellRange } | null) => void;
  setFocusBlock: (id: string | null) => void;
  markDirty: (ids: string[]) => void;
  clearDirty: (id: string) => void;
  setCalcMode: (mode: CalcMode) => void;
  setInitScript: (script: string) => void;
  setHoverRange: (hover: { sheetId: string; range: CellRange } | null) => void;
  setHoverBlock: (id: string | null) => void;
  setSelectedBlock: (id: string | null) => void;
  /** 속성 창의 해당 탭을 열고(창이 닫혀 있으면 연다) 보여 준다 */
  showPanelTab: (
    tab: WorkbookState["propsTopTab"] | WorkbookState["propsBottomTab"],
  ) => void;
  setLastEditorBlock: (id: string | null) => void;
  setAnchorPicking: (target: AnchorPickTarget | null) => void;
  /** 부록 J.2: 선택 범위에 셀 서식 병합 적용 — src 셀 제외, 한 트랜잭션(= 한 undo).
   *  b: false·fs: null은 해당 서식 제거. 빈 셀에도 적용 가능({v:null} 셀 생성) */
  applyCellStyle: (
    sheetId: string,
    range: CellRange,
    patch: { b?: boolean; fs?: number | null },
  ) => void;
  /** 부록 J.3: 성공 실행의 참조 기록 (null이면 제거) */
  setExecutedRefs: (blockId: string, refs: SheetRange[] | null) => void;
  setShowRefs: (show: boolean) => void;
  /** 부록 L.2: 채팅 근거 하이라이트 (빈 배열이면 표시 없음) */
  setChatRefs: (refs: SheetRange[]) => void;
  setTocOpen: (open: boolean) => void;
  setAiChatOpen: (open: boolean) => void;
  setPropsOpen: (open: boolean) => void;
  setPropsPinned: (pinned: boolean) => void;
  setGridCompact: (compact: boolean) => void;
  setGridMaximized: (max: boolean) => void;
  /** 부록 P: 작업 폴더 지정(워크북에 저장, null = 지정 해제). 실제 적용은 런타임 코드(os.chdir) */
  setWorkDir: (dir: string | null) => void;
  /** 그리드·Python 패널 접기 — 둘 다 접히면 화면이 비므로 반대쪽은 자동으로 펼친다 */
  setPanelCollapsed: (panel: "grid" | "python", collapsed: boolean) => void;
  setView: (view: "workbook" | "reference") => void;
}

const norm = (rg: CellRange): CellRange => ({
  r0: Math.min(rg.r0, rg.r1),
  c0: Math.min(rg.c0, rg.c1),
  r1: Math.max(rg.r0, rg.r1),
  c1: Math.max(rg.c0, rg.c1),
});

/** cells 레코드 키 재배치. map이 null을 반환하면 그 셀은 삭제된다 */
function remapCells(
  cells: Record<string, Cell>,
  map: (r: number, c: number) => [number, number] | null,
): Record<string, Cell> {
  const next: Record<string, Cell> = {};
  for (const key of Object.keys(cells)) {
    const { r, c } = parseCellKey(key);
    const to = map(r, c);
    if (to) next[cellKey(to[0], to[1])] = cells[key];
  }
  return next;
}

function remapWidths(
  widths: Record<number, number> | undefined,
  map: (c: number) => number | null,
): Record<number, number> | undefined {
  if (!widths) return widths;
  const next: Record<number, number> = {};
  for (const k of Object.keys(widths)) {
    const c = Number(k);
    const to = map(c);
    if (to !== null) next[to] = widths[c];
  }
  return next;
}

export const createWorkbookStore = () => {
  // partialize 메모화: workbook 참조가 같으면 같은 스냅샷 객체를 반환해
  // equality(===)로 selection/activeSheet 변경을 이력에서 제외한다.
  // 이력에서 빼는 블록 필드: last(실행 결과)·collapsed(카드 접기) — 문서 내용이 아니다.
  // 출력별 last도 같은 이유로 제외한다(실행 결과가 undo 단계를 만들지 않게).
  const stripBlocks = (blocks: PyBlock[]) =>
    blocks.map(({ last: _last, collapsed: _collapsed, outputs, ...b }) =>
      outputs ? { ...b, outputs: outputs.map(({ last: _l, ...o }) => o) } : b,
    );

  /** 이력에 남길 변경이 없으면 true — 접기 토글만으로 undo 단계가 생기지 않게 한다 */
  const sameHistory = (a: Workbook, b: Workbook): boolean => {
    for (const k of Object.keys(a) as (keyof Workbook)[]) {
      if (k !== "pyBlocks" && a[k] !== b[k]) return false;
    }
    return JSON.stringify(a.pyBlocks) === JSON.stringify(b.pyBlocks);
  };

  let cacheWb: Workbook | undefined;
  let cacheSnap: { workbook: Workbook } | undefined;
  const partialize = (s: WorkbookState): { workbook: Workbook } => {
    if (cacheWb === s.workbook && cacheSnap) return cacheSnap;
    cacheWb = s.workbook;
    const next = {
      workbook: { ...s.workbook, pyBlocks: stripBlocks(s.workbook.pyBlocks) },
    };
    if (!cacheSnap || !sameHistory(cacheSnap.workbook, next.workbook)) cacheSnap = next;
    return cacheSnap;
  };

  let cancelPending: () => void = () => {};
  let resetHistory: () => void = () => {};

  const store = create<WorkbookState>()(
    temporal(
      immer((set, get) => {
        const wb = createWorkbook();

        /** edited가 있으면 같은 트랜잭션에서 수식 재계산 (부록 I.2 — 공통 후처리 지점) */
        const mutateSheet = (
          sheetId: string,
          fn: (sheet: Sheet, wb: Workbook, state: WorkbookState) => void,
          edited?: CellRange[] | "all",
        ) =>
          set((state) => {
            const sheet = state.workbook.sheets.find((s) => s.id === sheetId);
            if (!sheet) return;
            fn(sheet, state.workbook, state);
            if (edited !== undefined) {
              recalcFormulas(
                state.workbook,
                edited === "all" ? "all" : edited.map((rg) => ({ ...norm(rg), sheetId })),
              );
            }
          });

        return {
          workbook: wb,
          activeSheetId: wb.sheets[0].id,
          selection: null,
          runningBlocks: {},
          flash: null,
          focusBlockId: null,
          dirtyBlocks: {},
          hoverRange: null,
          hoverBlockId: null,
          selectedBlockId: null,
          propsTopTab: "variables" as const,
          propsBottomTab: "files" as const,
          lastEditorBlockId: null,
          anchorPicking: null,
          tocOpen: false,
          aiChatOpen: false,
          propsOpen: false,
          propsPinned: true,
          gridCompact: true,
          gridMaximized: false,
          gridCollapsed: false,
          pyCollapsed: false,
          view: "workbook" as const,
          executedRefs: {},
          showRefs: true,
          chatRefs: [],

          setCellValue: (sheetId, r, c, cell) => {
            const sheet = get().workbook.sheets.find((s) => s.id === sheetId);
            if (!sheet) return false;
            const key = cellKey(r, c);
            if (sheet.cells[key]?.src) return false; // spill 셀은 직접 편집 금지
            mutateSheet(
              sheetId,
              (sh) => {
                if (cell === null) delete sh.cells[key];
                else sh.cells[key] = cell;
                if (r >= sh.rowCount) sh.rowCount = r + 1;
                if (c >= sh.colCount) sh.colCount = c + 1;
              },
              [{ r0: r, c0: c, r1: r, c1: c }],
            );
            return true;
          },

          setCells: (sheetId, edits) =>
            mutateSheet(
              sheetId,
              (sh) => {
                for (const { r, c, cell } of edits) {
                  const key = cellKey(r, c);
                  if (cell === null) delete sh.cells[key];
                  else sh.cells[key] = cell;
                  if (r >= sh.rowCount) sh.rowCount = r + 1;
                  if (c >= sh.colCount) sh.colCount = c + 1;
                }
              },
              edits.map(({ r, c }) => ({ r0: r, c0: c, r1: r, c1: c })),
            ),

          clearRange: (sheetId, range) =>
            mutateSheet(
              sheetId,
              (sh) => {
                const { r0, c0, r1, c1 } = norm(range);
                // ponytail: 저장된 셀 전체 스캔 O(cells) — 범위 인덱스가 필요해지면 교체
                for (const key of Object.keys(sh.cells)) {
                  const { r, c } = parseCellKey(key);
                  if (r >= r0 && r <= r1 && c >= c0 && c <= c1 && !sh.cells[key].src) {
                    delete sh.cells[key];
                  }
                }
              },
              [range],
            ),

          insertRows: (sheetId, index, count) => {
            let changed: string[] = [];
            mutateSheet(sheetId, (sh, _wb, st) => {
              if (count <= 0) return;
              changed = applyStructure(st, { sheetName: sh.name, axis: "row", index, count: count });
              sh.rowCount += count;
              sh.cells = remapCells(sh.cells, (r, c) =>
                r >= index ? [r + count, c] : [r, c],
              );
            }, "all");
            return changed;
          },

          deleteRows: (sheetId, index, count) => {
            let changed: string[] = [];
            mutateSheet(sheetId, (sh, _wb, st) => {
              if (count <= 0) return;
              changed = applyStructure(st, { sheetName: sh.name, axis: "row", index, count: -count });
              sh.rowCount = Math.max(1, sh.rowCount - count);
              sh.cells = remapCells(sh.cells, (r, c) => {
                if (r < index) return [r, c];
                if (r < index + count) return null;
                return [r - count, c];
              });
            }, "all");
            return changed;
          },

          insertCols: (sheetId, index, count) => {
            let changed: string[] = [];
            mutateSheet(sheetId, (sh, _wb, st) => {
              if (count <= 0) return;
              changed = applyStructure(st, { sheetName: sh.name, axis: "col", index, count: count });
              sh.colCount += count;
              sh.cells = remapCells(sh.cells, (r, c) =>
                c >= index ? [r, c + count] : [r, c],
              );
              sh.colWidths = remapWidths(sh.colWidths, (c) =>
                c >= index ? c + count : c,
              );
            }, "all");
            return changed;
          },

          deleteCols: (sheetId, index, count) => {
            let changed: string[] = [];
            mutateSheet(sheetId, (sh, _wb, st) => {
              if (count <= 0) return;
              changed = applyStructure(st, { sheetName: sh.name, axis: "col", index, count: -count });
              sh.colCount = Math.max(1, sh.colCount - count);
              sh.cells = remapCells(sh.cells, (r, c) => {
                if (c < index) return [r, c];
                if (c < index + count) return null;
                return [r, c - count];
              });
              sh.colWidths = remapWidths(sh.colWidths, (c) => {
                if (c < index) return c;
                if (c < index + count) return null;
                return c - count;
              });
              if (sh.frozenCols) {
                // setFrozenCols와 같은 불변식: 최대 colCount - 1
                sh.frozenCols = Math.min(sh.frozenCols, sh.colCount - 1);
              }
            }, "all");
            return changed;
          },

          moveRange: (sheetId, srcRange, dest) => {
            const st = get();
            const sheet = st.workbook.sheets.find((s) => s.id === sheetId);
            if (!sheet) return "시트를 찾을 수 없습니다";
            const src = norm(srcRange);
            const dr = dest.r - src.r0;
            const dc = dest.c - src.c0;
            if (dr === 0 && dc === 0) return null;
            const target = {
              r0: dest.r,
              c0: dest.c,
              r1: dest.r + src.r1 - src.r0,
              c1: dest.c + src.c1 - src.c0,
            };
            for (const rg of [src, target])
              for (let r = rg.r0; r <= rg.r1; r++)
                for (let c = rg.c0; c <= rg.c1; c++)
                  if (sheet.cells[cellKey(r, c)]?.src)
                    return "Python 출력(spill) 셀은 잘라내거나 덮어쓸 수 없습니다";
            mutateSheet(
              sheetId,
              (sh, wb) => {
                const move = (fx: string, own: string) =>
                  moveRefsInFormula(fx, own, sh.name, src, dr, dc);
                rewriteAllFormulas(wb, move);
                for (const n of wb.names ?? []) n.ref = move(`=${n.ref}`, sh.name).slice(1);
                const moved: [number, number, Cell][] = [];
                for (let r = src.r0; r <= src.r1; r++)
                  for (let c = src.c0; c <= src.c1; c++) {
                    const key = cellKey(r, c);
                    if (sh.cells[key]) moved.push([r + dr, c + dc, sh.cells[key]]);
                    delete sh.cells[key];
                  }
                for (let r = target.r0; r <= target.r1; r++)
                  for (let c = target.c0; c <= target.c1; c++) delete sh.cells[cellKey(r, c)];
                for (const [r, c, cell] of moved) sh.cells[cellKey(r, c)] = cell;
                sh.rowCount = Math.max(sh.rowCount, target.r1 + 1);
                sh.colCount = Math.max(sh.colCount, target.c1 + 1);
              },
              [src, target],
            );
            return null;
          },

          defineName: (name, ref) => {
            const trimmed = name.trim();
            const problem = nameProblem(trimmed, ref);
            if (problem) return problem;
            set((state) => {
              const list = (state.workbook.names ??= []);
              const i = list.findIndex((n) => n.name.toUpperCase() === trimmed.toUpperCase());
              if (i >= 0) list[i] = { name: trimmed, ref };
              else list.push({ name: trimmed, ref });
              recalcFormulas(state.workbook, "all");
            });
            return null;
          },

          removeName: (name) =>
            set((state) => {
              const list = state.workbook.names;
              if (!list) return;
              state.workbook.names = list.filter((n) => n.name !== name);
              recalcFormulas(state.workbook, "all");
            }),

          addSheet: () =>
            set((state) => {
              const names = new Set(state.workbook.sheets.map((s) => s.name));
              let i = state.workbook.sheets.length + 1;
              while (names.has(`Sheet${i}`)) i++;
              const sheet = createSheet(`Sheet${i}`);
              state.workbook.sheets.push(sheet);
              state.activeSheetId = sheet.id;
            }),

          addSheetWithCells: (edits) =>
            set((state) => {
              const names = new Set(state.workbook.sheets.map((s) => s.name));
              let i = state.workbook.sheets.length + 1;
              while (names.has(`Sheet${i}`)) i++;
              const sheet = createSheet(`Sheet${i}`);
              for (const { r, c, cell } of edits) {
                if (cell === null) continue;
                sheet.cells[cellKey(r, c)] = cell;
                if (r >= sheet.rowCount) sheet.rowCount = r + 1;
                if (c >= sheet.colCount) sheet.colCount = c + 1;
              }
              state.workbook.sheets.push(sheet);
              state.activeSheetId = sheet.id;
              state.selection = null;
              recalcFormulas(state.workbook, "all"); // 새 시트 이름 참조가 해석될 수 있다
            }),

          renameSheet: (sheetId, name) => {
            const trimmed = name.trim();
            if (trimmed === "") return;
            // 이름 참조("Sheet2!A1") 해석이 바뀌므로 전체 재계산
            mutateSheet(
              sheetId,
              (sh, wb, st) => {
                if (sh.name === trimmed) return;
                const from = sh.name;
                const ren = (fx: string) => renameSheetInFormula(fx, from, trimmed);
                rewriteAllFormulas(wb, ren);
                for (const n of wb.names ?? []) n.ref = ren(`=${n.ref}`).slice(1);
                // 코드의 xl("Sheet2!A1")도 새 이름을 따라간다 (부록 O.4)
                for (const b of wb.pyBlocks) {
                  if (b.kind === "markdown") continue;
                  const code = rewriteXlRefs(b.code, (ref) => ren(`=${ref}`).slice(1));
                  if (code !== b.code) {
                    b.code = code;
                    st.dirtyBlocks[b.id] = true;
                  }
                }
                sh.name = trimmed;
              },
              "all",
            );
          },

          removeSheet: (sheetId) =>
            set((state) => {
              const sheets = state.workbook.sheets;
              if (sheets.length <= 1) return;
              const idx = sheets.findIndex((s) => s.id === sheetId);
              if (idx < 0) return;
              sheets.splice(idx, 1);
              if (state.activeSheetId === sheetId) {
                state.activeSheetId = sheets[Math.max(0, idx - 1)].id;
                state.selection = null;
              }
              recalcFormulas(state.workbook, "all"); // 지워진 시트 참조 → #REF!
            }),

          moveSheet: (sheetId, offset) =>
            set((state) => {
              const sheets = state.workbook.sheets;
              const idx = sheets.findIndex((s) => s.id === sheetId);
              if (idx < 0) return;
              const to = Math.max(0, Math.min(sheets.length - 1, idx + offset));
              if (to === idx) return;
              const [sheet] = sheets.splice(idx, 1);
              sheets.splice(to, 0, sheet);
            }),

          setColWidth: (sheetId, col, width) =>
            mutateSheet(sheetId, (sh) => {
              (sh.colWidths ??= {})[col] = Math.max(30, Math.round(width));
            }),

          setFrozenCols: (sheetId, n) =>
            mutateSheet(sheetId, (sh) => {
              sh.frozenCols = Math.max(0, Math.min(n, sh.colCount - 1));
            }),

          setTitle: (title) => {
            const trimmed = title.trim();
            if (trimmed === "") return;
            set((state) => {
              state.workbook.title = trimmed;
            });
          },

          setSelection: (range) =>
            set((state) => {
              state.selection = range;
            }),

          setActiveSheet: (id) =>
            set((state) => {
              if (state.workbook.sheets.some((s) => s.id === id)) {
                state.activeSheetId = id;
                state.selection = null;
              }
            }),

          newWorkbook: () => {
            const fresh = createWorkbook();
            set((state) => {
              state.workbook = fresh;
              state.activeSheetId = fresh.sheets[0].id;
              state.selection = null;
              state.runningBlocks = {};
              state.dirtyBlocks = {};
              state.executedRefs = {};
              state.chatRefs = [];
              state.selectedBlockId = null;
              state.lastEditorBlockId = null;
              state.hoverBlockId = null;
              state.flash = null;
              state.anchorPicking = null;
            });
            resetHistory();
          },

          loadWorkbook: (loaded) => {
            set((state) => {
              // 구 워크북 정규화: 코드 블록마다 outputs ≥ 1, spill src 태그 이관 (부록 D.1)
              state.workbook = normalizeWorkbook(loaded);
              state.activeSheetId = loaded.sheets[0]?.id ?? "";
              state.selection = null;
              // 이전 워크북의 transient 상태(실행 중·dirty 등) 정리 (§M7.5)
              state.runningBlocks = {};
              state.dirtyBlocks = {};
              state.executedRefs = {};
              state.chatRefs = [];
              state.selectedBlockId = null;
              state.lastEditorBlockId = null;
              state.hoverBlockId = null;
              state.flash = null;
              state.anchorPicking = null;
            });
            resetHistory();
          },

          addPyBlock: (sheetId, anchor, kind) => {
            const st = get();
            const sheet = st.workbook.sheets.find((s) => s.id === sheetId);
            if (!sheet) return null;
            if (sheet.cells[cellKey(anchor.r, anchor.c)]?.src) return null;
            if (
              st.workbook.pyBlocks.some(
                (b) =>
                  b.sheetId === sheetId &&
                  b.anchor.r === anchor.r &&
                  b.anchor.c === anchor.c,
              )
            ) {
              return null;
            }
            const id = newId();
            set((state) => {
              const block: PyBlock = {
                id,
                sheetId,
                anchor,
                code: "",
                outputMode: "values",
                includeIndex: "auto",
                // 마크다운 블록은 실행되지 않고 셀에 아무것도 쓰지 않는다 (앵커 = 위치·목차 대상)
                ...(kind === "markdown" ? { kind, markdown: "" } : {}),
              };
              normalizeBlock(block); // 코드 블록은 출력 1개로 시작한다
              state.workbook.pyBlocks.push(block);
            });
            return id;
          },

          setBlockAnchor: (id, anchor, sheetId) => {
            // 레거시 경로: outputs[0]이 곧 블록 앵커다 (부록 D.1)
            const block = get().workbook.pyBlocks.find((b) => b.id === id);
            if (!block) return "블록을 찾을 수 없습니다";
            const outputId = block.outputs?.[0]?.id;
            if (!outputId) return "출력을 찾을 수 없습니다";
            return get().setOutputAnchor(id, outputId, { sheetId, r: anchor.r, c: anchor.c });
          },

          setOutputAnchor: (blockId, outputId, target) => {
            const st = get();
            const block = st.workbook.pyBlocks.find((b) => b.id === blockId);
            if (!block) return "블록을 찾을 수 없습니다";
            const binding = block.outputs?.find((o) => o.id === outputId);
            if (!binding) return "출력을 찾을 수 없습니다";
            const targetSheetId = target.sheetId ?? binding.sheetId ?? block.sheetId;
            const sheet = st.workbook.sheets.find((s) => s.id === targetSheetId);
            if (!sheet) return "시트를 찾을 수 없습니다";
            const currentSheetId = binding.sheetId ?? block.sheetId;
            if (
              targetSheetId === currentSheetId &&
              binding.anchor.r === target.r &&
              binding.anchor.c === target.c
            ) {
              return null; // 제자리
            }
            const tag = srcTag(blockId, outputId);
            const conflict = checkSpillConflict(sheet, st.workbook.pyBlocks, tag, target, [1, 1]);
            if (conflict) return conflict;
            const cell = sheet.cells[cellKey(target.r, target.c)];
            // 앵커 셀은 출력 소유라 checkSpillConflict가 봐주지만, 재지정은 빈 셀에만 허용한다
            if (cell && !cell.src && cell.v !== null && cell.v !== "") {
              return `비어 있지 않은 셀(${colToLetter(target.c)}${target.r + 1})과 겹칩니다`;
            }
            set((state) => {
              const b = state.workbook.pyBlocks.find((x) => x.id === blockId);
              const o = b?.outputs?.find((x) => x.id === outputId);
              if (!b || !o) return;
              const cleared = clearSpillCells(state.workbook, blockId, outputId);
              if (cleared.length > 0) recalcFormulas(state.workbook, cleared);
              if (o.last?.spillRange) delete o.last.spillRange; // 옛 위치의 spill 테두리 제거
              o.anchor = { r: target.r, c: target.c };
              if (b.outputs![0].id === outputId) {
                // 블록 시트가 따라 움직인다 — 다른 출력은 원래 시트에 남긴다
                if (targetSheetId !== b.sheetId) {
                  for (const other of b.outputs!) {
                    if (other.id !== outputId) other.sheetId ??= b.sheetId;
                  }
                  b.sheetId = targetSheetId;
                }
                delete o.sheetId;
              } else if (targetSheetId === b.sheetId) {
                delete o.sheetId;
              } else {
                o.sheetId = targetSheetId;
              }
              syncLegacy(b);
              state.dirtyBlocks[blockId] = true;
              state.anchorPicking = null;
            });
            return null;
          },

          swapBlockOrder: (id, direction) => {
            const ordered = blocksInOrder(get().workbook);
            const i = ordered.findIndex((b) => b.id === id);
            const j = direction === "up" ? i - 1 : i + 1;
            if (i < 0 || j < 0 || j >= ordered.length) return null; // 경계
            const otherId = ordered[j].id;
            set((state) => {
              const a = state.workbook.pyBlocks.find((b) => b.id === id);
              const b = state.workbook.pyBlocks.find((x) => x.id === otherId);
              if (!a || !b) return;
              const cleared: SheetRange[] = [];
              for (const blk of [a, b]) {
                cleared.push(...clearSpillCells(state.workbook, blk.id));
                for (const o of blk.outputs ?? []) {
                  if (o.last?.spillRange) delete o.last.spillRange;
                }
                if (blk.last?.spillRange) delete blk.last.spillRange;
                state.dirtyBlocks[blk.id] = true;
              }
              // draft 별칭을 피하려고 평범한 값으로 먼저 복사한다
              const posA = { sheetId: a.sheetId, anchor: { ...a.anchor } };
              const posB = { sheetId: b.sheetId, anchor: { ...b.anchor } };
              // 다른 시트로 옮기는 블록의 나머지 출력은 원래 시트에 남는다
              for (const blk of [a, b]) {
                for (const o of (blk.outputs ?? []).slice(1)) o.sheetId ??= blk.sheetId;
              }
              a.sheetId = posB.sheetId;
              a.anchor = { ...posB.anchor };
              b.sheetId = posA.sheetId;
              b.anchor = { ...posA.anchor };
              for (const blk of [a, b]) {
                const first = blk.outputs?.[0];
                if (first) {
                  first.anchor = { ...blk.anchor };
                  delete first.sheetId;
                }
              }
              if (cleared.length > 0) recalcFormulas(state.workbook, cleared);
            });
            return otherId;
          },

          removePyBlock: (id) =>
            set((state) => {
              const idx = state.workbook.pyBlocks.findIndex((b) => b.id === id);
              if (idx < 0) return;
              const cleared = clearSpillCells(state.workbook, id); // 다른 시트에 놓인 출력까지
              state.workbook.pyBlocks.splice(idx, 1);
              if (cleared.length > 0) recalcFormulas(state.workbook, cleared);
              delete state.runningBlocks[id];
              delete state.dirtyBlocks[id];
              delete state.executedRefs[id];
              if (state.focusBlockId === id) state.focusBlockId = null;
              if (state.selectedBlockId === id) state.selectedBlockId = null;
              if (state.lastEditorBlockId === id) state.lastEditorBlockId = null;
              if (state.hoverBlockId === id) state.hoverBlockId = null;
              if (state.anchorPicking?.blockId === id) state.anchorPicking = null;
            }),

          setBlockCode: (id, code) =>
            set((state) => {
              const block = state.workbook.pyBlocks.find((b) => b.id === id);
              if (block && block.code !== code) block.code = code;
            }),

          setBlockOutputMode: (id, mode) => {
            const outputId = get().workbook.pyBlocks.find((b) => b.id === id)?.outputs?.[0]?.id;
            if (outputId) get().setOutputMode(id, outputId, mode);
          },

          setBlockOutput: (id, patch) => {
            const outputId = get().workbook.pyBlocks.find((b) => b.id === id)?.outputs?.[0]?.id;
            if (outputId) get().setOutputSelection(id, outputId, patch);
          },

          addOutput: (blockId) => {
            const st = get();
            const block = st.workbook.pyBlocks.find((b) => b.id === blockId);
            if (!block || block.kind === "markdown") return null;
            const sheet = st.workbook.sheets.find((s) => s.id === block.sheetId);
            if (!sheet) return null;
            const r = block.anchor.r;
            const c = freeOutputCol(sheet, st.workbook.pyBlocks, block);
            const id = newId();
            set((state) => {
              const b = state.workbook.pyBlocks.find((x) => x.id === blockId);
              if (!b) return;
              (b.outputs ??= []).push({
                id,
                anchor: { r, c },
                mode: "values",
                includeIndex: "auto",
              });
              state.dirtyBlocks[blockId] = true;
            });
            return id;
          },

          addOutputs: (blockId, specs, start) => {
            const st = get();
            const block = st.workbook.pyBlocks.find((b) => b.id === blockId);
            if (!block || block.kind === "markdown" || specs.length === 0) return [];
            const sheetId = start?.sheetId ?? block.sheetId;
            const sheet = st.workbook.sheets.find((s) => s.id === sheetId);
            if (!sheet) return [];
            const r = start ? start.r : block.anchor.r;
            let c = start ? start.c : freeOutputCol(sheet, st.workbook.pyBlocks, block);
            const ids: string[] = [];
            set((state) => {
              const b = state.workbook.pyBlocks.find((x) => x.id === blockId);
              if (!b) return;
              for (const spec of specs) {
                const id = newId();
                ids.push(id);
                (b.outputs ??= []).push({
                  id,
                  ...(sheetId === b.sheetId ? {} : { sheetId }),
                  anchor: { r, c },
                  mode: "values",
                  includeIndex: spec.includeIndex ?? "auto",
                  selection: spec.selection,
                  label: spec.label,
                });
                c += Math.max(1, spec.width) + 1; // 한 칸 띄워 다음 출력
              }
              state.dirtyBlocks[blockId] = true;
            });
            return ids;
          },

          removeOutput: (blockId, outputId) =>
            set((state) => {
              const b = state.workbook.pyBlocks.find((x) => x.id === blockId);
              if (!b?.outputs || b.outputs.length <= 1) return; // 마지막 출력은 남긴다
              const i = b.outputs.findIndex((o) => o.id === outputId);
              if (i < 0) return;
              const cleared = clearSpillCells(state.workbook, blockId, outputId);
              b.outputs.splice(i, 1);
              syncLegacy(b);
              if (cleared.length > 0) recalcFormulas(state.workbook, cleared);
              state.dirtyBlocks[blockId] = true;
            }),

          setOutputSelection: (blockId, outputId, patch) =>
            set((state) => {
              const b = state.workbook.pyBlocks.find((x) => x.id === blockId);
              const o = b?.outputs?.find((x) => x.id === outputId);
              if (!b || !o) return;
              const next: OutputSelection = { ...o.selection, ...patch };
              for (const k of Object.keys(next) as (keyof OutputSelection)[]) {
                if (next[k] === undefined) delete next[k];
              }
              o.selection = Object.keys(next).length > 0 ? next : undefined;
              syncLegacy(b);
            }),

          setOutputMode: (blockId, outputId, mode) =>
            set((state) => {
              const b = state.workbook.pyBlocks.find((x) => x.id === blockId);
              const o = b?.outputs?.find((x) => x.id === outputId);
              // 객체→값 전환의 spill 충돌은 다음 실행에서 검사한다 (§2.3.6, M5)
              if (!b || !o) return;
              o.mode = mode;
              syncLegacy(b);
            }),

          setOutputIncludeIndex: (blockId, outputId, value) =>
            set((state) => {
              const b = state.workbook.pyBlocks.find((x) => x.id === blockId);
              const o = b?.outputs?.find((x) => x.id === outputId);
              if (!b || !o) return;
              o.includeIndex = value;
              syncLegacy(b);
            }),

          setOutputLabel: (blockId, outputId, label) =>
            set((state) => {
              const o = state.workbook.pyBlocks
                .find((x) => x.id === blockId)
                ?.outputs?.find((x) => x.id === outputId);
              if (o) o.label = label.trim() === "" ? undefined : label;
            }),

          setBlockMarkdown: (id, markdown) =>
            set((state) => {
              const block = state.workbook.pyBlocks.find((b) => b.id === id);
              if (!block || block.markdown === markdown) return;
              block.markdown = markdown;
              block.title = markdownTitle(markdown) || undefined; // 첫 헤딩 = 목차 제목
            }),

          setBlockNote: (id, note) =>
            set((state) => {
              const block = state.workbook.pyBlocks.find((b) => b.id === id);
              if (!block || block.kind === "markdown") return;
              if (note === null) delete block.note;
              else if (block.note !== note) block.note = note;
            }),

          setBlockTitle: (id, title) =>
            set((state) => {
              const block = state.workbook.pyBlocks.find((b) => b.id === id);
              if (block) block.title = title.trim() === "" ? undefined : title;
            }),

          setBlockCollapsed: (id, collapsed) =>
            set((state) => {
              const block = state.workbook.pyBlocks.find((b) => b.id === id);
              if (block) block.collapsed = collapsed || undefined;
            }),

          setAllCollapsed: (collapsed) =>
            set((state) => {
              for (const b of state.workbook.pyBlocks) b.collapsed = collapsed || undefined;
            }),

          patchOutputImage: (blockId, outputId, imageBlobId) =>
            set((state) => {
              const block = state.workbook.pyBlocks.find((b) => b.id === blockId);
              const binding = block?.outputs?.find((o) => o.id === outputId);
              if (!binding?.last) return;
              binding.last.imageBlobId = imageBlobId;
              if (block && block.outputs?.[0]?.id === outputId && block.last) {
                block.last.imageBlobId = imageBlobId; // 레거시 뷰 동기화
              }
            }),
          applyOutputResults: (blockId, results) =>
            set((state) => {
              const block = state.workbook.pyBlocks.find((b) => b.id === blockId);
              if (!block) return;
              const edited: SheetRange[] = []; // spill 반영도 수식 재계산 대상 (부록 I.2)
              for (const res of results) {
                const binding = block.outputs?.find((o) => o.id === res.outputId);
                if (!binding) continue;
                const sheet = state.workbook.sheets.find(
                  (s) => s.id === (binding.sheetId ?? block.sheetId),
                );
                if (!sheet) continue;
                const tag = srcTag(blockId, binding.id);
                if (res.clearPrevious) {
                  let box: SheetRange | null = null;
                  for (const key of Object.keys(sheet.cells)) {
                    if (sheet.cells[key].src !== tag) continue;
                    delete sheet.cells[key];
                    const { r, c } = parseCellKey(key);
                    box = box
                      ? {
                          sheetId: sheet.id,
                          r0: Math.min(box.r0, r),
                          c0: Math.min(box.c0, c),
                          r1: Math.max(box.r1, r),
                          c1: Math.max(box.c1, c),
                        }
                      : { sheetId: sheet.id, r0: r, c0: c, r1: r, c1: c };
                  }
                  if (box) edited.push(box);
                }
                res.cells.forEach((row, i) =>
                  row.forEach((cell, j) => {
                    const r = binding.anchor.r + i;
                    const c = binding.anchor.c + j;
                    sheet.cells[cellKey(r, c)] = { ...cell, src: tag };
                    if (r >= sheet.rowCount) sheet.rowCount = r + 1;
                    if (c >= sheet.colCount) sheet.colCount = c + 1;
                  }),
                );
                if (res.cells.length > 0 && res.cells[0].length > 0) {
                  edited.push({
                    sheetId: sheet.id,
                    r0: binding.anchor.r,
                    c0: binding.anchor.c,
                    r1: binding.anchor.r + res.cells.length - 1,
                    c1: binding.anchor.c + res.cells[0].length - 1,
                  });
                }
                if (res.last) binding.last = res.last;
              }
              syncLegacy(block);
              if (edited.length > 0) recalcFormulas(state.workbook, edited);
            }),

          applyBlockResult: (blockId, cells, opts) => {
            const outputId = get().workbook.pyBlocks.find((b) => b.id === blockId)?.outputs?.[0]
              ?.id;
            if (!outputId) return;
            get().applyOutputResults(blockId, [
              { outputId, cells, clearPrevious: opts?.clearPrevious, last: opts?.last },
            ]);
          },

          setBlockRunning: (id, running) =>
            set((state) => {
              if (running) state.runningBlocks[id] = true;
              else delete state.runningBlocks[id];
            }),

          setFlash: (flash) =>
            set((state) => {
              state.flash = flash;
            }),

          setFocusBlock: (id) =>
            set((state) => {
              state.focusBlockId = id;
            }),

          markDirty: (ids) =>
            set((state) => {
              for (const id of ids) state.dirtyBlocks[id] = true;
            }),

          clearDirty: (id) =>
            set((state) => {
              delete state.dirtyBlocks[id];
            }),

          setCalcMode: (mode) =>
            set((state) => {
              state.workbook.calcMode = mode;
            }),

          setInitScript: (script) =>
            set((state) => {
              state.workbook.initScript = script;
            }),

          setHoverRange: (hover) =>
            set((state) => {
              state.hoverRange = hover;
            }),

          setHoverBlock: (id) =>
            set((state) => {
              state.hoverBlockId = id;
            }),

          setSelectedBlock: (id) =>
            set((state) => {
              state.selectedBlockId = id;
            }),

          showPanelTab: (tab) =>
            set((state) => {
              if (tab === "files" || tab === "preview") state.propsBottomTab = tab;
              else state.propsTopTab = tab;
              state.propsOpen = true;
            }),

          setLastEditorBlock: (id) =>
            set((state) => {
              state.lastEditorBlockId = id;
            }),

          setAnchorPicking: (target) =>
            set((state) => {
              state.anchorPicking = target;
            }),

          applyCellStyle: (sheetId, range, patch) =>
            mutateSheet(sheetId, (sh) => {
              const { r0, c0, r1, c1 } = norm(range);
              for (let r = r0; r <= r1; r++) {
                for (let c = c0; c <= c1; c++) {
                  const key = cellKey(r, c);
                  const cell = sh.cells[key];
                  if (cell?.src) continue; // spill 잠금 셀 제외
                  const st = { ...cell?.st };
                  if (patch.b !== undefined) {
                    if (patch.b) st.b = true;
                    else delete st.b;
                  }
                  if (patch.fs !== undefined) {
                    if (patch.fs === null) delete st.fs;
                    else st.fs = patch.fs;
                  }
                  const hasSt = st.b !== undefined || st.fs !== undefined;
                  if (cell) {
                    if (hasSt) cell.st = st;
                    else if (cell.v === null && !cell.fx) delete sh.cells[key]; // 서식만 있던 빈 셀
                    else delete cell.st;
                  } else if (hasSt) {
                    sh.cells[key] = { v: null, t: "s", st }; // 빈 셀에도 서식 적용 (엑셀 동일)
                  }
                }
              }
            }),

          setExecutedRefs: (blockId, refs) =>
            set((state) => {
              if (refs === null) delete state.executedRefs[blockId];
              else state.executedRefs[blockId] = refs;
            }),

          setShowRefs: (show) =>
            set((state) => {
              state.showRefs = show;
            }),

          setChatRefs: (refs) =>
            set((state) => {
              state.chatRefs = refs;
            }),

          setTocOpen: (open) =>
            set((state) => {
              state.tocOpen = open;
            }),

          setPropsOpen: (open) =>
            set((state) => {
              state.propsOpen = open;
            }),

          setPropsPinned: (pinned) =>
            set((state) => {
              state.propsPinned = pinned;
            }),

          setGridCompact: (compact) =>
            set((state) => {
              state.gridCompact = compact;
            }),

          setGridMaximized: (max) =>
            set((state) => {
              state.gridMaximized = max;
              if (max) state.gridCollapsed = false; // 접힌 채로는 전체 화면이 될 수 없다
            }),

          setWorkDir: (dir) =>
            set((state) => {
              if (dir === null) delete state.workbook.workDir;
              else state.workbook.workDir = dir;
            }),

          setAiChatOpen: (open) =>
            set((state) => {
              state.aiChatOpen = open;
            }),

          setPanelCollapsed: (panel, collapsed) =>
            set((state) => {
              if (panel === "grid") {
                state.gridCollapsed = collapsed;
                if (collapsed) state.pyCollapsed = false; // 마지막 하나는 남긴다
              } else {
                state.pyCollapsed = collapsed;
                if (collapsed) state.gridCollapsed = false;
              }
            }),

          setView: (view) =>
            set((state) => {
              state.view = view;
            }),
        };
      }),
      {
        limit: 100,
        partialize,
        equality: (a, b) => a === b,
        handleSet: (handle) => {
          const throttled = throttle(handle as (...args: unknown[]) => void, 300, {
            leading: true,
            trailing: true,
          });
          cancelPending = () => throttled.cancel();
          return throttled;
        },
      },
    ),
  );

  resetHistory = () => {
    cancelPending(); // 대기 중인 trailing push가 초기화 후 이력을 오염시키지 않도록
    store.temporal.getState().clear();
  };

  return store;
};

export const useWorkbookStore = createWorkbookStore();
