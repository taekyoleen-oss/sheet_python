// lib/runtime/calc-engine.ts — 의존성 그래프·위상 정렬·dirty 전파·실행 조율 (설계서 §2.4)
// 순수 TS. 스토어를 import하지 않는다 — grid-ui는 CalcHost 콜백으로 반영한다.
// API 문서: /output/calc-engine-api.md

import { parseA1 } from "@/lib/grid/a1";
import type {
  BlockKind,
  OutputBinding,
  Cell,
  CellRange,
  CellType,
  IncludeIndex,
  OutputMode,
  OutputSelection,
} from "@/types/workbook";
import { spillRange, toCells } from "./converters";
import type { OutputItem, OutputRequest, RangeSnapshot, RunPayload } from "./protocol";

// ── 데이터 형태 ──────────────────────────────────────────

export interface CalcBlock {
  id: string;
  sheetId: string;
  anchor: { r: number; c: number };
  code: string;
  outputMode: OutputMode;
  includeIndex: IncludeIndex;
  /** 출력 선택(변수·열·행). 워커 run 메시지로 그대로 전달된다 */
  output?: OutputSelection;
  /** 'markdown' 블록은 실행 대상이 아니다 — 큐·그래프에서 제외 */
  kind?: BlockKind;
  /** 다중 출력. 있으면 outputMode/includeIndex/output 대신 이쪽이 정본 */
  outputs?: OutputBinding[];
}

/** 블록이 점유한 출력 영역 (sheetId 생략 시 블록 시트) */
export type OutputArea = CellRange & { sheetId?: string };

export interface SheetRange extends CellRange {
  sheetId: string;
}

export interface CalcGraph {
  /** id → 선행 블록(내가 의존하는) */
  deps: Map<string, Set<string>>;
  /** id → 후행 블록(나에게 의존하는) */
  dependents: Map<string, Set<string>>;
}

/** 실행 스냅샷: 스토어의 현재 상태를 평면 데이터로 넘긴다 */
export interface WorkbookView {
  blocks: CalcBlock[];
  /** 시트 id, 표시 순서대로 (동순위 tie-break 기준) */
  sheetOrder: string[];
  /** 값 모드 블록의 현재 spill 범위. 없으면 앵커 1×1로 간주 */
  spills: Map<string, CellRange | undefined>;
  /**
   * 다중 출력 블록이 점유한 영역들(값 모드 spill 범위, 객체 모드 앵커 1×1).
   * 있으면 spills보다 우선한다 — 의존성 대상은 이 영역 전부다.
   */
  areas?: Map<string, OutputArea[]>;
}

/** 마크다운 블록은 실행 대상이 아니다 — 그래프·순서·dirty·실행 큐 전부에서 제외한다 */
export const isExecutable = (b: CalcBlock): boolean => b.kind !== "markdown";

export const overlaps = (a: SheetRange, b: SheetRange): boolean =>
  a.sheetId === b.sheetId && a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;

/** analyze가 돌려준 ref 문자열들 → 시트가 확정된 범위. 잘못된 참조는 조용히 건너뛴다
 *  (그래프용 — 실행 시 buildSnapshots가 같은 참조를 다시 만나 한국어 오류로 보고한다) */
export function resolveRefs(
  refs: string[],
  ownSheetId: string,
  resolveSheet: (name: string) => string | undefined,
): SheetRange[] {
  const out: SheetRange[] = [];
  for (const ref of refs) {
    try {
      const p = parseA1(ref);
      const sheetId = p.sheetName === undefined ? ownSheetId : resolveSheet(p.sheetName);
      if (sheetId !== undefined) out.push({ ...p.range, sheetId });
    } catch {
      // 무시
    }
  }
  return out;
}

// ── a. 그래프 ────────────────────────────────────────────

/** B의 참조가 A의 출력 영역(다중 출력이면 전부, 아니면 spill 범위/앵커 셀)과 겹치면 B는 A에 의존 (§2.4) */
export function buildGraph(
  all: CalcBlock[],
  resolved: Map<string, SheetRange[]>,
  spills: Map<string, CellRange | undefined>,
  areas?: Map<string, OutputArea[]>,
): CalcGraph {
  const blocks = all.filter(isExecutable); // 마크다운은 노드가 아니다(참조 대상도 아님)
  const targets = (b: CalcBlock): SheetRange[] => {
    const own = areas?.get(b.id);
    // 다중 출력: 등록된 영역 전부가 의존 대상 (areas가 spills를 이긴다)
    if (own && own.length > 0) return own.map((a) => ({ ...a, sheetId: a.sheetId ?? b.sheetId }));
    const spill = b.outputMode === "values" ? spills.get(b.id) : undefined;
    const r = spill ?? { r0: b.anchor.r, c0: b.anchor.c, r1: b.anchor.r, c1: b.anchor.c };
    return [{ ...r, sheetId: b.sheetId }];
  };
  const deps = new Map(blocks.map((b) => [b.id, new Set<string>()]));
  const dependents = new Map(blocks.map((b) => [b.id, new Set<string>()]));
  // ponytail: O(블록² × 참조) 전수 비교 — 블록 수백 개 규모라 무해. 병목이 되면 시트별 인덱스로
  for (const a of blocks) {
    const ts = targets(a);
    for (const b of blocks) {
      if (a.id === b.id) continue;
      if ((resolved.get(b.id) ?? []).some((ref) => ts.some((t) => overlaps(ref, t)))) {
        deps.get(b.id)!.add(a.id);
        dependents.get(a.id)!.add(b.id);
      }
    }
  }
  return { deps, dependents };
}

// ── b. 위상 정렬 + 순환 ──────────────────────────────────

const makeCmp = (blocks: CalcBlock[], sheetOrder: string[]) => {
  const byId = new Map(blocks.map((b) => [b.id, b]));
  const sheetIdx = new Map(sheetOrder.map((s, i) => [s, i]));
  return (x: string, y: string): number => {
    const a = byId.get(x)!;
    const b = byId.get(y)!;
    return (
      (sheetIdx.get(a.sheetId) ?? Number.MAX_SAFE_INTEGER) -
        (sheetIdx.get(b.sheetId) ?? Number.MAX_SAFE_INTEGER) ||
      a.anchor.r - b.anchor.r ||
      a.anchor.c - b.anchor.c ||
      x.localeCompare(y)
    );
  };
};

/** 순환에 속한 블록 전부. 양방향 가지치기 후 자기 도달 검사 */
function cycleMembers(blocks: CalcBlock[], graph: CalcGraph): Set<string> {
  const alive = new Set(blocks.map((b) => b.id));
  const inAlive = (set?: Set<string>) => [...(set ?? [])].filter((x) => alive.has(x));
  let changed = true;
  while (changed) {
    changed = false;
    for (const id of [...alive]) {
      if (inAlive(graph.deps.get(id)).length === 0 || inAlive(graph.dependents.get(id)).length === 0) {
        alive.delete(id);
        changed = true;
      }
    }
  }
  // ponytail: 후보마다 DFS 자기 도달 — O(n·e). 워크북 블록 규모라 충분, 커지면 Tarjan SCC로
  const out = new Set<string>();
  for (const id of alive) {
    const seen = new Set<string>();
    const stack = inAlive(graph.dependents.get(id));
    while (stack.length) {
      const x = stack.pop()!;
      if (x === id) {
        out.add(id);
        break;
      }
      if (!seen.has(x)) {
        seen.add(x);
        stack.push(...inAlive(graph.dependents.get(x)));
      }
    }
  }
  return out;
}

/** Kahn 위상 정렬. 동순위는 (시트 순, 앵커 행, 열). 순환 구성원은 order에서 제외된다.
 *  순환의 하위(순환에 의존하지만 순환은 아닌) 블록은 순환 의존을 무시하고 order에 포함된다
 *  — 실행 시 sheet() 안전망 또는 이전 spill 값으로 진행된다(문서 참조) */
export function calcOrder(
  all: CalcBlock[],
  sheetOrder: string[],
  graph: CalcGraph,
): { order: string[]; cycle: string[] } {
  const blocks = all.filter(isExecutable);
  const cmp = makeCmp(blocks, sheetOrder);
  const cycle = cycleMembers(blocks, graph);
  const aliveIds = blocks.map((b) => b.id).filter((id) => !cycle.has(id));
  const aliveSet = new Set(aliveIds);
  const indeg = new Map(
    aliveIds.map((id) => [
      id,
      [...(graph.deps.get(id) ?? [])].filter((d) => aliveSet.has(d)).length,
    ]),
  );
  const ready = aliveIds.filter((id) => indeg.get(id) === 0);
  const order: string[] = [];
  while (ready.length) {
    ready.sort(cmp);
    const id = ready.shift()!;
    order.push(id);
    for (const d of graph.dependents.get(id) ?? []) {
      if (!indeg.has(d)) continue;
      const n = indeg.get(d)! - 1;
      indeg.set(d, n);
      if (n === 0) ready.push(d);
    }
  }
  return { order, cycle: [...cycle].sort(cmp) };
}

// ── c. dirty 전파 ────────────────────────────────────────

/** 편집 범위와 참조가 겹치는 블록 + 코드 수정 블록 + 그 하위 전부 (§2.4) */
export function dirtyPropagation(
  resolved: Map<string, SheetRange[]>,
  graph: CalcGraph,
  editedRanges: SheetRange[],
  editedBlockIds: string[],
): Set<string> {
  // graph에 없는 id = 실행 대상 아님(마크다운 블록 등) → dirty 배지에서도 제외
  const known = (id: string): boolean => graph.deps.has(id);
  const dirty = new Set(editedBlockIds.filter(known));
  for (const [id, refs] of resolved) {
    if (!known(id)) continue;
    if (refs.some((ref) => editedRanges.some((e) => overlaps(ref, e)))) dirty.add(id);
  }
  const stack = [...dirty];
  while (stack.length) {
    const id = stack.pop()!;
    for (const d of graph.dependents.get(id) ?? []) {
      if (!dirty.has(d)) {
        dirty.add(d);
        stack.push(d);
      }
    }
  }
  return dirty;
}

// ── 실행 조율 ────────────────────────────────────────────

/** grid-ui가 구현하는 반영 인터페이스 (스토어 트랜잭션은 grid-ui 몫) */
export interface CalcHost {
  /** 스냅샷 구성용 셀 조회. 빈 셀 → undefined */
  getCell(
    sheetId: string,
    r: number,
    c: number,
  ): { v: string | number | boolean | null; t: CellType } | undefined;
  /** 시트 이름 → 시트 id. 없으면 undefined */
  resolveSheet(name: string): string | undefined;
  /** 실행 대기열 진입 — '#BUSY!' 표시 */
  onBusy(blockId: string): void;
  /** 실행 결과(성공·실패 공통). cells/spill은 값 모드 성공에만 있다 */
  onResult(
    blockId: string,
    payload: RunPayload,
    cells: Cell[][] | null,
    spill: CellRange | null,
  ): void;
  /**
   * 다중 출력 결과 적용 (block.outputs가 있고 이 콜백을 구현했을 때 onResult 대신 호출된다).
   * items[i].cells/spill은 값 모드 성공일 때만 채워진다. 실행 자체가 실패하면 items는 빈 배열.
   */
  onOutputs?(
    blockId: string,
    payload: RunPayload,
    items: Array<{
      outputId: string;
      item: OutputItem;
      cells: Cell[][] | null;
      spill: CellRange | null;
    }>,
  ): void;
}

/** RuntimeClient가 구조적으로 만족하는 최소 표면 (테스트 mock용) */
export interface RuntimeLike {
  analyze(code: string): Promise<string[]>;
  run(
    blockId: string,
    code: string,
    snapshots: Record<string, RangeSnapshot>,
    outputMode: OutputMode,
    includeIndex: IncludeIndex,
    timeoutSec?: number,
    output?: OutputSelection,
    outputs?: OutputRequest[],
  ): Promise<RunPayload>;
}

const failure = (errorType: string, message: string): RunPayload => ({
  ok: false,
  errorType,
  message,
  traceback: "",
  stdout: "",
  stderr: "",
  durationMs: 0,
});

export class RunCoordinator {
  /** 자동(기본)/수동 — 수동이면 notifyEdit는 무시된다(§2.4, dirty 배지는 grid-ui가 dirtyPropagation으로) */
  mode: "auto" | "manual" = "auto";
  /** 블록 실행 타임아웃(초). undefined면 클라이언트 기본값(60) */
  timeoutSec?: number;
  /** 자동 모드 디바운스(ms) — 연속 셀 편집을 묶는다 */
  debounceMs = 500;

  private queue: Promise<void> = Promise.resolve();
  // ponytail: 코드 문자열 자체가 캐시 키, 무한 보관 — 블록 수 규모라 무해. 커지면 LRU로.
  // 실패도 캐시한다(분석은 코드에 결정론적). 일시 오류(워커 재부트 중 등)까지 캐시되는
  // 한계는 코드 수정 시 키가 바뀌므로 실사용에서 자연 해소된다.
  private analyzeCache = new Map<string, string[] | Error>();
  private debounceTimer: ReturnType<typeof setTimeout> | undefined;
  private pendingRanges: SheetRange[] = [];
  private pendingBlockIds: string[] = [];
  private pendingView: WorkbookView | null = null;

  constructor(
    private client: RuntimeLike,
    private host: CalcHost,
  ) {}

  /** 모든 예약된 실행이 끝날 때까지 (테스트·저장 전 대기용) */
  whenIdle(): Promise<void> {
    return this.queue;
  }

  /** 셀 편집/코드 수정 통지. 자동 모드에서만 디바운스 후 dirty 블록을 실행한다 */
  notifyEdit(editedRanges: SheetRange[], editedBlockIds: string[], view: WorkbookView): void {
    if (this.mode !== "auto") return;
    this.pendingRanges.push(...editedRanges);
    this.pendingBlockIds.push(...editedBlockIds);
    this.pendingView = view;
    if (this.debounceTimer) clearTimeout(this.debounceTimer);
    this.debounceTimer = setTimeout(() => {
      const ranges = this.pendingRanges;
      const blockIds = this.pendingBlockIds;
      const v = this.pendingView!;
      this.pendingRanges = [];
      this.pendingBlockIds = [];
      this.pendingView = null;
      this.enqueue(async () => {
        const { resolved, graph } = await this.prepare(v);
        const dirty = dirtyPropagation(resolved, graph, ranges, blockIds);
        const { order, cycle } = calcOrder(v.blocks, v.sheetOrder, graph);
        await this.execute(v, order.filter((id) => dirty.has(id)), cycle);
      });
    }, this.debounceMs);
  }

  /** 전체 실행 — 변수 공유만으로 이어진 블록의 순서는 여기서만 보장된다(§2.4) */
  runAll(view: WorkbookView): Promise<void> {
    return this.enqueue(async () => {
      const { graph } = await this.prepare(view);
      const { order, cycle } = calcOrder(view.blocks, view.sheetOrder, graph);
      await this.execute(view, order, cycle);
    });
  }

  /** ▶ 실행: 지정 블록 + 하위 의존 블록을 위상 순서로 */
  runBlocks(seedIds: string[], view: WorkbookView): Promise<void> {
    return this.enqueue(async () => {
      const { resolved, graph } = await this.prepare(view);
      const dirty = dirtyPropagation(resolved, graph, [], seedIds);
      const { order, cycle } = calcOrder(view.blocks, view.sheetOrder, graph);
      await this.execute(view, order.filter((id) => dirty.has(id)), cycle);
    });
  }

  // ── 내부 ──────────────────────────────────────────────

  private enqueue(fn: () => Promise<void>): Promise<void> {
    this.queue = this.queue.then(fn).catch(() => undefined);
    return this.queue;
  }

  private async analyzeCached(code: string): Promise<string[]> {
    const hit = this.analyzeCache.get(code);
    if (hit !== undefined) {
      if (hit instanceof Error) throw hit;
      return hit;
    }
    try {
      const refs = await this.client.analyze(code);
      this.analyzeCache.set(code, refs);
      return refs;
    } catch (e) {
      const err = e instanceof Error ? e : new Error(String(e));
      this.analyzeCache.set(code, err);
      throw err;
    }
  }

  private async prepare(
    view: WorkbookView,
  ): Promise<{ resolved: Map<string, SheetRange[]>; graph: CalcGraph }> {
    const resolved = new Map<string, SheetRange[]>();
    for (const b of view.blocks) {
      if (!isExecutable(b)) continue; // 마크다운 코드는 워커로 보내지 않는다
      try {
        resolved.set(
          b.id,
          resolveRefs(await this.analyzeCached(b.code), b.sheetId, (n) =>
            this.host.resolveSheet(n),
          ),
        );
      } catch {
        resolved.set(b.id, []); // 분석 실패 블록: 의존성 없음. 실행 단계에서 오류로 보고된다
      }
    }
    return { resolved, graph: buildGraph(view.blocks, resolved, view.spills, view.areas) };
  }

  private buildSnapshots(refs: string[], ownSheetId: string): Record<string, RangeSnapshot> {
    const out: Record<string, RangeSnapshot> = {};
    for (const ref of refs) {
      const p = parseA1(ref); // A1Error(한국어) → 호출부에서 보고
      const sheetId = p.sheetName === undefined ? ownSheetId : this.host.resolveSheet(p.sheetName);
      if (sheetId === undefined) throw new Error(`시트를 찾을 수 없습니다: ${p.sheetName}`);
      const values: (string | number | boolean | null)[][] = [];
      const types: CellType[][] = [];
      for (let r = p.range.r0; r <= p.range.r1; r++) {
        const vr: (string | number | boolean | null)[] = [];
        const tr: CellType[] = [];
        for (let c = p.range.c0; c <= p.range.c1; c++) {
          const cell = this.host.getCell(sheetId, r, c);
          vr.push(cell ? cell.v : null);
          tr.push(cell ? cell.t : "s");
        }
        values.push(vr);
        types.push(tr);
      }
      out[ref] = { values, types, scalar: p.scalar };
    }
    return out;
  }

  private async execute(view: WorkbookView, order: string[], cycle: string[]): Promise<void> {
    // 순환 구성원 전부 오류 표시 + 실행 큐 제외 (§2.4)
    for (const id of cycle) {
      this.host.onResult(id, failure("PyGridCycleError", "순환 참조"), null, null);
    }
    for (const id of order) this.host.onBusy(id);
    const byId = new Map(view.blocks.map((b) => [b.id, b]));
    for (const id of order) {
      const block = byId.get(id);
      if (!block) continue;
      let refs: string[];
      try {
        refs = await this.analyzeCached(block.code);
      } catch (e) {
        // sheet() 비리터럴 인수 등 → #PYTHON! + 한국어 메시지
        this.host.onResult(id, failure("PyGridAnalyzeError", msg(e)), null, null);
        continue;
      }
      let snapshots: Record<string, RangeSnapshot>;
      try {
        snapshots = this.buildSnapshots(refs, block.sheetId);
      } catch (e) {
        this.host.onResult(id, failure("PyGridRefError", msg(e)), null, null);
        continue;
      }
      // 다중 출력: 호스트가 onOutputs를 구현했을 때만 (미구현이면 레거시 단일 출력으로)
      const multi = block.outputs?.length && this.host.onOutputs ? block.outputs : undefined;
      let payload: RunPayload;
      try {
        payload = await this.client.run(
          id,
          block.code,
          snapshots,
          block.outputMode,
          block.includeIndex,
          this.timeoutSec,
          block.output,
          multi?.map((o) => ({
            id: o.id,
            mode: o.mode,
            includeIndex: o.includeIndex,
            selection: o.selection,
          })),
        );
      } catch (e) {
        // 타임아웃 → interrupt → 재부트로 요청이 거부된 경우 등
        const err = failure("WorkerError", msg(e));
        if (multi) this.host.onOutputs!(id, err, []);
        else this.host.onResult(id, err, null, null);
        continue;
      }
      if (multi) {
        const anchors = new Map(multi.map((o) => [o.id, o.anchor]));
        this.host.onOutputs!(
          id,
          payload,
          (payload.ok ? (payload.outputs ?? []) : []).map((item) => {
            const anchor = anchors.get(item.id);
            const cells = item.ok && anchor ? (item.cells ?? null) : null;
            return {
              outputId: item.id,
              item,
              cells: cells ? toCells(cells) : null,
              spill: cells
                ? spillRange(anchor!, cells.length, cells[0]?.length ?? 0)
                : null,
            };
          }),
        );
        continue;
      }
      if (payload.ok && block.outputMode === "values" && payload.cells) {
        this.host.onResult(
          id,
          payload,
          toCells(payload.cells),
          spillRange(block.anchor, payload.cells.length, payload.cells[0]?.length ?? 0),
        );
      } else {
        this.host.onResult(id, payload, null, null);
      }
    }
  }
}

const msg = (e: unknown): string => (e instanceof Error ? e.message : String(e));
