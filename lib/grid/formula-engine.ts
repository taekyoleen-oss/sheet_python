// 수식 셀 레지스트리·의존성 그래프·재계산 — 부록 I.2 → O.4(인덱스·위상 정렬)
// 순수: 전달된 Workbook(immer draft)만 수정하고 바뀐 셀 범위를 돌려준다. 스토어 무관.

import {
  cellKey,
  parseCellKey,
  type CellRange,
  type Workbook,
} from "@/types/workbook";
import { namesMap, parseFormula, type FormulaResult, type ParsedFormula } from "./formula";

/**
 * 워크북의 모든 수식 원문에 fn 적용 (행/열 삽입·삭제, 시트 이름 변경 — 부록 O.2).
 * fn은 (수식, 수식이 놓인 시트 이름)을 받는다. 값 재계산은 호출부가 한다.
 */
export function rewriteAllFormulas(
  wb: Workbook,
  fn: (fx: string, ownSheetName: string) => string,
): void {
  for (const sheet of wb.sheets) {
    for (const cell of Object.values(sheet.cells)) {
      if (!cell.fx) continue;
      const next = fn(cell.fx, sheet.name);
      if (next !== cell.fx) cell.fx = next;
    }
  }
}

export interface SheetRange extends CellRange {
  sheetId: string;
}

// ponytail: (이름 정의 서명 + fx) 키 파스 캐시(무한 보관) — 수식 수 규모라 무해
const parseCache = new Map<string, ParsedFormula>();

interface Entry {
  sheetId: string;
  r: number;
  c: number;
  key: string;
  pf: ParsedFormula;
  /** 시트 이름을 id로 해석한 참조 (모르는 시트는 null — 평가 시 #REF!) */
  refs: { sheetId: string | null; range: CellRange }[];
}

const intersects = (a: CellRange, b: CellRange): boolean =>
  a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;

const cellOf = (e: Entry): CellRange => ({ r0: e.r, c0: e.c, r1: e.r, c1: e.c });

/** 이 면적 이하의 참조는 셀 단위 색인으로 선행 수식을 찾고, 넘으면 수식 목록을 훑는다 */
const INDEX_AREA = 4096;

/**
 * 셀 변경 후 수식 재계산. edited === null이면 전체 재계산(시트 추가/삭제/이름 변경·이름 정의 변경).
 * 영향받는 수식만 의존 순서(Kahn 위상 정렬)로 계산하고, 순환 구성원(Tarjan SCC)은 전부 #CIRC!.
 * 순환의 하류 수식은 구성원이 끝난 뒤 정상 평가된다(#CIRC! 셀 참조는 #CIRC! 전파).
 * v가 실제로 바뀐 셀 범위를 반환한다 (호출부가 notifyWorkbookEdit로 전달).
 */
export function recalcAfter(wb: Workbook, edited: SheetRange[] | null): SheetRange[] {
  // 1) 수식 셀 수집 — ponytail: 전 셀 스캔 O(cells)/쓰기 (clearRange와 동일 규모)
  const byName = new Map(wb.sheets.map((s) => [s.name, s.id]));
  const byId = new Map(wb.sheets.map((s) => [s.id, s]));
  const names = namesMap(wb.names);
  const sig = (wb.names ?? []).map((n) => `${n.name}=${n.ref}`).join(";");
  const entries: Entry[] = [];
  const at = new Map<string, Entry>(); // "sheetId|r:c" → 그 셀의 수식
  for (const sheet of wb.sheets) {
    for (const key of Object.keys(sheet.cells)) {
      const fx = sheet.cells[key].fx;
      if (!fx) continue;
      const { r, c } = parseCellKey(key);
      const ck = `${sig}\u0000${fx}`;
      let pf = parseCache.get(ck);
      if (!pf) {
        pf = parseFormula(fx, names);
        parseCache.set(ck, pf);
      }
      const e: Entry = {
        sheetId: sheet.id,
        r,
        c,
        key,
        pf,
        refs: pf.refs.map((rf) => ({
          sheetId: rf.sheetName === undefined ? sheet.id : (byName.get(rf.sheetName) ?? null),
          range: rf.range,
        })),
      };
      entries.push(e);
      at.set(`${sheet.id}|${key}`, e);
    }
  }
  if (entries.length === 0) return [];

  // 2) 선행 수식(precedents) 색인 — e가 읽는 셀에 놓인 수식들
  const big = entries; // 큰 범위 참조는 전체 수식 목록과 교차 검사
  const precedents = new Map<Entry, Set<Entry>>();
  const dependents = new Map<Entry, Entry[]>();
  for (const e of entries) {
    const pre = new Set<Entry>();
    for (const rf of e.refs) {
      if (rf.sheetId === null) continue;
      const { r0, c0, r1, c1 } = rf.range;
      if ((r1 - r0 + 1) * (c1 - c0 + 1) <= INDEX_AREA) {
        for (let r = r0; r <= r1; r++)
          for (let c = c0; c <= c1; c++) {
            const d = at.get(`${rf.sheetId}|${cellKey(r, c)}`);
            if (d) pre.add(d);
          }
      } else {
        for (const d of big) if (d.sheetId === rf.sheetId && intersects(rf.range, cellOf(d))) pre.add(d);
      }
    }
    precedents.set(e, pre);
    for (const d of pre) {
      const list = dependents.get(d);
      if (list) list.push(e);
      else dependents.set(d, [e]);
    }
  }

  // 3) 시드: 편집 범위 안의 수식 자신 + 편집 범위와 겹치는 참조를 가진 수식 → 의존자 방향 폐포
  const isSeed = (e: Entry): boolean =>
    edited === null ||
    edited.some(
      (ed) =>
        (ed.sheetId === e.sheetId && intersects(ed, cellOf(e))) ||
        e.refs.some((rf) => rf.sheetId === ed.sheetId && intersects(rf.range, ed)),
    );
  const affected = new Set<Entry>();
  const stack = entries.filter(isSeed);
  while (stack.length > 0) {
    const e = stack.pop()!;
    if (affected.has(e)) continue;
    affected.add(e);
    for (const d of dependents.get(e) ?? []) if (!affected.has(d)) stack.push(d);
  }
  if (affected.size === 0) return [];

  // 4) 평가 준비
  const changed: SheetRange[] = [];
  const write = (e: Entry, { v, t }: FormulaResult): void => {
    const sheet = byId.get(e.sheetId);
    if (!sheet) return;
    const old = sheet.cells[e.key];
    if (old && old.v === v && old.t === t) return;
    sheet.cells[e.key] = { ...old, fx: old?.fx ?? "", v, t };
    changed.push({ sheetId: e.sheetId, ...cellOf(e) });
  };
  const getCellFor =
    (e: Entry) => (sheetName: string | undefined, r: number, c: number) => {
      const sid = sheetName === undefined ? e.sheetId : byName.get(sheetName);
      const sheet = sid === undefined ? undefined : byId.get(sid);
      if (!sheet) return "#REF!" as const;
      return sheet.cells[cellKey(r, c)];
    };
  const boundsFor = (e: Entry) => (sheetName: string | undefined) => {
    const sid = sheetName === undefined ? e.sheetId : byName.get(sheetName);
    const sheet = sid === undefined ? undefined : byId.get(sid);
    return sheet ? { rows: sheet.rowCount, cols: sheet.colCount } : undefined;
  };
  const evaluate = (e: Entry) => write(e, e.pf.eval(getCellFor(e), boundsFor(e)));

  // 5) Kahn: affected 안의 선행 수만큼 진입 차수. 막히면 남은 그래프에서 순환(SCC) 제거
  const indeg = new Map<Entry, number>();
  for (const e of affected) {
    let n = 0;
    for (const d of precedents.get(e)!) if (affected.has(d)) n++;
    indeg.set(e, n);
  }
  const done = new Set<Entry>();
  const release = (e: Entry, queue: Entry[]) => {
    done.add(e);
    for (const d of dependents.get(e) ?? []) {
      if (!affected.has(d) || done.has(d)) continue;
      const n = indeg.get(d)! - 1;
      indeg.set(d, n);
      if (n === 0) queue.push(d);
    }
  };
  let queue = [...affected].filter((e) => indeg.get(e) === 0);
  while (done.size < affected.size) {
    for (let h = 0; h < queue.length; h++) {
      const e = queue[h];
      if (done.has(e)) continue;
      evaluate(e);
      release(e, queue);
    }
    if (done.size === affected.size) break;
    const members = cycleMembers([...affected].filter((e) => !done.has(e)), precedents);
    if (members.length === 0) break; // 방어 — 이론상 도달 불가
    queue = [];
    for (const e of members) write(e, { v: "#CIRC!", t: "e" });
    for (const e of members) release(e, queue);
  }
  return changed;
}

/** 남은 노드 중 순환 구성원 (크기 ≥2 SCC 또는 자기 참조) — 반복형 Tarjan */
function cycleMembers(nodes: Entry[], precedents: Map<Entry, Set<Entry>>): Entry[] {
  const inSet = new Set(nodes);
  const index = new Map<Entry, number>();
  const low = new Map<Entry, number>();
  const onStack = new Set<Entry>();
  const st: Entry[] = [];
  const out: Entry[] = [];
  let next = 0;
  for (const root of nodes) {
    if (index.has(root)) continue;
    const work: { v: Entry; it: Iterator<Entry> }[] = [];
    const open = (v: Entry) => {
      index.set(v, next);
      low.set(v, next++);
      st.push(v);
      onStack.add(v);
      work.push({ v, it: precedents.get(v)![Symbol.iterator]() });
    };
    open(root);
    while (work.length > 0) {
      const top = work[work.length - 1];
      const step = top.it.next();
      if (!step.done) {
        const w = step.value;
        if (!inSet.has(w)) continue;
        if (!index.has(w)) open(w);
        else if (onStack.has(w)) low.set(top.v, Math.min(low.get(top.v)!, index.get(w)!));
        continue;
      }
      work.pop();
      if (work.length > 0) {
        const parent = work[work.length - 1].v;
        low.set(parent, Math.min(low.get(parent)!, low.get(top.v)!));
      }
      if (low.get(top.v) === index.get(top.v)) {
        const comp: Entry[] = [];
        let w: Entry;
        do {
          w = st.pop()!;
          onStack.delete(w);
          comp.push(w);
        } while (w !== top.v);
        if (comp.length > 1 || precedents.get(top.v)!.has(top.v)) out.push(...comp);
      }
    }
  }
  return out;
}
