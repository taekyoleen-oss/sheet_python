"use client";

// 붙여넣기 미리보기 다이얼로그 + 붙여넣기 적용 플로우 — 설계서 §4.5.2·4.5.3
// startPasteFlow: 5행 이하(설정으로 항상 표시 가능)는 즉시 반영, 그 외 다이얼로그.

import { useEffect, useMemo, useState } from "react";
import { create } from "zustand";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { Label } from "@/components/ui/label";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { colToLetter } from "@/lib/grid/a1";
import { notifyWorkbookEdit } from "@/lib/grid/calc-host";
import {
  classifyCell,
  inferCells,
  type DateOrder,
  type InferResult,
} from "@/lib/grid/clipboard/infer";
import {
  buildFillEdits,
  buildPasteEdits,
  forgetCopy,
  snapshotRange,
  type CopiedBlock,
} from "@/lib/grid/clipboard/internal";
import { formatCellDisplay } from "@/lib/grid/format";
import { useWorkbookStore, type CellEdit } from "@/lib/grid/model";
import { loadSettings } from "@/lib/storage/db";
import { cellKey, type Cell, type CellRange } from "@/types/workbook";

/** 열 유형 재정의: 자동/숫자/문자/날짜/불리언 */
type ColumnOverride = "auto" | "n" | "s" | "d" | "b";
type PasteTarget = "anchor" | "newSheet";

const usePasteStore = create<{ raw: string[][] | null; dateOrder: DateOrder }>()(
  () => ({ raw: null, dateOrder: "ymd" }),
);

const SPILL_OVERLAP =
  "붙여넣기 범위가 Python 블록의 spill 셀과 겹칩니다. 다른 위치를 선택하세요.";

/** 편집 목록 반영 (한 setCells = 한 undo 단계) + 통지. spill 셀과 겹치면 중단 */
function applyEdits(edits: CellEdit[]): boolean {
  const state = useWorkbookStore.getState();
  const sheetId = state.activeSheetId;
  const sheet = state.workbook.sheets.find((s) => s.id === sheetId);
  if (!sheet || edits.length === 0) return false;
  if (edits.some((e) => sheet.cells[cellKey(e.r, e.c)]?.src)) {
    toast.error(SPILL_OVERLAP);
    return false;
  }
  state.setCells(sheetId, edits);
  const rs = edits.map((e) => e.r);
  const cs = edits.map((e) => e.c);
  notifyWorkbookEdit([
    { sheetId, r0: Math.min(...rs), c0: Math.min(...cs), r1: Math.max(...rs), c1: Math.max(...cs) },
  ]);
  return true;
}

/** 부록 O.3: 앱 안에서 복사한 범위 붙여넣기 — 수식 상대 참조 이동, 선택이 정수배면 타일링 */
export const applyInternalPaste = (src: CopiedBlock): boolean => {
  const state = useWorkbookStore.getState();
  const sel = state.selection ?? { r0: 0, c0: 0, r1: 0, c1: 0 };
  if (!src.cut) return applyEdits(buildPasteEdits(src, sel));
  // 부록 O.4 잘라내기 → 이동 (1회용)
  forgetCopy();
  const srcRange = {
    r0: src.r0,
    c0: src.c0,
    r1: src.r0 + src.cells.length - 1,
    c1: src.c0 + (src.cells[0]?.length ?? 1) - 1,
  };
  if (src.sheetId === state.activeSheetId) {
    const problem = state.moveRange(state.activeSheetId, srcRange, { r: sel.r0, c: sel.c0 });
    if (problem) {
      toast.error(problem);
      return false;
    }
    const dest = {
      r0: sel.r0,
      c0: sel.c0,
      r1: sel.r0 + srcRange.r1 - srcRange.r0,
      c1: sel.c0 + srcRange.c1 - srcRange.c0,
    };
    notifyWorkbookEdit([
      { sheetId: state.activeSheetId, ...srcRange },
      { sheetId: state.activeSheetId, ...dest },
    ]);
    return true;
  }
  // ponytail: 다른 시트로의 잘라내기는 "수식째 복사 + 원본 지우기" — 다른 수식의 참조는 따라가지 않는다
  if (!applyEdits(buildPasteEdits({ ...src, cut: false }, { ...sel, r1: sel.r0, c1: sel.c0 }))) return false;
  if (src.sheetId) {
    state.clearRange(src.sheetId, srcRange);
    notifyWorkbookEdit([{ sheetId: src.sheetId, ...srcRange }]);
  }
  return true;
};

/** 부록 O.4 채우기 핸들 — 원본 범위를 대상까지 연속 패턴으로 채운다 */
export function applyFill(src: CellRange, dest: CellRange): boolean {
  const state = useWorkbookStore.getState();
  const sheet = state.workbook.sheets.find((s) => s.id === state.activeSheetId);
  if (!sheet) return false;
  return applyEdits(buildFillEdits(sheet, src, dest));
}

/**
 * 부록 O.3: 채우기 — Ctrl+D(아래로)·Ctrl+R(오른쪽으로). 선택 첫 행/열을 나머지에 복사.
 * 한 행/열만 선택했으면 바로 위/왼쪽 셀을 가져온다 (엑셀 동일).
 */
export function fillSelection(dir: "down" | "right"): boolean {
  const state = useWorkbookStore.getState();
  const sheet = state.workbook.sheets.find((s) => s.id === state.activeSheetId);
  const sel = state.selection;
  if (!sheet || !sel) return false;
  const r0 = Math.min(sel.r0, sel.r1);
  const r1 = Math.max(sel.r0, sel.r1);
  const c0 = Math.min(sel.c0, sel.c1);
  const c1 = Math.max(sel.c0, sel.c1);
  if (dir === "down") {
    const from = r1 > r0 ? r0 : r0 - 1;
    if (from < 0) return false;
    const src = snapshotRange(sheet, { r0: from, c0, r1: from, c1 });
    return applyEdits(buildPasteEdits(src, { r0: from + 1, c0, r1, c1 }));
  }
  const from = c1 > c0 ? c0 : c0 - 1;
  if (from < 0) return false;
  const src = snapshotRange(sheet, { r0, c0: from, r1, c1: from });
  return applyEdits(buildPasteEdits(src, { r0, c0: from + 1, r1, c1 }));
}

/** 붙여넣은 셀을 스토어에 반영 (한 setCells = 한 undo 단계). spill 셀과 겹치면 중단 */
export function applyPastedCells(
  cells: (Cell | null)[][],
  target: PasteTarget,
): boolean {
  const state = useWorkbookStore.getState();
  if (target === "newSheet") {
    const edits: CellEdit[] = [];
    cells.forEach((row, i) =>
      row.forEach((cell, j) => edits.push({ r: i, c: j, cell })),
    );
    state.addSheetWithCells(edits); // 시트 생성 + 채우기 = 한 undo 단계
    return true;
  }
  const sheetId = state.activeSheetId;
  const base = { r: state.selection?.r0 ?? 0, c: state.selection?.c0 ?? 0 };
  const sheet = state.workbook.sheets.find((s) => s.id === sheetId);
  if (!sheet) return false;
  for (let i = 0; i < cells.length; i++) {
    for (let j = 0; j < cells[i].length; j++) {
      if (sheet.cells[cellKey(base.r + i, base.c + j)]?.src) {
        toast.error(SPILL_OVERLAP);
        return false;
      }
    }
  }
  const edits: CellEdit[] = [];
  cells.forEach((row, i) =>
    row.forEach((cell, j) => edits.push({ r: base.r + i, c: base.c + j, cell })),
  );
  state.setCells(sheetId, edits);
  notifyWorkbookEdit([
    {
      sheetId,
      r0: base.r,
      c0: base.c,
      r1: base.r + cells.length - 1,
      c1: base.c + (cells[0]?.length ?? 1) - 1,
    },
  ]);
  return true;
}

/** 파싱된 raw 표를 받아 즉시 반영하거나 다이얼로그를 연다 */
export async function startPasteFlow(raw: string[][]): Promise<void> {
  if (raw.length === 0 || raw[0].length === 0) return;
  let always = false;
  let dateOrder: DateOrder = "ymd";
  try {
    const settings = await loadSettings();
    always = settings?.pasteAlwaysDialog ?? false;
    dateOrder = settings?.dateOrder ?? "ymd";
  } catch {
    /* 설정 로드 실패 시 기본 동작 */
  }
  if (!always && raw.length <= 5) {
    // 5행 이하·단일 셀은 다이얼로그 생략
    applyPastedCells(inferCells(raw, { dateOrder }).cells, "anchor");
    return;
  }
  usePasteStore.setState({ raw, dateOrder });
}

function computeCells(
  raw: string[][],
  headerRow: boolean,
  overrides: Record<number, ColumnOverride>,
  dateOrder: DateOrder,
): InferResult {
  const result = inferCells(raw, { dateOrder, forceHeader: headerRow });
  const start = headerRow ? 1 : 0;
  for (const [colStr, override] of Object.entries(overrides)) {
    if (override === "auto") continue;
    const c = Number(colStr);
    for (let r = start; r < raw.length; r++) {
      const rawVal = (raw[r][c] ?? "").trim();
      if (rawVal === "") {
        result.cells[r][c] = null;
        continue;
      }
      if (override === "s") {
        result.cells[r][c] = { v: rawVal, t: "s" };
        continue;
      }
      const cl = classifyCell(rawVal, dateOrder);
      result.cells[r][c] =
        cl.cls === override ? cl.cell : { v: rawVal, t: "s" };
    }
  }
  return result;
}

const OVERRIDE_LABELS: Record<ColumnOverride, string> = {
  auto: "자동",
  n: "숫자",
  s: "문자",
  d: "날짜",
  b: "불리언",
};

const PREVIEW_ROWS = 20;

export default function PasteImportDialog() {
  const raw = usePasteStore((s) => s.raw);
  const dateOrder = usePasteStore((s) => s.dateOrder);
  const [headerRow, setHeaderRow] = useState(false);
  const [overrides, setOverrides] = useState<Record<number, ColumnOverride>>({});
  const [target, setTarget] = useState<PasteTarget>("anchor");

  // 새 붙여넣기마다 상태 초기화 (헤더는 추론값으로 미리 체크)
  useEffect(() => {
    if (raw) {
      setHeaderRow(inferCells(raw, { dateOrder }).headerRow);
      setOverrides({});
      setTarget("anchor");
    }
  }, [raw, dateOrder]);

  const computed = useMemo(
    () => (raw ? computeCells(raw, headerRow, overrides, dateOrder) : null),
    [raw, headerRow, overrides, dateOrder],
  );

  if (!raw || !computed) return null;
  const width = computed.cells[0]?.length ?? 0;
  const close = () => usePasteStore.setState({ raw: null });

  const apply = () => {
    if (applyPastedCells(computed.cells, target)) close();
  };

  return (
    <Dialog open onOpenChange={(open) => !open && close()}>
      <DialogContent className="max-w-[880px]">
        <DialogHeader>
          <DialogTitle>붙여넣기 미리보기</DialogTitle>
          <DialogDescription>
            {raw.length.toLocaleString()}행 × {width}열 — 열별 추론 유형을 확인하고
            수정하세요. (상위 {PREVIEW_ROWS}행 표시)
          </DialogDescription>
        </DialogHeader>

        <div className="max-h-80 overflow-auto rounded border">
          <table className="w-full border-collapse text-xs">
            <thead className="sticky top-0 bg-muted">
              <tr>
                {Array.from({ length: width }, (_, c) => (
                  <th key={c} className="border-b border-r p-1 text-left font-normal">
                    <div className="mb-1 font-mono text-[10px] text-muted-foreground">
                      {colToLetter(c)}
                    </div>
                    <Select
                      value={overrides[c] ?? "auto"}
                      onValueChange={(v) =>
                        setOverrides((prev) => ({ ...prev, [c]: v as ColumnOverride }))
                      }
                    >
                      <SelectTrigger className="h-6 w-full text-xs" aria-label={`${colToLetter(c)}열 유형`}>
                        <SelectValue />
                      </SelectTrigger>
                      <SelectContent>
                        {Object.entries(OVERRIDE_LABELS).map(([value, label]) => (
                          <SelectItem key={value} value={value}>
                            {label}
                          </SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {computed.cells.slice(0, PREVIEW_ROWS).map((row, r) => (
                <tr key={r} className={headerRow && r === 0 ? "bg-accent/60 font-medium" : ""}>
                  {row.map((cell, c) => (
                    <td
                      key={c}
                      className={`border-b border-r p-1 ${cell?.t === "n" ? "text-right tabular-nums" : ""}`}
                    >
                      {cell ? formatCellDisplay(cell) : ""}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        <div className="flex flex-wrap items-center gap-6 text-sm">
          <label className="flex items-center gap-2">
            <Checkbox
              checked={headerRow}
              onCheckedChange={(v) => setHeaderRow(v === true)}
            />
            첫 행은 헤더
          </label>
          <div className="flex items-center gap-2">
            <Label htmlFor="paste-target">붙여넣을 위치</Label>
            <Select value={target} onValueChange={(v) => setTarget(v as PasteTarget)}>
              <SelectTrigger id="paste-target" className="h-7 w-36 text-xs">
                <SelectValue />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="anchor">현재 셀 위치</SelectItem>
                <SelectItem value="newSheet">새 시트</SelectItem>
              </SelectContent>
            </Select>
          </div>
        </div>

        <DialogFooter>
          <Button variant="ghost" onClick={close}>
            취소
          </Button>
          <Button onClick={apply}>적용</Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
