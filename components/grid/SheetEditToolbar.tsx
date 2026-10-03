"use client";

// 시트 편집 툴바 — 붙여넣기·행/열 삽입 삭제·열 고정·정렬·서식.
// 그리드 패널 **안**에 산다: 스프레드시트를 접으면 이 툴바도 함께 사라진다.
// 패널 접기·파일 계열 메뉴는 셸 바(components/shell/ShellBar.tsx)에 있다.

import { useEffect, useState } from "react";
import { toast } from "sonner";
import {
  ClipboardText,
  PushPin,
  SortAscending,
  SortDescending,
  TextB,
} from "@phosphor-icons/react";
import { startPasteFlow } from "@/components/grid/PasteImportDialog";
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
import { Separator } from "@/components/ui/separator";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { ToolButton } from "@/components/shell/ShellBar";
import type { DateOrder } from "@/lib/grid/clipboard/infer";
import { parseClipboard } from "@/lib/grid/clipboard/parse";
import { notifyWorkbookEdit } from "@/lib/grid/calc-host";
import { isRangeBold, useWorkbookStore, type CellEdit } from "@/lib/grid/model";
import { loadSettings, saveSettings } from "@/lib/storage/db";
import { cellKey, parseCellKey, type Cell } from "@/types/workbook";

/** 텍스트로 붙여넣기 폴백 (모바일·클립보드 권한 차단 환경) + 항상 미리보기 설정 */
function PasteOptionsDialog({ open, onClose }: { open: boolean; onClose: () => void }) {
  const [value, setValue] = useState("");
  const [alwaysDialog, setAlwaysDialog] = useState(false);
  const [dateOrder, setDateOrder] = useState<DateOrder>("ymd");

  // 열릴 때 설정 한 번 로드
  useEffect(() => {
    if (open) {
      void loadSettings().then((s) => {
        setAlwaysDialog(s?.pasteAlwaysDialog ?? false);
        setDateOrder(s?.dateOrder ?? "ymd");
      });
    }
  }, [open]);

  const paste = () => {
    const raw = parseClipboard({ text: value });
    if (raw.length > 0) void startPasteFlow(raw);
    setValue("");
    onClose();
  };

  return (
    <Dialog open={open} onOpenChange={(next) => !next && onClose()}>
      <DialogContent>
        <DialogHeader>
          <DialogTitle>텍스트로 붙여넣기</DialogTitle>
          <DialogDescription>
            클립보드 권한이 막힌 환경에서는 아래에 표를 직접 붙여넣으세요 (탭 구분).
          </DialogDescription>
        </DialogHeader>
        <textarea
          value={value}
          onChange={(e) => setValue(e.target.value)}
          rows={8}
          placeholder={"이름\t나이\n철수\t20"}
          className="w-full resize-y rounded border border-input bg-background p-2 font-mono text-xs outline-none focus:border-ring"
          aria-label="붙여넣을 텍스트"
        />
        <div className="flex flex-wrap items-center gap-x-6 gap-y-2 text-sm">
          <label className="flex items-center gap-2">
            <Checkbox
              checked={alwaysDialog}
              onCheckedChange={(v) => {
                const next = v === true;
                setAlwaysDialog(next);
                void saveSettings({ pasteAlwaysDialog: next });
              }}
            />
            붙여넣기 시 항상 미리보기 대화상자 표시
          </label>
          <div className="flex items-center gap-2">
            <span className="text-muted-foreground">날짜 순서</span>
            <Select
              value={dateOrder}
              onValueChange={(v) => {
                const next = v as DateOrder;
                setDateOrder(next);
                void saveSettings({ dateOrder: next });
              }}
            >
              <SelectTrigger className="h-7 w-32 text-xs" aria-label="날짜 순서">
                <SelectValue />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="ymd">연-월-일</SelectItem>
                <SelectItem value="mdy">월-일-연</SelectItem>
              </SelectContent>
            </Select>
          </div>
        </div>
        <DialogFooter>
          <Button variant="ghost" onClick={onClose}>
            닫기
          </Button>
          <Button onClick={paste} disabled={value.trim() === ""}>
            붙여넣기
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}

const isEmpty = (cell?: Cell): boolean =>
  !cell || cell.v === null || cell.v === "";

function compareCells(a: Cell, b: Cell): number {
  const aNum = typeof a.v === "number" || typeof a.v === "boolean";
  const bNum = typeof b.v === "number" || typeof b.v === "boolean";
  if (aNum && bNum) return Number(a.v) - Number(b.v);
  if (aNum) return -1; // 숫자가 문자열보다 앞 (Excel과 동일)
  if (bNum) return 1;
  return String(a.v).localeCompare(String(b.v), "ko");
}

/** 활성 시트 사용 범위를 선택 열 기준 정렬. 첫 행이 전부 문자열이면 헤더로 보고 건너뜀 */
function sortByColumn(direction: 1 | -1): void {
  const { workbook, activeSheetId, selection, setCells } = useWorkbookStore.getState();
  const sheet = workbook.sheets.find((s) => s.id === activeSheetId);
  if (!sheet || !selection) return;
  const keys = Object.keys(sheet.cells);
  if (keys.length === 0) return;
  // spill 잠금 우회 차단: 정렬 대상에 블록 결과(src) 셀이 있으면 거부 (§3.4)
  if (keys.some((k) => sheet.cells[k].src)) {
    toast.error("Python 블록의 결과(spill) 셀이 포함되어 정렬할 수 없습니다. 블록을 삭제하거나 코드를 수정하세요.");
    return;
  }

  let r0 = Infinity, r1 = -1, c0 = Infinity, c1 = -1;
  for (const key of keys) {
    const { r, c } = parseCellKey(key);
    if (r < r0) r0 = r;
    if (r > r1) r1 = r;
    if (c < c0) c0 = c;
    if (c > c1) c1 = c;
  }
  const col = Math.min(Math.max(selection.c0, c0), c1);

  // 헤더 판정: 사용 범위 첫 행의 비어 있지 않은 셀이 모두 문자열
  let hasAny = false;
  let allStrings = true;
  for (let c = c0; c <= c1; c++) {
    const cell = sheet.cells[cellKey(r0, c)];
    if (cell) {
      hasAny = true;
      if (cell.t !== "s") allStrings = false;
    }
  }
  const dataStart = hasAny && allStrings ? r0 + 1 : r0;

  const rows: number[] = [];
  for (let r = dataStart; r <= r1; r++) rows.push(r);
  const valueAt = (r: number) => sheet.cells[cellKey(r, col)];
  rows.sort((ra, rb) => {
    const a = valueAt(ra);
    const b = valueAt(rb);
    const ea = isEmpty(a);
    const eb = isEmpty(b);
    if (ea && eb) return 0;
    if (ea) return 1; // 빈 셀은 방향과 무관하게 마지막
    if (eb) return -1;
    return compareCells(a!, b!) * direction;
  });

  const edits: CellEdit[] = [];
  rows.forEach((srcRow, i) => {
    const dstRow = dataStart + i;
    for (let c = c0; c <= c1; c++) {
      edits.push({ r: dstRow, c, cell: sheet.cells[cellKey(srcRow, c)] ?? null });
    }
  });
  setCells(activeSheetId, edits); // 한 트랜잭션 = 한 undo 단계
}

export default function SheetEditToolbar() {
  const selection = useWorkbookStore((s) => s.selection);
  const frozenCols = useWorkbookStore(
    (s) => s.workbook.sheets.find((sh) => sh.id === s.activeSheetId)?.frozenCols ?? 0,
  );
  const [pasteDialogOpen, setPasteDialogOpen] = useState(false);
  // 부록 J.2: 선택 전체가 굵으면 토글은 해제로 동작
  const selectionBold = useWorkbookStore((s) => {
    if (!s.selection) return false;
    const sheet = s.workbook.sheets.find((sh) => sh.id === s.activeSheetId);
    return sheet ? isRangeBold(sheet, s.selection) : false;
  });
  // 크기 표시는 활성 셀(선택 좌상단) 기준 — 엑셀과 동일
  const selectionFs = useWorkbookStore((s) => {
    const sel = s.selection;
    const sheet = s.workbook.sheets.find((sh) => sh.id === s.activeSheetId);
    const fs = sel && sheet ? sheet.cells[cellKey(sel.r0, sel.c0)]?.st?.fs : undefined;
    return fs ? String(fs) : "default";
  });

  const store = () => useWorkbookStore.getState();
  const rowIndex = selection?.r0 ?? 0;
  const rowSpan = selection ? selection.r1 - selection.r0 + 1 : 1;
  const colIndex = selection?.c0 ?? 0;
  const colSpan = selection ? selection.c1 - selection.c0 + 1 : 1;
  const sid = () => store().activeSheetId;

  /** 부록 O.4: 행/열 구조 변경 뒤 — 앵커·xl() 참조가 바뀐 블록 + 이 시트를 읽는 블록 재실행 */
  const structural = (changedBlocks: string[]) => {
    const sheet = store().workbook.sheets.find((s) => s.id === sid());
    if (!sheet) return;
    notifyWorkbookEdit(
      [{ sheetId: sheet.id, r0: 0, c0: 0, r1: sheet.rowCount - 1, c1: sheet.colCount - 1 }],
      changedBlocks,
    );
  };

  const toggleFreeze = () => {
    const target = colIndex + 1;
    store().setFrozenCols(sid(), frozenCols === target ? 0 : target);
  };

  return (
    <div
      data-testid="sheet-edit-toolbar"
      className="flex h-10 shrink-0 flex-wrap items-center gap-1 overflow-hidden border-b bg-muted/40 px-2"
    >
      <ToolButton label="붙여넣기 옵션 (텍스트로 붙여넣기)" onClick={() => setPasteDialogOpen(true)}>
        <ClipboardText />
      </ToolButton>
      <PasteOptionsDialog open={pasteDialogOpen} onClose={() => setPasteDialogOpen(false)} />
      <Separator orientation="vertical" className="mx-1 h-5" />

      <ToolButton label={`행 삽입 (${rowSpan}개)`} disabled={!selection}
        onClick={() => structural(store().insertRows(sid(), rowIndex, rowSpan))}>
        <span className="text-xs font-semibold">행+</span>
      </ToolButton>
      <ToolButton label={`행 삭제 (${rowSpan}개)`} disabled={!selection}
        onClick={() => structural(store().deleteRows(sid(), rowIndex, rowSpan))}>
        <span className="text-xs font-semibold">행−</span>
      </ToolButton>
      <ToolButton label={`열 삽입 (${colSpan}개)`} disabled={!selection}
        onClick={() => structural(store().insertCols(sid(), colIndex, colSpan))}>
        <span className="text-xs font-semibold">열+</span>
      </ToolButton>
      <ToolButton label={`열 삭제 (${colSpan}개)`} disabled={!selection}
        onClick={() => structural(store().deleteCols(sid(), colIndex, colSpan))}>
        <span className="text-xs font-semibold">열−</span>
      </ToolButton>
      <Separator orientation="vertical" className="mx-1 h-5" />

      <ToolButton
        label={frozenCols > 0 ? `열 고정 해제 (현재 ${frozenCols}열)` : "선택 열까지 고정"}
        disabled={!selection && frozenCols === 0}
        active={frozenCols > 0}
        onClick={toggleFreeze}
      >
        <PushPin weight={frozenCols > 0 ? "fill" : "regular"} />
      </ToolButton>
      <ToolButton label="선택 열 기준 오름차순 정렬" disabled={!selection}
        onClick={() => sortByColumn(1)}>
        <SortAscending />
      </ToolButton>
      <ToolButton label="선택 열 기준 내림차순 정렬" disabled={!selection}
        onClick={() => sortByColumn(-1)}>
        <SortDescending />
      </ToolButton>
      <Separator orientation="vertical" className="mx-1 h-5" />

      {/* 부록 J.2 서식 그룹 — 굵게 토글 + 글자 크기 */}
      <ToolButton
        label={selectionBold ? "굵게 해제" : "굵게"}
        disabled={!selection}
        active={selectionBold}
        onClick={() => selection && store().applyCellStyle(sid(), selection, { b: !selectionBold })}
      >
        <TextB weight="bold" />
      </ToolButton>
      <Select
        value={selectionFs}
        onValueChange={(v) =>
          selection &&
          store().applyCellStyle(sid(), selection, { fs: v === "default" ? null : Number(v) })
        }
        disabled={!selection}
      >
        <SelectTrigger className="h-7 w-20 text-xs" aria-label="글자 크기">
          <SelectValue />
        </SelectTrigger>
        <SelectContent>
          <SelectItem value="default">기본</SelectItem>
          {[12, 14, 16, 20].map((n) => (
            <SelectItem key={n} value={String(n)}>
              {n}px
            </SelectItem>
          ))}
        </SelectContent>
      </Select>
    </div>
  );
}
