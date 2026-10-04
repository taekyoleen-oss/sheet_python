"use client";

// 부록 P: 변수 → 스프레드시트 — 속성 창 변수마다 원하는 형태로 시트에 보인다.
//  · 연결 출력: 변수를 만든 블록에 출력 바인딩을 붙여 블록을 다시 실행할 때마다 갱신 (권장)
//  · 값 복사: 지금 값을 한 번만 복사 (선택 셀 또는 새 시트)
// 공통 옵션: 열 고르기·상위 N행·index 포함

import { useEffect, useMemo, useState } from "react";
import { toast } from "sonner";
import { applyPastedCells } from "@/components/grid/PasteImportDialog";
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
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { formatA1 } from "@/lib/grid/a1";
import { notifyWorkbookEdit } from "@/lib/grid/calc-host";
import { blocksInOrder, useWorkbookStore } from "@/lib/grid/model";
import { definingBlock } from "@/lib/grid/model-output";
import { getRuntimeClient } from "@/lib/runtime/client";
import { toCells } from "@/lib/runtime/converters";
import type { VariableInfo } from "@/lib/runtime/protocol";
import type { IncludeIndex } from "@/types/workbook";

type Mode = "linked" | "copy";
type Where = "selection" | "newSheet";

export default function ShowVariableDialog({
  info,
  onClose,
}: {
  info: VariableInfo | null;
  onClose: () => void;
}) {
  const wb = useWorkbookStore((s) => s.workbook);
  const selection = useWorkbookStore((s) => s.selection);
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);
  const blocks = useMemo(() => blocksInOrder(wb).filter((b) => b.kind !== "markdown"), [wb]);

  const [mode, setMode] = useState<Mode>("linked");
  const [where, setWhere] = useState<Where>("selection");
  const [blockId, setBlockId] = useState("");
  const [cols, setCols] = useState<string[]>([]);
  const [rows, setRows] = useState("");
  const [index, setIndex] = useState<IncludeIndex>("auto");

  useEffect(() => {
    if (!info) return;
    setBlockId(definingBlock(useWorkbookStore.getState().workbook, info.name)?.id ?? "");
    setCols(info.columns ?? []);
    setRows("");
    setIndex("auto");
    setMode(useWorkbookStore.getState().workbook.pyBlocks.length > 0 ? "linked" : "copy");
  }, [info]);

  if (!info) return null;
  const all = info.columns ?? [];
  const rowLimit = Number(rows) > 0 ? Math.floor(Number(rows)) : undefined;
  const columns = all.length > 0 && cols.length > 0 && cols.length < all.length ? cols : undefined;
  const selAddr = selection
    ? formatA1({ r0: selection.r0, c0: selection.c0, r1: selection.r0, c1: selection.c0 })
    : null;

  const apply = async () => {
    const st = useWorkbookStore.getState();
    if (mode === "linked") {
      const block = st.workbook.pyBlocks.find((b) => b.id === blockId);
      if (!block) return;
      const start =
        where === "selection" && selection
          ? { sheetId: activeSheetId, r: selection.r0, c: selection.c0 }
          : undefined;
      if (where === "newSheet") {
        st.addSheet();
      }
      const target = where === "newSheet" ? { sheetId: useWorkbookStore.getState().activeSheetId, r: 0, c: 0 } : start;
      const ids = useWorkbookStore.getState().addOutputs(
        block.id,
        [{ selection: { variable: info.name, columns, rowLimit }, label: info.name, width: (columns?.length ?? all.length) + 1, includeIndex: index }],
        target,
      );
      if (ids.length === 0) return;
      notifyWorkbookEdit([], [block.id]);
      toast.success(`${info.name}을(를) 블록 출력으로 연결했습니다 — 블록을 다시 실행할 때마다 갱신됩니다`);
      onClose();
      return;
    }
    try {
      const payload = await getRuntimeClient().run(
        `__var_export_${info.name}`,
        info.name,
        {},
        "values",
        index,
        undefined,
        { columns, rowLimit },
      );
      if (!payload.ok || !payload.cells) {
        toast.error("값으로 변환할 수 없습니다 (이미지·객체는 연결 출력의 '객체' 모드로 보세요)");
        return;
      }
      if (applyPastedCells(toCells(payload.cells), where === "newSheet" ? "newSheet" : "anchor")) {
        toast.success(`${info.name} 값을 ${where === "newSheet" ? "새 시트" : selAddr}에 복사했습니다`);
        onClose();
      }
    } catch (e) {
      toast.error(`복사 실패: ${(e as Error).message}`);
    }
  };

  return (
    <Dialog open onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="max-h-[85vh] overflow-y-auto sm:max-w-lg">
        <DialogHeader>
          <DialogTitle>
            <span className="font-mono">{info.name}</span> → 스프레드시트
          </DialogTitle>
          <DialogDescription>
            {info.type}
            {info.shape ? ` · ${info.shape[0]}행 × ${info.shape[1]}열` : ""} — 원하는 형태로 시트에 보입니다.
          </DialogDescription>
        </DialogHeader>
        <div className="space-y-4 text-sm">
          <section className="space-y-1.5">
            <Label>방식</Label>
            <Select value={mode} onValueChange={(v) => setMode(v as Mode)}>
              <SelectTrigger className="h-8 w-full" aria-label="표시 방식">
                <SelectValue />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="linked" disabled={blocks.length === 0}>
                  연결 출력 — 블록을 다시 실행하면 갱신
                </SelectItem>
                <SelectItem value="copy">값 복사 — 지금 값을 한 번만</SelectItem>
              </SelectContent>
            </Select>
            {mode === "linked" && (
              <Select value={blockId} onValueChange={setBlockId}>
                <SelectTrigger className="h-8 w-full text-xs" aria-label="연결할 블록">
                  <SelectValue placeholder="블록" />
                </SelectTrigger>
                <SelectContent>
                  {blocks.map((b, i) => (
                    <SelectItem key={b.id} value={b.id} className="text-xs">
                      {i + 1}. {b.title || b.code.split("\n").find((l) => l.trim() && !l.startsWith("#"))?.slice(0, 40) || "(빈 블록)"}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            )}
          </section>

          <section className="space-y-1.5">
            <Label>위치</Label>
            <Select value={where} onValueChange={(v) => setWhere(v as Where)}>
              <SelectTrigger className="h-8 w-full" aria-label="표시 위치">
                <SelectValue />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="selection" disabled={!selAddr}>
                  선택한 셀부터 {selAddr ? `(${selAddr})` : ""}
                </SelectItem>
                <SelectItem value="newSheet">새 시트 A1</SelectItem>
              </SelectContent>
            </Select>
          </section>

          {all.length > 0 && (
            <section className="space-y-1.5">
              <div className="flex items-center justify-between">
                <Label>열 ({cols.length}/{all.length})</Label>
                <button
                  type="button"
                  className="text-xs text-primary"
                  onClick={() => setCols(cols.length === all.length ? [] : all)}
                >
                  {cols.length === all.length ? "모두 해제" : "모두 선택"}
                </button>
              </div>
              <div className="grid max-h-36 grid-cols-2 gap-1 overflow-y-auto rounded border p-2 sm:grid-cols-3">
                {all.map((c) => (
                  <label key={c} className="flex items-center gap-1.5 text-xs">
                    <Checkbox
                      checked={cols.includes(c)}
                      onCheckedChange={(v) =>
                        setCols((cur) => (v === true ? all.filter((x) => cur.includes(x) || x === c) : cur.filter((x) => x !== c)))
                      }
                    />
                    <span className="truncate font-mono">{c}</span>
                  </label>
                ))}
              </div>
            </section>
          )}

          <section className="flex flex-wrap items-center gap-3">
            <label className="flex items-center gap-2 text-xs">
              상위
              <Input
                type="number"
                min={1}
                value={rows}
                placeholder="전체"
                onChange={(e) => setRows(e.target.value)}
                className="h-7 w-20"
                aria-label="상위 N행"
              />
              행
            </label>
            <label className="flex items-center gap-2 text-xs">
              index
              <Select value={index} onValueChange={(v) => setIndex(v as IncludeIndex)}>
                <SelectTrigger className="h-7 w-28 text-xs" aria-label="index 포함">
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="auto">자동</SelectItem>
                  <SelectItem value="always">포함</SelectItem>
                  <SelectItem value="never">제외</SelectItem>
                </SelectContent>
              </Select>
            </label>
          </section>
        </div>
        <DialogFooter>
          <Button variant="ghost" onClick={onClose}>
            취소
          </Button>
          <Button onClick={() => void apply()} disabled={mode === "linked" && !blockId}>
            시트에 보이기
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
