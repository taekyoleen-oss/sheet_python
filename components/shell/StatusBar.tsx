"use client";

// 상태 바 — 선택 범위(A1) · 시트 수 · 계산 모드/블록 자리표시 · 자동 저장 상태

import { useEffect, useState } from "react";
import { toast } from "sonner";
import { formatA1 } from "@/lib/grid/a1";
import { setCalcModeEverywhere } from "@/lib/grid/calc-host";
import { useWorkbookStore } from "@/lib/grid/model";
import { getRuntimeClient } from "@/lib/runtime/client";
import type { SaveStatus } from "@/lib/storage/autosave";
import { saveSettings } from "@/lib/storage/db";

const SAVE_LABEL: Record<SaveStatus, string> = {
  saved: "✓ 저장됨",
  saving: "저장 중…",
  error: "저장 오류 — 파일로 내보내 백업하세요",
};

/** 워커 FS 데이터 파일 칩 (R5) — 클릭: 같은 이름 시트로 이동, 없으면 파일명 복사 */
function DataFileChips() {
  const [files, setFiles] = useState<string[]>([]);
  useEffect(() => {
    const read = () =>
      setFiles(getRuntimeClient().listFiles().filter((f) => !f.startsWith("_")));
    read();
    window.addEventListener("pygrid:data-files", read);
    return () => window.removeEventListener("pygrid:data-files", read);
  }, []);
  const sheets = useWorkbookStore((s) => s.workbook.sheets);
  if (files.length === 0) return null;

  return (
    <span className="flex min-w-0 items-center gap-1 overflow-hidden">
      {files.map((f) => {
        const stem = f.replace(/\.[^.]+$/, "");
        const sheet = sheets.find(
          (s) =>
            s.name === stem ||
            s.name === f ||
            s.name.startsWith(`${stem}-`) ||
            s.name.startsWith(`${stem} (`),
        );
        return (
          <button
            key={f}
            onClick={() => {
              if (sheet) useWorkbookStore.getState().setActiveSheet(sheet.id);
              else {
                void navigator.clipboard?.writeText(f);
                toast(`"${f}" 파일명을 복사했습니다 — 코드에서 pd.read_*로 읽을 수 있습니다`);
              }
            }}
            className="max-w-36 truncate rounded-full border bg-background px-2 py-px font-mono text-[10px] hover:border-primary hover:text-foreground"
            title={sheet ? `${f} — "${sheet.name}" 시트로 이동` : `${f} — 클릭하여 파일명 복사`}
          >
            {f}
          </button>
        );
      })}
    </span>
  );
}

export default function StatusBar({ saveStatus }: { saveStatus: SaveStatus }) {
  const selection = useWorkbookStore((s) => s.selection);
  const sheetCount = useWorkbookStore((s) => s.workbook.sheets.length);
  const calcMode = useWorkbookStore((s) => s.workbook.calcMode);
  const blockCount = useWorkbookStore((s) => s.workbook.pyBlocks.length);
  const dirtyCount = useWorkbookStore((s) => Object.keys(s.dirtyBlocks).length);
  const picking = useWorkbookStore((s) => !!s.anchorPicking);
  const showRefs = useWorkbookStore((s) => s.showRefs);

  return (
    <div className="flex h-7 shrink-0 items-center gap-4 border-t bg-muted/60 px-3 text-xs text-muted-foreground">
      <span className="font-mono">{selection ? formatA1(selection) : "선택 없음"}</span>
      <span>시트 {sheetCount}</span>
      <DataFileChips />
      {picking && (
        <span className="font-medium text-primary">
          결과를 놓을 셀을 클릭하세요 · Esc 취소
        </span>
      )}
      <span className="ml-auto">
        블록 {blockCount}
        {dirtyCount > 0 ? ` (dirty ${dirtyCount})` : ""}
      </span>
      {/* 부록 J.3: 실행 참조 표시 토글 (기본 켬) */}
      <button
        onClick={() => {
          useWorkbookStore.getState().setShowRefs(!showRefs);
          void saveSettings({ showRefs: !showRefs });
        }}
        className={showRefs ? "font-medium text-[#1F6E64]" : "hover:text-foreground"}
        title="블록이 마지막 성공 실행에서 읽은 sheet() 참조 범위를 그리드에 표시"
        aria-pressed={showRefs}
      >
        참조 표시 {showRefs ? "켬" : "끔"}
      </button>
      <button
        onClick={() => setCalcModeEverywhere(calcMode === "auto" ? "manual" : "auto")}
        className={calcMode === "manual" ? "font-medium text-warning-text" : "hover:text-foreground"}
        title="클릭하여 자동/수동 전환"
      >
        계산 {calcMode === "auto" ? "자동" : "수동"}
      </button>
      <span className={saveStatus === "error" ? "font-medium text-destructive" : ""}>
        {SAVE_LABEL[saveStatus]}
      </span>
    </div>
  );
}
