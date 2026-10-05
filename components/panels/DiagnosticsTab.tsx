"use client";

// 진단 탭 — 블록의 stdout/stderr + 출력별 상태·한국어 요약·traceback, 셀·블록 이동 (§4.4, 부록 D.1)

import { Badge } from "@/components/ui/badge";
import { formatA1 } from "@/lib/grid/a1";
import { useWorkbookStore } from "@/lib/grid/model";
import { onSheet, outputsOf } from "@/lib/grid/outputs";
import { blocksInOrder } from "@/lib/grid/run-block";
import type { OutputBinding, PyBlock } from "@/types/workbook";

const STATUS_LABEL: Record<string, { label: string; cls: string }> = {
  ok: { label: "성공", cls: "bg-primary/10 text-primary" },
  error: { label: "오류", cls: "bg-destructive/10 text-destructive" },
  spill: { label: "#SPILL!", cls: "bg-destructive/10 text-destructive" },
};

const anchorLabel = (sheet: string, anchor: { r: number; c: number }) =>
  `${sheet}!${formatA1({ r0: anchor.r, c0: anchor.c, r1: anchor.r, c1: anchor.c })}`;

/** 출력 하나의 진단 행 — 다중 출력이면 출력마다 상태·오류가 따로 표시된다 */
function OutputRow({
  block,
  output,
  index,
  sheetName,
}: {
  block: PyBlock;
  output: OutputBinding;
  index: number;
  sheetName: (id: string) => string;
}) {
  const sheetId = output.sheetId ?? block.sheetId;
  const last = output.last;
  const status = last ? STATUS_LABEL[last.status] : undefined;
  const name =
    output.label ?? output.selection?.variable ?? `출력 ${index + 1}`;

  const goToCell = () => {
    const st = useWorkbookStore.getState();
    st.setActiveSheet(sheetId);
    st.setSelection({
      r0: output.anchor.r,
      c0: output.anchor.c,
      r1: output.anchor.r,
      c1: output.anchor.c,
    });
  };

  return (
    <div className="mt-1 border-l-2 border-border pl-2">
      <div className="flex items-center gap-2">
        <span className="text-muted-foreground">{name}</span>
        {/* 시트에 놓인 출력만 셀 주소가 있다 — Python 결과로만 보는 블록·위치 미정 출력은 주소 없음 */}
        {onSheet(block, output) && (
          <button
            onClick={goToCell}
            className="font-mono text-primary hover:underline"
            title="셀로 이동"
          >
            {anchorLabel(sheetName(sheetId), output.anchor)}
          </button>
        )}
        {status && <Badge className={status.cls}>{status.label}</Badge>}
        {!last && <span className="text-muted-foreground">실행 전</span>}
      </div>
      {last?.summaryKo && (
        <p className={last.status === "ok" ? "" : "text-destructive"}>
          {last.summaryKo}
        </p>
      )}
      {last?.traceback && (
        <details>
          <summary className="cursor-pointer text-muted-foreground">
            traceback
          </summary>
          <pre className="mt-1 overflow-x-auto rounded bg-code-bg p-2 font-mono text-[11px] leading-4">
            {last.traceback}
          </pre>
        </details>
      )}
    </div>
  );
}

function BlockRow({
  block,
  sheetName,
}: {
  block: PyBlock;
  sheetName: (id: string) => string;
}) {
  const outputs = outputsOf(block);
  const placed = outputs.find((o) => onSheet(block, o)); // 셀에 놓인 첫 출력 (없으면 주소 없음)
  // stdout/stderr·소요 시간은 실행 단위(코드 1회 실행) — 출력마다 반복하지 않고 블록에 한 번만 보인다
  const run = outputs.find((o) => o.last)?.last;
  const streams = [
    run?.stdout && `── stdout ──`,
    run?.stdout,
    run?.stderr && `── stderr ──`,
    run?.stderr,
  ]
    .filter(Boolean)
    .join("\n");

  return (
    <div className="border-b px-3 py-2 text-xs">
      <div className="flex items-center gap-2">
        {placed ? (
          <span className="font-mono">
            {anchorLabel(sheetName(placed.sheetId ?? block.sheetId), placed.anchor)}
          </span>
        ) : (
          !block.title && <span className="text-muted-foreground">Python 블록</span>
        )}
        {block.title && <span className="truncate">{block.title}</span>}
        <span className="text-muted-foreground">
          {run ? `${run.durationMs}ms` : "실행 전"}
          {outputs.length > 1 && ` · 출력 ${outputs.length}개`}
        </span>
        <button
          onClick={() => useWorkbookStore.getState().setFocusBlock(block.id)}
          className="ml-auto text-primary hover:underline"
        >
          블록으로 이동
        </button>
      </div>
      {streams && (
        <details className="mt-1">
          <summary className="cursor-pointer text-muted-foreground">
            stdout / stderr
          </summary>
          <pre className="mt-1 overflow-x-auto rounded bg-code-bg p-2 font-mono text-[11px] leading-4">
            {streams}
          </pre>
        </details>
      )}
      {outputs.map((o, i) => (
        <OutputRow
          key={o.id}
          block={block}
          output={o}
          index={i}
          sheetName={sheetName}
        />
      ))}
    </div>
  );
}

export default function DiagnosticsTab() {
  const workbook = useWorkbookStore((s) => s.workbook);
  // 마크다운 블록은 실행되지 않으므로 진단 대상이 아니다
  const blocks = blocksInOrder(workbook).filter((b) => b.kind !== "markdown");
  const sheetName = (id: string) =>
    workbook.sheets.find((s) => s.id === id)?.name ?? "?";

  if (blocks.length === 0) {
    return (
      <p
        data-testid="diagnostics-tab"
        className="px-3 py-6 text-center text-xs text-muted-foreground"
      >
        Python 블록이 없습니다
      </p>
    );
  }
  return (
    <div data-testid="diagnostics-tab">
      {blocks.map((b) => (
        <BlockRow key={b.id} block={b} sheetName={sheetName} />
      ))}
    </div>
  );
}
