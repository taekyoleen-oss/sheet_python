"use client";

// 블록 카드 (Colab 스타일 셀) — 왼쪽 원형 ▶ + hover 시 떠오르는 우상단 툴바(위·아래·편집·삭제·더보기).
// 헤더에는 접기·앵커·제목·상태만 남기고 보조 조작은 ⋮ 메뉴로 모은다.
// kind==='markdown'이면 실행 UI 없이 마크다운 편집/미리보기만 (셀에 아무것도 쓰지 않는다).

import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import {
  ArrowDown,
  ArrowUp,
  CaretDown,
  CaretRight,
  ChartLineUp,
  DotsThreeVertical,
  Eye,
  NotePencil,
  Play,
  Plus,
  TrashSimple,
  X,
} from "@phosphor-icons/react";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import {
  DropdownMenu,
  DropdownMenuCheckboxItem,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuRadioGroup,
  DropdownMenuRadioItem,
  DropdownMenuSeparator,
  DropdownMenuShortcut,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { Input } from "@/components/ui/input";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { AiAssist } from "@/components/python/AiAssist";
import CodeEditor from "@/components/python/CodeEditor";
import ModelResultDialog from "@/components/python/ModelResultDialog";
import { PreviewImage, PreviewTable } from "@/components/panels/OutputPreviewTab";
import type { PreviewPayload } from "@/lib/runtime/protocol";
import { formatA1 } from "@/lib/grid/a1";
import { notifyCodeEdit, notifyWorkbookEdit } from "@/lib/grid/calc-host";
import { codeTitle } from "@/lib/grid/code-sections";
import { toast } from "sonner";
import { applyMdAction, renderMarkdown, type MdAction } from "@/lib/grid/markdown";
import { useWorkbookStore } from "@/lib/grid/model";
import { onSheet, outputsOf } from "@/lib/grid/outputs";
import { moveBlock, runBlock } from "@/lib/grid/run-block";
import { getRuntimeClient } from "@/lib/runtime/client";
import { cn } from "@/lib/utils";
import {
  cellKey,
  type OutputBinding,
  type OutputMode,
  type OutputSelection,
  type PyBlock,
} from "@/types/workbook";

/** Radix Select는 빈 값을 못 쓴다 — '마지막 표현식' 자리표시 값 */
const LAST_EXPR = "__last__";

const store = () => useWorkbookStore.getState();

function statusBadge(block: PyBlock, running: boolean) {
  if (running) {
    return <Badge className="bg-warning/15 text-warning-text">실행 중</Badge>;
  }
  switch (block.last?.status) {
    case "ok":
      return <Badge className="bg-primary/10 text-primary">성공</Badge>;
    case "error":
      return <Badge variant="destructive">오류</Badge>;
    case "spill":
      return <Badge variant="destructive">#SPILL!</Badge>;
    default:
      return <Badge variant="secondary">준비</Badge>;
  }
}

/** 출력 선택 변경 → dirty 표시 + (자동 모드) 재실행. 필터링은 런타임이 한다 */
const applyOutput = (blockId: string, outputId: string, patch: OutputSelection) => {
  store().setOutputSelection(blockId, outputId, patch);
  notifyWorkbookEdit([], [blockId]);
};

/** 출력 변수 — 런타임 전역 변수 목록(inspect) + '마지막 표현식' */
function VariableSelect({
  block,
  output,
  index,
}: {
  block: PyBlock;
  output: OutputBinding;
  index: number;
}) {
  const [vars, setVars] = useState<string[]>([]);
  const loaded = useRef(false);

  const refresh = useCallback(async () => {
    const client = getRuntimeClient();
    if (client.getStatus() !== "ready") return;
    try {
      setVars((await client.inspect()).map((v) => v.name));
      loaded.current = true;
    } catch {
      /* 재부트 중 등 — 다음 열기에서 회복 */
    }
  }, []);

  // 실행이 끝나면 갱신 (한 번이라도 목록을 연 카드만 — 유휴 inspect 폭주 방지)
  const ranAt = output.last?.ranAt;
  useEffect(() => {
    if (loaded.current) void refresh();
  }, [ranAt, refresh]);

  const current = output.selection?.variable;
  const names = current && !vars.includes(current) ? [current, ...vars] : vars;

  return (
    <Select
      value={current ?? LAST_EXPR}
      onValueChange={(v) =>
        applyOutput(block.id, output.id, { variable: v === LAST_EXPR ? undefined : v })
      }
      onOpenChange={(open) => {
        if (open) void refresh();
      }}
    >
      <SelectTrigger className="h-6 w-32 text-xs" aria-label={`출력 ${index + 1} 변수`}>
        <SelectValue />
      </SelectTrigger>
      <SelectContent>
        <SelectItem value={LAST_EXPR}>마지막 표현식</SelectItem>
        {names.map((n) => (
          <SelectItem key={n} value={n} className="font-mono">
            {n}
          </SelectItem>
        ))}
      </SelectContent>
    </Select>
  );
}

/**
 * 결과 표의 열 이름. 객체 모드는 preview.columns, 값 모드는 preview가 없으므로
 * spill 헤더 행(전부 문자열)에서 읽는다.
 * ponytail: 열을 거른 뒤에는 남은 열만 보인다 — '전체'로 되돌리면 목록도 복구된다.
 */
function useTableColumns(block: PyBlock, output: OutputBinding): string[] {
  const sheetId = output.sheetId ?? block.sheetId;
  const sheet = useWorkbookStore((s) => s.workbook.sheets.find((sh) => sh.id === sheetId));
  const preview = output.last?.preview as
    | { kind?: string; columns?: string[] }
    | undefined;
  const selected = output.selection?.columns ?? [];
  let names: string[] = [];
  if (preview?.kind === "table" && preview.columns) {
    names = preview.columns;
  } else {
    const rg = output.last?.spillRange;
    if (sheet && rg && output.last?.kind === "table" && rg.r1 > rg.r0) {
      for (let c = rg.c0; c <= rg.c1; c++) {
        const cell = sheet.cells[cellKey(rg.r0, c)];
        if (!cell || cell.t !== "s") {
          names = []; // 헤더 행이 아니다 (Series·목록 등)
          break;
        }
        names.push(String(cell.v ?? ""));
      }
      if (names[0] === "") names.shift(); // index 라벨 열
    }
  }
  return [...names, ...selected.filter((c) => !names.includes(c))];
}

/** 열 선택 — 마지막 결과가 표일 때만 표시 */
function ColumnsMenu({
  block,
  output,
  index,
}: {
  block: PyBlock;
  output: OutputBinding;
  index: number;
}) {
  const all = useTableColumns(block, output);
  if (all.length === 0) return null;

  const selected = output.selection?.columns;
  const current = selected ?? all;
  const commit = (next: string[]) => {
    const full = next.length === 0 || next.length === all.length;
    applyOutput(block.id, output.id, { columns: full ? undefined : next });
  };

  return (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <Button
          variant="outline"
          size="sm"
          className="h-6 px-2 text-xs"
          aria-label={`출력 ${index + 1} 열 선택`}
        >
          열 {selected ? `${selected.length}/${all.length}` : "전체"}
        </Button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="start" className="max-h-64 overflow-y-auto">
        <DropdownMenuCheckboxItem
          checked={!selected}
          onSelect={(e) => e.preventDefault()}
          onCheckedChange={() => commit(all)}
        >
          전체
        </DropdownMenuCheckboxItem>
        <DropdownMenuSeparator />
        {all.map((c) => (
          <DropdownMenuCheckboxItem
            key={c}
            checked={current.includes(c)}
            onSelect={(e) => e.preventDefault()}
            onCheckedChange={(v) =>
              commit(v ? [...all.filter((x) => current.includes(x) || x === c)] : current.filter((x) => x !== c))
            }
            className="font-mono text-xs"
          >
            {c}
          </DropdownMenuCheckboxItem>
        ))}
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

const OUT_STATUS: Record<string, { label: string; cls: string }> = {
  ok: { label: "성공", cls: "bg-primary/10 text-primary" },
  error: { label: "오류", cls: "bg-destructive/10 text-destructive" },
  spill: { label: "#SPILL!", cls: "bg-destructive/10 text-destructive" },
};

/** 출력 한 줄: [셀 주소][상태][변수][열][상위 N행][값/객체][삭제] */
function OutputRow({
  block,
  output,
  index,
  last,
}: {
  block: PyBlock;
  output: OutputBinding;
  index: number;
  /** 마지막 출력 — 지우면 블록이 시트에서 빠진다(Python 결과로만 보기) */
  last: boolean;
}) {
  const sheetId = output.sheetId ?? block.sheetId;
  const sheetName = useWorkbookStore(
    (s) => s.workbook.sheets.find((sh) => sh.id === sheetId)?.name ?? "?",
  );
  const picking = useWorkbookStore(
    (s) => s.anchorPicking?.blockId === block.id && s.anchorPicking.outputId === output.id,
  );
  const addr = formatA1({
    r0: output.anchor.r,
    c0: output.anchor.c,
    r1: output.anchor.r,
    c1: output.anchor.c,
  });
  const status = output.last ? OUT_STATUS[output.last.status] : undefined;

  return (
    <div
      className="flex flex-wrap items-center gap-1 border-b bg-muted/20 px-2 py-1"
      data-output-id={output.id}
    >
      <button
        onClick={() =>
          store().setAnchorPicking(picking ? null : { blockId: block.id, outputId: output.id })
        }
        aria-label={`출력 ${index + 1} 위치`}
        aria-pressed={picking}
        title="클릭 후 결과를 놓을 셀을 고르세요"
        className={cn(
          "rounded border px-1.5 py-0.5 font-mono text-xs hover:bg-accent",
          picking
            ? "border-primary text-primary"
            : output.unplaced
              ? "border-dashed border-muted-foreground/50 text-muted-foreground"
              : "border-transparent text-foreground/80",
        )}
      >
        {output.unplaced ? "셀 선택" : sheetId === block.sheetId ? addr : `${sheetName}!${addr}`}
      </button>
      {status && (
        <Badge className={status.cls} title={output.last?.summaryKo}>
          {status.label}
        </Badge>
      )}
      <VariableSelect block={block} output={output} index={index} />
      <ColumnsMenu block={block} output={output} index={index} />
      <Input
        type="number"
        min={1}
        inputMode="numeric"
        value={output.selection?.rowLimit ?? ""}
        placeholder="상위 N행"
        aria-label={`출력 ${index + 1} 상위 N행`}
        onChange={(e) => {
          const n = Number(e.target.value);
          applyOutput(block.id, output.id, {
            rowLimit: Number.isFinite(n) && n > 0 ? Math.floor(n) : undefined,
          });
        }}
        className="h-6 w-24 px-1.5"
      />
      <Select
        value={output.mode}
        onValueChange={(v) => {
          store().setOutputMode(block.id, output.id, v as OutputMode);
          notifyWorkbookEdit([], [block.id]);
        }}
      >
        <SelectTrigger className="h-6 w-20 text-xs" aria-label={`출력 ${index + 1} 모드`}>
          <SelectValue />
        </SelectTrigger>
        <SelectContent>
          <SelectItem value="values">값</SelectItem>
          <SelectItem value="object">객체</SelectItem>
        </SelectContent>
      </Select>
      <Button
        variant="ghost"
        size="icon-xs"
        onClick={() => store().removeOutput(block.id, output.id)}
        aria-label={`출력 ${index + 1} 삭제`}
        title={last ? "시트에서 빼기 — Python 결과로만 봅니다" : "이 출력 삭제"}
      >
        <X />
      </Button>
      {output.last?.status !== "ok" && output.last?.summaryKo && (
        <span className="w-full truncate text-xs text-destructive" title={output.last.summaryKo}>
          {output.last.summaryKo}
        </span>
      )}
    </div>
  );
}

/** 출력 목록 — 한 블록의 결과를 여러 셀에 나눠 놓는다 (부록 D.1) */
function OutputList({ block }: { block: PyBlock }) {
  // 꺼진 출력(속성 창으로 처음 추가할 때의 블록 앵커)은 숨기고, 위치를 아직 안 고른 출력은 보인다
  const outputs = outputsOf(block).filter((o) => block.sheetOut === true && !o.off);
  const [modelOpen, setModelOpen] = useState(false);
  return (
    // 카드 내부 밴드 구분: 출력 설정은 옅은 muted 배경 + 상하 경계 (설명·코드와 시각 분리)
    <div data-testid="output-list" className="border-y bg-muted/25">
      {outputs.map((o, i) => (
        <OutputRow
          key={o.id}
          block={block}
          output={o}
          index={i}
          last={outputs.length === 1}
        />
      ))}
      <div className="border-b bg-muted/20 px-2 py-1">
        <Button
          variant="ghost"
          size="sm"
          className="h-6 px-2 text-xs text-primary"
          onClick={() => setModelOpen(true)}
          title="적합한 모델의 계수·결정계수·예측값을 골라 셀에 놓습니다 (부록 O.5)"
        >
          <ChartLineUp className="size-3" /> 모델 결과 → 시트
        </Button>
        {modelOpen && (
          <ModelResultDialog block={block} open={modelOpen} onOpenChange={setModelOpen} />
        )}
      </div>
    </div>
  );
}

/** 노트북식 실행 결과 — 코드 바로 아래에 stdout/stderr·오류·마지막 값(표·이미지·repr). 숨기기 가능 */
function CellResult({ block }: { block: PyBlock }) {
  const [hidden, setHidden] = useState(false);
  const last = block.last;
  if (!last) return null;
  const preview = last.preview as PreviewPayload | undefined;
  const isError = last.status === "error";
  return (
    <div data-testid="cell-result" className="border-t bg-card">
      <button
        onClick={() => setHidden((v) => !v)}
        aria-expanded={!hidden}
        className="flex w-full items-center gap-1 px-2 py-0.5 text-[11px] text-muted-foreground hover:text-foreground"
      >
        {hidden ? <CaretRight className="size-3" /> : <CaretDown className="size-3" />}
        {hidden ? "실행 결과 보기" : "실행 결과 숨기기"}
        <span className="ml-auto font-mono">{last.durationMs}ms</span>
      </button>
      {!hidden && (
        <div className="max-h-80 overflow-auto pb-1">
          {last.stdout && (
            <pre className="whitespace-pre-wrap px-3 py-1 font-mono text-xs leading-5">
              {last.stdout}
            </pre>
          )}
          {last.stderr && (
            <pre className="whitespace-pre-wrap bg-destructive/5 px-3 py-1 font-mono text-xs leading-5 text-destructive">
              {last.stderr}
            </pre>
          )}
          {isError ? (
            <div className="px-3 py-1 text-xs text-destructive">
              {last.summaryKo && <p>{last.summaryKo}</p>}
              {last.traceback && (
                <pre className="mt-1 overflow-x-auto font-mono text-[11px] leading-4">
                  {last.traceback}
                </pre>
              )}
            </div>
          ) : last.imageBlobId ? (
            <PreviewImage blobId={last.imageBlobId} />
          ) : preview?.kind === "table" ? (
            <PreviewTable preview={preview} />
          ) : preview?.kind === "repr" && preview.repr ? (
            <pre className="overflow-x-auto px-3 py-1 font-mono text-xs leading-5">
              {preview.repr}
            </pre>
          ) : null}
        </div>
      )}
    </div>
  );
}

/** 해당 셀(앵커)로 이동 (헤더 주소 버튼 · ⋮ 메뉴 공용) */
/** 시트에 놓인 첫 출력 위치 (없으면 블록 앵커) */
function sheetPos(block: PyBlock): { sheetId: string; r: number; c: number } {
  const o = outputsOf(block).find((x) => onSheet(block, x));
  return o
    ? { sheetId: o.sheetId ?? block.sheetId, r: o.anchor.r, c: o.anchor.c }
    : { sheetId: block.sheetId, r: block.anchor.r, c: block.anchor.c };
}

function goToAnchor(block: PyBlock): void {
  const st = store();
  const p = sheetPos(block);
  st.setActiveSheet(p.sheetId);
  st.setSelection({ r0: p.r, c0: p.c, r1: p.r, c1: p.c });
}

/** ⋮ 더보기 — 헤더에서 밀어낸 보조 조작 */
function MoreMenu({
  block,
  onRun,
  onNote,
}: {
  block: PyBlock;
  onRun: () => void;
  /** 부록 J.4: 설명 추가/편집 (코드 블록만) */
  onNote?: () => void;
}) {
  // 설명 추가/편집 선택 시 Radix가 트리거로 포커스를 되돌리면 방금 연 편집기가 blur로 닫힌다 — 억제
  const suppressRestore = useRef(false);
  const isMarkdown = block.kind === "markdown";
  const collapsed = !!block.collapsed;
  const firstOutputId = block.outputs?.[0]?.id;
  const picking = useWorkbookStore((s) => s.anchorPicking?.blockId === block.id);
  return (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <Button variant="ghost" size="icon-xs" aria-label="더보기" title="더보기">
          <DotsThreeVertical weight="bold" />
        </Button>
      </DropdownMenuTrigger>
      <DropdownMenuContent
        align="end"
        className="min-w-44"
        onCloseAutoFocus={(e) => {
          if (suppressRestore.current) {
            e.preventDefault();
            suppressRestore.current = false;
          }
        }}
      >
        {!isMarkdown && (
          <>
            <DropdownMenuItem onClick={onRun}>
              실행
              <DropdownMenuShortcut>Ctrl+Enter</DropdownMenuShortcut>
            </DropdownMenuItem>
            <DropdownMenuSeparator />
            <DropdownMenuLabel>출력 모드</DropdownMenuLabel>
            <DropdownMenuRadioGroup
              value={block.outputMode}
              onValueChange={(v) => store().setBlockOutputMode(block.id, v as OutputMode)}
            >
              <DropdownMenuRadioItem value="values">값 (셀로 펼치기)</DropdownMenuRadioItem>
              <DropdownMenuRadioItem value="object">객체 (카드)</DropdownMenuRadioItem>
            </DropdownMenuRadioGroup>
            <DropdownMenuSeparator />
          </>
        )}
        {(isMarkdown || firstOutputId) && (
          <DropdownMenuItem
            onClick={() =>
              store().setAnchorPicking(
                picking ? null : { blockId: block.id, outputId: firstOutputId ?? "" },
              )
            }
          >
            {isMarkdown ? "위치 지정" : "출력 위치 지정 (첫 출력)"}
            {picking && <DropdownMenuShortcut>지정 중</DropdownMenuShortcut>}
          </DropdownMenuItem>
        )}
        {!isMarkdown && onNote && (
          <DropdownMenuItem
            onClick={() => {
              suppressRestore.current = true;
              onNote();
            }}
          >
            {block.note === undefined ? "설명 추가" : "설명 편집"}
            <DropdownMenuShortcut>제목에서 Ctrl+Enter</DropdownMenuShortcut>
          </DropdownMenuItem>
        )}
        {outputsOf(block).some((o) => onSheet(block, o)) && (
          <DropdownMenuItem onClick={() => goToAnchor(block)}>해당 셀로 이동</DropdownMenuItem>
        )}
        <DropdownMenuItem
          onClick={() => store().setBlockCollapsed(block.id, !collapsed)}
        >
          {collapsed ? "펼치기" : "접기"}
        </DropdownMenuItem>
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

// ── 부록 J.1: 마크다운 이미지 내장 (data URI, 500KB 초과 시 캔버스 축소) ──

const MAX_IMG_BYTES = 500 * 1024;
const dataUriBytes = (uri: string): number =>
  Math.ceil(((uri.length - uri.indexOf(",") - 1) * 3) / 4);

async function imageFileToDataUri(file: File): Promise<string> {
  if (file.size <= MAX_IMG_BYTES) {
    return await new Promise<string>((res, rej) => {
      const r = new FileReader();
      r.onload = () => res(r.result as string);
      r.onerror = () => rej(new Error("read"));
      r.readAsDataURL(file);
    });
  }
  // 캔버스 리사이즈 → JPEG, 500KB 이하가 될 때까지 단계 축소
  const bmp = await createImageBitmap(file);
  for (let scale = 1; scale >= 0.1; scale *= 0.7) {
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(bmp.width * scale));
    canvas.height = Math.max(1, Math.round(bmp.height * scale));
    canvas.getContext("2d")!.drawImage(bmp, 0, 0, canvas.width, canvas.height);
    const uri = canvas.toDataURL("image/jpeg", 0.82);
    if (dataUriBytes(uri) <= MAX_IMG_BYTES) return uri;
  }
  throw new Error("too-large");
}

/** J.1 서식 툴바 버튼 정의 — 라벨은 한국어, 표시는 축약 기호 */
const MD_TOOLS: [MdAction, string, ReactNode][] = [
  ["h1", "제목1", "H1"],
  ["h2", "제목2", "H2"],
  ["h3", "제목3", "H3"],
  ["subheading", "하위 제목", "H+"],
  ["bold", "굵게", <b key="b">B</b>],
  ["ul", "기호 목록", "•―"],
  ["ol", "번호 목록", "1."],
  ["code", "인라인 코드", "<>"],
  ["hr", "구분선", "—"],
];

/** J.4 설명 전용 툴바 — 제목1/2/3 없음 (블록 제목이 사실상 #, 하위 제목이 ##부터 삽입) */
const NOTE_TOOLS = MD_TOOLS.filter(([a]) => a !== "h1" && a !== "h2" && a !== "h3");

export default function PyBlockCard({
  block,
  isFirst,
  isLast,
}: {
  block: PyBlock;
  /** 계산 순서상 처음/마지막 — ↑↓ 비활성화용 (패널이 계산해 넘긴다) */
  isFirst?: boolean;
  isLast?: boolean;
}) {
  const isMarkdown = block.kind === "markdown";
  const running = useWorkbookStore((s) => !!s.runningBlocks[block.id]);
  const dirty = useWorkbookStore((s) => !!s.dirtyBlocks[block.id]);
  const execCount = useWorkbookStore((s) => s.execCounts[block.id]);
  const hovered = useWorkbookStore((s) => s.hoverBlockId === block.id);
  const picking = useWorkbookStore((s) => s.anchorPicking?.blockId === block.id);
  const focusRequested = useWorkbookStore((s) => s.focusBlockId === block.id);
  // 헤더 주소 = 시트에 놓인 첫 출력 위치 (시트에 추가 전에는 블록 앵커 — 표시되지 않는다)
  const pos = sheetPos(block);
  const sheetName = useWorkbookStore(
    (s) => s.workbook.sheets.find((sh) => sh.id === pos.sheetId)?.name ?? "?",
  );
  const collapsed = !!block.collapsed;
  const sheetOut = !isMarkdown && block.sheetOut === true;
  // 실제로 셀에 놓인 출력이 있을 때만 주소·'해당 셀로 이동'을 보인다 (위치 미정은 셀이 없다)
  const placed = sheetOut && outputsOf(block).some((o) => onSheet(block, o));
  const cardRef = useRef<HTMLDivElement>(null);
  const mdRef = useRef<HTMLTextAreaElement>(null);
  const noteRef = useRef<HTMLTextAreaElement>(null);
  // J.4: 설명(note) 편집/미리보기 — 마크다운 본문과 동일 패턴
  const [editingNote, setEditingNote] = useState(false);
  const codeRef = useRef(block.code);
  const debounceRef = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  // 노트북처럼: 새 마크다운은 편집 상태로 열리고, 미리보기 더블클릭으로 다시 편집
  const [editingMd, setEditingMd] = useState(isMarkdown && !block.markdown);
  // §4.7 <640: 인라인 편집 대신 전체 화면 편집기
  const [narrow, setNarrow] = useState(false);
  const [editorOpen, setEditorOpen] = useState(false);
  useEffect(() => {
    const mq = window.matchMedia("(max-width: 639px)");
    const update = () => setNarrow(mq.matches);
    update();
    mq.addEventListener("change", update);
    return () => mq.removeEventListener("change", update);
  }, []);

  useEffect(() => () => clearTimeout(debounceRef.current), []);

  // 외부 변경(undo 등)으로 코드가 바뀌면 ▶ 커밋 기준도 따라간다
  useEffect(() => {
    codeRef.current = block.code;
  }, [block.code]);

  // 목차·블록으로 이동 → 카드를 펼치고 카드 상단을 화면 상단에 맞춘다
  // (코드 블록 포커스는 CodeEditor가 스크롤 없이 처리 — 자식 effect가 먼저 돌고 이 정렬이 마지막)
  useEffect(() => {
    if (!focusRequested) return;
    if (collapsed) store().setBlockCollapsed(block.id, false);
    cardRef.current?.scrollIntoView({ block: "start" });
    if (isMarkdown) {
      mdRef.current?.focus({ preventScroll: true });
      store().setFocusBlock(null);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [focusRequested, isMarkdown]);

  /** 코드 확정. notify면 자동 재계산/dirty 배지 통지 (§2.3.3 코드 저장 → dirty) */
  const commitCode = (value: string, notify: boolean) => {
    clearTimeout(debounceRef.current);
    const changed =
      store().workbook.pyBlocks.find((b) => b.id === block.id)?.code !== value;
    store().setBlockCode(block.id, value);
    if (notify && changed) notifyCodeEdit(block.id); // 수정 중엔 실행하지 않는다 — 이전 결과 유지, ▶로 실행
  };

  const onChange = (value: string) => {
    codeRef.current = value;
    clearTimeout(debounceRef.current);
    debounceRef.current = setTimeout(() => commitCode(value, true), 500);
  };

  const run = () => {
    commitCode(codeRef.current, false); // ▶가 직접 실행하므로 통지 없이 확정만 (이중 실행 방지)
    void runBlock(block.id);
  };

  const anchorLabel = `${sheetName}!${formatA1({ r0: pos.r, c0: pos.c, r1: pos.r, c1: pos.c })}`;

  const rendered = useMemo(
    () => (isMarkdown ? renderMarkdown(block.markdown ?? "") : null),
    [isMarkdown, block.markdown],
  );
  // J.4: 설명(note) 미리보기 — 마크다운 본문과 동일 렌더러
  const renderedNote = useMemo(
    () => (!isMarkdown && block.note?.trim() ? renderMarkdown(block.note) : null),
    [isMarkdown, block.note],
  );

  // 편집 대상 필드 — 마크다운 본문 또는 코드 블록 설명(note, 부록 J.4)
  const fieldRef = (field: "markdown" | "note") =>
    field === "markdown" ? mdRef.current : noteRef.current;
  const fieldSet = (field: "markdown" | "note", value: string) => {
    if (field === "markdown") store().setBlockMarkdown(block.id, value);
    else store().setBlockNote(block.id, value);
  };

  /** J.1 서식 툴바 — textarea 선택 기준으로 문법 삽입/감싸기 (undo는 코드 커밋과 동일 1회).
   *  note는 baseLevel 1: 하위 제목이 ##부터 (블록 제목이 사실상 #, 부록 J.4) */
  const applyTool = (action: MdAction, field: "markdown" | "note") => {
    const ta = fieldRef(field);
    if (!ta) return;
    const res = applyMdAction(
      ta.value,
      ta.selectionStart,
      ta.selectionEnd,
      action,
      field === "note" ? 1 : 0,
    );
    fieldSet(field, res.text);
    requestAnimationFrame(() => {
      ta.focus();
      ta.setSelectionRange(res.start, res.end);
    });
  };

  /** J.1 이미지 → data URI → 커서 위치에 ![이름](…) 삽입 (본문·설명 공용) */
  const insertImage = useCallback(
    async (file: File, field: "markdown" | "note") => {
      const big = file.size > MAX_IMG_BYTES;
      if (big) toast.loading("이미지를 500KB 이하로 줄이는 중…", { id: "md-img" });
      try {
        const uri = await imageFileToDataUri(file);
        const b = useWorkbookStore.getState().workbook.pyBlocks.find((x) => x.id === block.id);
        const cur = (field === "markdown" ? b?.markdown : b?.note) ?? "";
        const ta = field === "markdown" ? mdRef.current : noteRef.current;
        const pos = ta?.selectionStart ?? cur.length;
        fieldSet(field, `${cur.slice(0, pos)}![${file.name}](${uri})${cur.slice(pos)}`);
        if (big) toast.success("이미지를 넣었습니다 (축소됨)", { id: "md-img" });
      } catch {
        toast.error("이미지를 넣지 못했습니다 — 500KB 이하로 줄일 수 없습니다", {
          id: "md-img",
        });
      }
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [block.id],
  );

  /** 이미지 붙여넣기·드롭 핸들러 (본문·설명 공용) — 워크북 교체 DnD와 분리 */
  const imageDndProps = (field: "markdown" | "note") => ({
    onPaste: (e: React.ClipboardEvent) => {
      const file = Array.from(e.clipboardData.files).find((f) => f.type.startsWith("image/"));
      if (file) {
        e.preventDefault();
        void insertImage(file, field);
      }
    },
    onDragOver: (e: React.DragEvent) => {
      if (e.dataTransfer.types.includes("Files")) {
        e.preventDefault();
        e.stopPropagation();
      }
    },
    onDrop: (e: React.DragEvent) => {
      const file = Array.from(e.dataTransfer.files).find((f) => f.type.startsWith("image/"));
      if (file) {
        e.preventDefault();
        e.stopPropagation();
        void insertImage(file, field);
      }
    },
  });

  /** J.4: 설명 영역 열기 — 없으면 빈 note 생성, 편집 모드 + 포커스 */
  const openNote = () => {
    if (collapsed) store().setBlockCollapsed(block.id, false);
    const cur = useWorkbookStore.getState().workbook.pyBlocks.find((b) => b.id === block.id);
    if (cur?.note === undefined) store().setBlockNote(block.id, "");
    setEditingNote(true);
    // 이중 rAF: Radix 메뉴가 닫히며 트리거로 포커스를 되돌린 뒤에 편집기로 포커스
    requestAnimationFrame(() => requestAnimationFrame(() => noteRef.current?.focus()));
  };

  // 제목 폴백 (부록 F.3) — 표시 전용, 스토어에 쓰지 않는다
  const fallbackTitle = useMemo(
    () => (isMarkdown ? "" : codeTitle(block.code)),
    [isMarkdown, block.code],
  );

  return (
    <div
      ref={cardRef}
      className={cn(
        // 한 묶음(제목·설명·코드·출력)이 하나의 카드로 읽히도록: 라운드 클리핑 + 좌측 종류 바.
        // 코드 블록은 Sky Blue(파이썬 관여), 마크다운은 중립 회색 — 이웃 카드와 시각적으로 분리된다.
        "group relative scroll-mt-1 overflow-hidden rounded-md border bg-card shadow-sm transition-shadow",
        "border-l-[3px]",
        isMarkdown ? "border-l-muted-foreground/35" : "border-l-primary/70",
        hovered && "border-primary/60 shadow-[0_0_0_2px_#EAF3FA]", // spill hover → 카드 강조 (§4.8)
        picking && "border-primary",
        focusRequested && "ring-1 ring-primary/40", // 목차·이동으로 지목된 블록
      )}
      data-block-id={block.id}
      data-block-kind={isMarkdown ? "markdown" : "code"}
    >
      {/* 헤더 (부록 F.3) — [접기][제목 맨 앞][상태·dirty·마크다운][앵커 주소 맨 뒤].
          접기 시에도 이 줄은 통째로 남는다 — 숨는 것은 아래 본문뿐 */}
      <div className="flex items-center gap-1.5 border-b bg-muted/60 px-2 py-1">
        <button
          onClick={() => store().setBlockCollapsed(block.id, !collapsed)}
          className="text-muted-foreground hover:text-foreground"
          aria-label={collapsed ? "블록 펼치기" : "블록 접기"}
          aria-expanded={!collapsed}
          title={collapsed ? "펼치기" : "접기"}
        >
          {collapsed ? <CaretRight className="size-3.5" /> : <CaretDown className="size-3.5" />}
        </button>
        {isMarkdown ? (
          <span className="min-w-0 flex-1 truncate text-xs font-medium">
            {block.title || <span className="text-muted-foreground">제목 없음</span>}
          </span>
        ) : (
          <Input
            value={block.title ?? ""}
            onChange={(e) => store().setBlockTitle(block.id, e.target.value)}
            // 제목이 비어 있으면 코드 첫 주석에서 유도한 표시 전용 제목 (저장 안 함)
            placeholder={fallbackTitle || "제목 없음"}
            // J.4 추천 제목 채택: 비어 있고 폴백이 있으면 포커스 시 실제 값으로 시드 + 전체 선택
            onFocus={(e) => {
              if (!block.title && fallbackTitle) {
                store().setBlockTitle(block.id, fallbackTitle);
                const input = e.target;
                requestAnimationFrame(() => input.select());
              }
            }}
            // J.4: 제목에서 Ctrl+Enter → 설명(note) 편집으로
            onKeyDown={(e) => {
              if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
                e.preventDefault();
                openNote();
              }
            }}
            aria-label="블록 제목"
            className="h-6 min-w-0 flex-1 px-1.5"
          />
        )}
        {isMarkdown ? (
          <Badge variant="secondary">마크다운</Badge>
        ) : (
          <>
            {statusBadge(block, running)}
            <button
              onClick={() => {
                // 처음엔 결과 하나를 시트에 놓고, 다시 누를 때마다 출력이 하나씩 늘어난다
                // 셀이 정해지지 않은 출력을 만들고 바로 '셀 선택' 상태로 — 셀을 고르면 그때 실행·반영된다
                const id = store().addOutput(block.id);
                // (상태 바가 '결과를 놓을 셀을 클릭하세요 · Esc 취소'를 띄운다)
                if (id) store().setAnchorPicking({ blockId: block.id, outputId: id });
              }}
              title={
                sheetOut
                  ? "출력을 하나 더 시트에 추가합니다 — 위치·변수·열·행은 아래에서 고릅니다"
                  : "결과를 시트에 추가합니다 — 기본은 Python 결과로만 봅니다"
              }
              className="flex shrink-0 items-center gap-0.5 rounded border px-1.5 py-0.5 text-[11px] text-muted-foreground hover:bg-accent hover:text-foreground"
            >
              <Plus className="size-3" />
              시트에 추가
            </button>
            {dirty && !running && (
              <span
                className="size-1.5 shrink-0 rounded-full bg-warning"
                title="변경됨 — 재실행 필요"
                aria-label="dirty"
              />
            )}
          </>
        )}
        {/* 셀 툴바 — 헤더 행 안에 공간을 항상 예약(인라인)해 아무것도 가리지 않는다.
            hover·포커스에서만 보이고, DOM에는 항상 있어 Tab으로 닿는다 */}
        <div
          data-testid="cell-toolbar"
          className="pointer-events-none flex shrink-0 items-center rounded border bg-card p-0.5 opacity-0 transition-opacity group-focus-within:pointer-events-auto group-focus-within:opacity-100 group-hover:pointer-events-auto group-hover:opacity-100"
        >
        <Button
          variant="ghost"
          size="icon-xs"
          disabled={isFirst}
          onClick={() => moveBlock(block.id, "up")}
          aria-label="위로"
          title={isFirst ? "첫 블록입니다" : "위로 — 앞 블록과 실행 순서·자리 바꾸기"}
        >
          <ArrowUp />
        </Button>
        <Button
          variant="ghost"
          size="icon-xs"
          disabled={isLast}
          onClick={() => moveBlock(block.id, "down")}
          aria-label="아래로"
          title={isLast ? "마지막 블록입니다" : "아래로 — 뒤 블록과 실행 순서·자리 바꾸기"}
        >
          <ArrowDown />
        </Button>
        {isMarkdown && (
          <Button
            variant="ghost"
            size="icon-xs"
            onMouseDown={(e) => e.preventDefault()} // 편집 중 blur(→미리보기)가 먼저 발생해 클릭 토글이 원상복구되는 것 방지
            onClick={() => setEditingMd((v) => !v)}
            aria-label="편집/미리보기 전환"
            title="편집/미리보기 전환"
          >
            <NotePencil />
          </Button>
        )}
        <Button
          variant="ghost"
          size="icon-xs"
          onClick={() => store().removePyBlock(block.id)}
          aria-label="블록 삭제"
          title="블록 삭제"
        >
          <TrashSimple />
        </Button>
        <MoreMenu block={block} onRun={run} onNote={isMarkdown ? undefined : openNote} />
      </div>
        {placed && (
          <button
            onClick={() => goToAnchor(block)}
            className="shrink-0 font-mono text-xs text-foreground/80 hover:text-primary"
            title="해당 셀로 이동"
          >
            {anchorLabel}
          </button>
        )}
      </div>

      {/* 본문 — 왼쪽 원형 실행 레일 + 내용 */}
      {!collapsed && (
        <div className="flex">
          <div className="flex w-9 shrink-0 flex-col items-center gap-0.5 py-1.5">
            {isMarkdown ? (
              <Button
                variant="ghost"
                className="size-7 rounded-full p-0"
                onClick={() => setEditingMd((v) => !v)}
                aria-label={editingMd ? "마크다운 미리보기" : "마크다운 편집"}
                title={editingMd ? "미리보기" : "편집"}
              >
                {editingMd ? <Eye /> : <NotePencil />}
              </Button>
            ) : (
              <Button
                variant="ghost"
                className="size-7 rounded-full border p-0 hover:bg-accent"
                onClick={run}
                disabled={running}
                aria-label="실행"
                title="실행 (Ctrl+Enter)"
              >
                <Play weight="fill" className="text-primary" />
              </Button>
            )}
            {/* 노트북식 실행 순번 — 실행 중 [*], 미실행이면 비워 둔다 */}
            {!isMarkdown && (running || execCount !== undefined) && (
              <span
                data-testid="exec-count"
                className="font-mono text-[10px] leading-none text-muted-foreground"
                title={running ? "실행 중" : `${execCount}번째로 실행됨`}
              >
                [{running ? "*" : execCount}]
              </span>
            )}
          </div>

          <div className="min-w-0 flex-1 border-l">
            {isMarkdown ? (
              editingMd ? (
                <>
                  {/* J.1 서식 툴바 — mousedown preventDefault로 textarea blur(→미리보기 전환) 방지 */}
                  <div
                    data-testid="md-toolbar"
                    className="flex flex-wrap items-center gap-0.5 border-b bg-muted/30 px-1 py-0.5"
                  >
                    {MD_TOOLS.map(([action, label, glyph]) => (
                      <button
                        key={action}
                        type="button"
                        onMouseDown={(e) => e.preventDefault()}
                        onClick={() => applyTool(action, "markdown")}
                        aria-label={label}
                        title={label}
                        className="min-w-6 rounded px-1.5 py-0.5 font-mono text-[11px] font-medium text-muted-foreground hover:bg-accent hover:text-foreground"
                      >
                        {glyph}
                      </button>
                    ))}
                  </div>
                  <textarea
                    ref={mdRef}
                    value={block.markdown ?? ""}
                    onChange={(e) => store().setBlockMarkdown(block.id, e.target.value)}
                    onBlur={() => setEditingMd(false)}
                    {...imageDndProps("markdown")}
                    rows={6}
                    placeholder={
                      "# 제목\n\n설명을 적으세요. **굵게**, `코드`, [링크](https://example.com)\n이미지는 드래그 앤 드롭 또는 붙여넣기"
                    }
                    aria-label="마크다운"
                    className="w-full resize-y bg-card p-2 font-mono text-xs outline-none placeholder:text-muted-foreground"
                  />
                </>
              ) : (
                <div
                  onDoubleClick={() => setEditingMd(true)}
                  data-testid="markdown-preview"
                  className="space-y-1 px-2 py-2"
                  title="더블클릭하여 편집"
                >
                  {block.markdown?.trim() ? (
                    rendered
                  ) : (
                    <p className="text-xs text-muted-foreground">
                      더블클릭하여 마크다운을 입력하세요
                    </p>
                  )}
                </div>
              )
            ) : (
              <>
                {/* J.4 설명(note) — 헤더 아래·본문 위, 편집/미리보기. 없으면 영역 자체가 없다 */}
                {block.note !== undefined && (
                  <div className="border-b" data-testid="block-note">
                    {editingNote ? (
                      <>
                        <div
                          data-testid="note-toolbar"
                          className="flex flex-wrap items-center gap-0.5 border-b bg-muted/30 px-1 py-0.5"
                        >
                          {NOTE_TOOLS.map(([action, label, glyph]) => (
                            <button
                              key={action}
                              type="button"
                              onMouseDown={(e) => e.preventDefault()} // blur(→미리보기) 방지
                              onClick={() => applyTool(action, "note")}
                              aria-label={label}
                              title={label}
                              className="min-w-6 rounded px-1.5 py-0.5 font-mono text-[11px] font-medium text-muted-foreground hover:bg-accent hover:text-foreground"
                            >
                              {glyph}
                            </button>
                          ))}
                          <button
                            type="button"
                            onMouseDown={(e) => e.preventDefault()}
                            onClick={() => {
                              store().setBlockNote(block.id, null); // 확인 없이 삭제, 1 undo
                              setEditingNote(false);
                            }}
                            aria-label="설명 삭제"
                            title="설명 삭제"
                            className="ml-auto rounded px-1.5 py-0.5 text-[11px] text-muted-foreground hover:bg-accent hover:text-destructive"
                          >
                            ✕
                          </button>
                        </div>
                        <textarea
                          ref={noteRef}
                          value={block.note}
                          onChange={(e) => store().setBlockNote(block.id, e.target.value)}
                          onBlur={() => setEditingNote(false)}
                          {...imageDndProps("note")}
                          rows={3}
                          placeholder="블록 설명 (마크다운) — ## 헤딩은 목차 서브 항목이 됩니다"
                          aria-label="설명"
                          className="w-full resize-y bg-card p-2 font-mono text-xs outline-none placeholder:text-muted-foreground"
                        />
                      </>
                    ) : (
                      <div
                        onDoubleClick={() => {
                          setEditingNote(true);
                          requestAnimationFrame(() => noteRef.current?.focus());
                        }}
                        data-testid="note-preview"
                        className="space-y-1 px-2 py-1.5"
                        title="더블클릭하여 편집"
                      >
                        {renderedNote ?? (
                          <p className="text-xs text-muted-foreground">
                            더블클릭하여 설명 입력
                          </p>
                        )}
                      </div>
                    )}
                  </div>
                )}
                {sheetOut && <OutputList block={block} />}
                {narrow ? (
                  <>
                    <button
                      onClick={() => setEditorOpen(true)}
                      className="w-full truncate bg-code-bg px-2 py-2 text-left font-mono text-xs text-muted-foreground"
                    >
                      {block.code.split("\n")[0] || "코드 편집 (전체 화면)…"}
                    </button>
                    <Dialog open={editorOpen} onOpenChange={setEditorOpen}>
                      <DialogContent className="h-[85vh] max-w-full p-3">
                        <DialogHeader>
                          <DialogTitle className="text-sm">{anchorLabel} 코드 편집</DialogTitle>
                        </DialogHeader>
                        <div className="min-h-0 flex-1 overflow-auto rounded border">
                          <CodeEditor
                            blockId={block.id}
                            sheetId={block.sheetId}
                            value={block.code}
                            onChange={onChange}
                            onRun={run}
                            className="max-h-full"
                          />
                        </div>
                        <Button onClick={run} disabled={running} className="shrink-0">
                          실행 (Ctrl+Enter)
                        </Button>
                      </DialogContent>
                    </Dialog>
                  </>
                ) : (
                  <CodeEditor
                    blockId={block.id}
                    sheetId={block.sheetId}
                    value={block.code}
                    onChange={onChange}
                    onRun={run}
                    placeholder={'df = sheet("A1:C10", headers=True)\ndf.describe()'}
                  />
                )}
                <CellResult block={block} />
                <AiAssist block={block} />
              </>
            )}
          </div>
        </div>
      )}
    </div>
  );
}
