"use client";

// 목차 전용 패널 (부록 D.2) — 마크다운 헤딩 계층 + 그 아래 코드 블록.
// 현재 블록은 primary 색 + 좌측 세로 바로 강조하고, hover 시 ▶(실행)·⋮(이동·이름·삭제)를 노출한다.

import { Fragment, useMemo } from "react";
import { DotsThreeVertical, Play, X } from "@phosphor-icons/react";
import { Button } from "@/components/ui/button";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { editorRegistry } from "@/components/python/CodeEditor";
import { blockSubEntries, buildToc, type SubEntry, type TocEntry } from "@/lib/grid/markdown";
import { useWorkbookStore } from "@/lib/grid/model";
import { blocksInOrder, runBlocks } from "@/lib/grid/run-block";
import { cn } from "@/lib/utils";

const store = () => useWorkbookStore.getState();

/** 실행 상태 점 — 오류는 destructive, 성공은 primary(Python 관여) */
function StatusDot({ entry }: { entry: TocEntry }) {
  if (entry.kind === "markdown") return null;
  const cls =
    entry.status === "error" || entry.status === "spill"
      ? "bg-destructive"
      : entry.status === "ok"
        ? "bg-primary"
        : "bg-muted-foreground/40";
  return <span className={cn("size-1.5 shrink-0 rounded-full", cls)} aria-hidden />;
}

/** 목차 클릭 — 카드 포커스 + 그리드 선택을 앵커로 (§부록 C.1 #5) */
function goToBlock(blockId: string): void {
  const st = store();
  const block = st.workbook.pyBlocks.find((b) => b.id === blockId);
  if (!block) return;
  st.setBlockCollapsed(blockId, false);
  st.setActiveSheet(block.sheetId);
  st.setSelection({
    r0: block.anchor.r,
    c0: block.anchor.c,
    r1: block.anchor.r,
    c1: block.anchor.c,
  });
  st.setSelectedBlock(blockId);
  requestAnimationFrame(() => store().setFocusBlock(blockId));
}

/** 서브 항목 클릭 (부록 F.2) — 카드 포커스 + 편집기에서 해당 줄로 스크롤 */
function goToSection(blockId: string, line: number): void {
  goToBlock(blockId);
  // goToBlock의 rAF(setFocusBlock) → 편집기 마운트·포커스 다음 프레임에 줄 스크롤.
  // 좁은 화면(편집기가 전체 화면 다이얼로그)에서는 registry가 비어 있어 포커스만 된다.
  requestAnimationFrame(() =>
    requestAnimationFrame(() => editorRegistry.get(blockId)?.scrollToLine(line)),
  );
}

/** 이름 바꾸기 — 카드로 이동한 뒤 제목 입력(마크다운은 본문)에 포커스 */
function renameBlock(blockId: string): void {
  goToBlock(blockId);
  requestAnimationFrame(() => {
    const card = document.querySelector<HTMLElement>(`[data-block-id="${blockId}"]`);
    const target = card?.querySelector<HTMLElement>(
      '[aria-label="블록 제목"], [aria-label="마크다운"]',
    );
    target?.focus();
  });
}

/** 마크다운 헤딩 항목 = 그 절(다음 동급 이상 헤딩 전까지)의 코드 블록 전부 */
function sectionBlockIds(toc: TocEntry[], index: number): string[] {
  const entry = toc[index];
  if (entry.kind === "code") return [entry.blockId];
  const ids: string[] = [];
  for (let i = index + 1; i < toc.length && toc[i].level > entry.level; i++) {
    if (toc[i].kind === "code") ids.push(toc[i].blockId);
  }
  return ids;
}

export function TocList() {
  const workbook = useWorkbookStore((s) => s.workbook);
  const blocks = useMemo(() => blocksInOrder(workbook), [workbook]);
  const toc = useMemo(() => buildToc(blocks), [blocks]);
  // 코드 블록 서브 항목 = 설명(note) 헤딩 + 섹션 주석 (부록 F.2 + J.4)
  const sections = useMemo(() => {
    const m = new Map<string, SubEntry[]>();
    for (const b of blocks) if (b.kind !== "markdown") m.set(b.id, blockSubEntries(b));
    return m;
  }, [blocks]);
  // 현재 항목: 최근 선택한 블록 → 없으면 마지막으로 편집한 블록 (한 블록의 첫 항목만 강조)
  const activeId = useWorkbookStore((s) => s.selectedBlockId ?? s.lastEditorBlockId);
  const activeKey = toc.find((e) => e.blockId === activeId)?.key;

  if (toc.length === 0) {
    return (
      <p className="px-2 py-8 text-center text-xs text-muted-foreground">
        마크다운 블록으로 제목을 추가하면 목차가 만들어집니다
      </p>
    );
  }

  return (
    <ul className="py-1">
      {toc.map((entry, i) => {
        const active = entry.key === activeKey;
        return (
          <Fragment key={entry.key}>
          <li
            className={cn(
              "group flex items-center border-l-2 pr-1",
              active ? "border-primary bg-accent/40" : "border-transparent",
            )}
          >
            <button
              onClick={() => goToBlock(entry.blockId)}
              // 접근성 이름은 제목만 — 3단계 · 접두는 시각 표시일 뿐이다
              aria-label={entry.label}
              style={{ paddingLeft: `${(entry.level - 1) * 12 + 6}px` }}
              className={cn(
                "flex min-w-0 flex-1 items-center gap-1.5 rounded py-1 text-left text-xs hover:bg-accent",
                active
                  ? "font-medium text-primary"
                  : entry.kind === "markdown"
                    ? "font-medium text-foreground"
                    : "font-mono text-muted-foreground",
              )}
            >
              <StatusDot entry={entry} />
              {entry.level >= 3 && <span className="shrink-0 text-muted-foreground">·</span>}
              <span className="truncate">{entry.label}</span>
            </button>
            <div className="flex shrink-0 items-center opacity-0 transition-opacity focus-within:opacity-100 group-hover:opacity-100">
              <Button
                variant="ghost"
                size="icon-xs"
                aria-label={`${entry.label} 실행`}
                title="이 절의 블록 실행"
                onClick={() => void runBlocks(sectionBlockIds(toc, i))}
              >
                <Play weight="fill" className="text-primary" />
              </Button>
              <DropdownMenu>
                <DropdownMenuTrigger asChild>
                  <Button variant="ghost" size="icon-xs" aria-label={`${entry.label} 메뉴`}>
                    <DotsThreeVertical weight="bold" />
                  </Button>
                </DropdownMenuTrigger>
                <DropdownMenuContent align="end">
                  <DropdownMenuItem onClick={() => goToBlock(entry.blockId)}>
                    해당 셀로 이동
                  </DropdownMenuItem>
                  <DropdownMenuItem onClick={() => renameBlock(entry.blockId)}>
                    이름 바꾸기
                  </DropdownMenuItem>
                  <DropdownMenuItem onClick={() => store().removePyBlock(entry.blockId)}>
                    삭제
                  </DropdownMenuItem>
                </DropdownMenuContent>
              </DropdownMenu>
            </div>
          </li>
          {/* 서브 항목: 설명(note) 헤딩 먼저 + 코드 섹션 주석 (부록 F.2 + J.4) */}
          {entry.kind === "code" &&
            (sections.get(entry.blockId) ?? []).map((s, j) => (
              <li
                key={`${entry.key}:s${j}`}
                className="flex items-center border-l-2 border-transparent pr-1"
                data-testid="toc-section"
              >
                <button
                  onClick={() =>
                    s.line === undefined
                      ? goToBlock(entry.blockId) // note 헤딩 — 카드 포커스로 충분
                      : goToSection(entry.blockId, s.line)
                  }
                  aria-label={s.label}
                  style={{ paddingLeft: `${(entry.level - 1 + s.depth) * 12 + 6}px` }}
                  className="flex min-w-0 flex-1 items-center gap-1.5 rounded py-0.5 text-left text-xs text-muted-foreground hover:bg-accent"
                >
                  <span className="shrink-0">·</span>
                  <span className="truncate">{s.label}</span>
                </button>
              </li>
            ))}
          </Fragment>
        );
      })}
    </ul>
  );
}

/** 전용 패널 — onClose가 있으면 헤더에 ✕ (좁은 화면에서는 탭이라 생략) */
export default function TocPanel({ onClose }: { onClose?: () => void }) {
  return (
    <div className="flex h-full flex-col border-l bg-card" data-testid="toc-panel">
      <div className="flex h-8 shrink-0 items-center gap-2 border-b px-2">
        <span className="text-xs font-medium">목차</span>
        {onClose && (
          <Button
            variant="ghost"
            size="icon-xs"
            className="ml-auto"
            onClick={onClose}
            aria-label="목차 패널 닫기"
            title="닫기"
          >
            <X />
          </Button>
        )}
      </div>
      <div className="min-h-0 flex-1 overflow-y-auto">
        <TocList />
      </div>
    </div>
  );
}
