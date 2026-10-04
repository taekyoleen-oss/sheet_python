"use client";

// Python 패널 — 헤더(블록 추가·전체 실행·중단·계산 모드·목차/AI 토글) + 블록 목록(계산 순서)
// + 스니펫·초기화 스크립트 + 참조 삽입 바 (§2.3.2).
// Python 조작은 전부 이 패널 헤더에 있다 — 시트 편집은 그리드 패널 안 SheetEditToolbar.
// 목차는 전용 패널(TocPanel)로 분리되었다 (부록 D.2).

import { useMemo } from "react";
import { Article, ChatCircleText, ListBullets, Play, Plus, Stop } from "@phosphor-icons/react";
import { Button } from "@/components/ui/button";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  Tooltip,
  TooltipContent,
  TooltipTrigger,
} from "@/components/ui/tooltip";
import { ToolButton } from "@/components/shell/ShellBar";
import { AiGenerateBar } from "@/components/python/AiAssist";
import InitScriptDialog from "@/components/python/InitScriptDialog";
import PyBlockCard from "@/components/python/PyBlockCard";
import SnippetInsertDialog from "@/components/python/SnippetInsertDialog";
import SnippetMenu from "@/components/python/SnippetMenu";
import { editorRegistry } from "@/components/python/CodeEditor";
import { setCalcModeEverywhere } from "@/lib/grid/calc-host";
import {
  addBlockAtSelection,
  addMarkdownAtSelection,
  blocksInOrder,
  runAllBlocks,
} from "@/lib/grid/run-block";
import { useWorkbookStore } from "@/lib/grid/model";
import { xlRefForSelection } from "@/lib/grid/xl-ref";
import { getRuntimeClient } from "@/lib/runtime/client";
import { saveSettings } from "@/lib/storage/db";
import type { CalcMode } from "@/types/workbook";

function RefInsertBar() {
  const selection = useWorkbookStore((s) => s.selection);
  const lastEditorBlockId = useWorkbookStore((s) => s.lastEditorBlockId);
  const block = useWorkbookStore((s) =>
    s.workbook.pyBlocks.find((b) => b.id === s.lastEditorBlockId),
  );
  if (!selection || !lastEditorBlockId || !block) return null;
  const ref = xlRefForSelection(block.sheetId);
  if (!ref) return null;

  return (
    <div className="border-t bg-accent/60 px-2 py-1.5">
      <button
        onClick={() => editorRegistry.get(lastEditorBlockId)?.insert(ref)}
        className="w-full truncate rounded border border-primary/40 bg-background px-2 py-1 text-left font-mono text-xs text-primary hover:bg-accent"
        title="블록 코드의 커서 위치에 참조 삽입"
      >
        {ref} 삽입
      </button>
    </div>
  );
}

export default function PythonPanel() {
  const workbook = useWorkbookStore((s) => s.workbook);
  const blocks = useMemo(() => blocksInOrder(workbook), [workbook]);
  const calcMode = useWorkbookStore((s) => s.workbook.calcMode);
  const tocOpen = useWorkbookStore((s) => s.tocOpen);
  const aiChatOpen = useWorkbookStore((s) => s.aiChatOpen);

  const allCollapsed = blocks.length > 0 && blocks.every((b) => b.collapsed);

  return (
    <div className="flex h-full flex-col border-l bg-code-bg">
      {/* Python 조작은 전부 이 헤더에 모인다. 좁은 패널에서는 아이콘 전용 + 줄바꿈(flex-wrap)으로 흘린다 */}
      <div
        data-testid="python-panel-header"
        className="flex flex-wrap items-center gap-x-0.5 gap-y-0.5 border-b px-2 py-1"
      >
        <span className="text-xs font-medium">블록</span>
        <span className="mr-1 text-xs text-muted-foreground">{blocks.length}</span>

        <ToolButton label="Python 블록 추가 (Ctrl+Shift+P)" onClick={addBlockAtSelection}>
          <Plus className="text-primary" />
        </ToolButton>
        <ToolButton label="마크다운 블록 추가" onClick={addMarkdownAtSelection}>
          <Article />
        </ToolButton>
        {/* 실행 버튼은 --primary 채움 (§4.6 Button) */}
        <Tooltip>
          <TooltipTrigger asChild>
            <Button
              variant="default"
              size="icon"
              className="ml-1 size-8"
              onClick={() => void runAllBlocks()}
              aria-label="전체 실행"
            >
              <Play weight="fill" />
            </Button>
          </TooltipTrigger>
          <TooltipContent>전체 실행 — 계산 순서대로 모든 블록</TooltipContent>
        </Tooltip>
        <ToolButton label="실행 중단" onClick={() => getRuntimeClient().interrupt()}>
          <Stop />
        </ToolButton>
        <Select value={calcMode} onValueChange={(v) => setCalcModeEverywhere(v as CalcMode)}>
          <SelectTrigger
            className="ml-0.5 mr-1 h-7 w-16 text-xs"
            aria-label="계산 모드"
            title="계산 모드"
          >
            <SelectValue />
          </SelectTrigger>
          <SelectContent>
            <SelectItem value="auto">자동</SelectItem>
            <SelectItem value="manual">수동</SelectItem>
          </SelectContent>
        </Select>
        <ToolButton
          label={tocOpen ? "목차 패널 닫기" : "목차 패널 열기"}
          active={tocOpen}
          onClick={() => {
            useWorkbookStore.getState().setTocOpen(!tocOpen);
            void saveSettings({ tocOpen: !tocOpen });
          }}
        >
          <ListBullets />
        </ToolButton>
        {/* 부록 G.2: AI 채팅 패널 토글 */}
        <ToolButton
          label={aiChatOpen ? "AI 채팅 패널 닫기" : "AI 채팅 패널 열기"}
          active={aiChatOpen}
          onClick={() => {
            useWorkbookStore.getState().setAiChatOpen(!aiChatOpen);
            void saveSettings({ aiChatOpen: !aiChatOpen });
          }}
        >
          <ChatCircleText />
        </ToolButton>

        <div className="ml-auto flex items-center">
          {blocks.length > 0 && (
            <Button
              variant="ghost"
              size="sm"
              className="h-6 px-2 text-xs"
              onClick={() => useWorkbookStore.getState().setAllCollapsed(!allCollapsed)}
            >
              {allCollapsed ? "모두 펼치기" : "모두 접기"}
            </Button>
          )}
          <SnippetMenu />
          <InitScriptDialog />
        </div>
      </div>
      {/* 상단 행 — 코드 삽입(부록 F.1) + ✦ AI 생성 바 */}
      <div className="flex items-center border-b pl-2">
        <SnippetInsertDialog />
        <AiGenerateBar />
      </div>

      <div className="min-h-0 flex-1 space-y-3 overflow-y-auto p-2">
        {blocks.length === 0 ? (
          <p className="px-2 py-8 text-center text-xs text-muted-foreground">
            셀을 선택하고 ＋ Python 블록을 누르세요
            <br />
            (Ctrl+Shift+P)
          </p>
        ) : (
          <>
            {blocks.map((block, i) => (
              <PyBlockCard
                key={block.id}
                block={block}
                isFirst={i === 0}
                isLast={i === blocks.length - 1}
              />
            ))}
            {/* 노트북식 끝 여백 — 마지막 카드도 상단 기준으로 맞출 수 있게 */}
            <div aria-hidden className="h-[70vh]" />
          </>
        )}
      </div>
      <RefInsertBar />
    </div>
  );
}
