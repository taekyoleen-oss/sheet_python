"use client";

// 셸 바 (2행) — 화면 전체에 대한 조작만 모은다.
// 왼쪽: [스프레드시트 감추기] | [파일 ▾][샘플 워크북][최근 워크북][? 단축키]
// 오른쪽: [(Python 접힘 시 ▶ 전체 실행)][Python 패널 감추기]
// 시트 편집(붙여넣기·행/열·정렬·서식)은 그리드 패널 안 SheetEditToolbar로 내려갔다 —
// 스프레드시트를 접으면 함께 사라진다. Python 조작은 PythonPanel 헤더에 있다.

import { type ReactNode } from "react";
import { Code, Play, SidebarSimple, Table } from "@phosphor-icons/react";
import FileMenu, {
  RecentWorkbookMenu,
  SampleWorkbookMenu,
} from "@/components/shell/FileMenu";
import { Button } from "@/components/ui/button";
import { Separator } from "@/components/ui/separator";
import {
  Tooltip,
  TooltipContent,
  TooltipTrigger,
} from "@/components/ui/tooltip";
import { useWorkbookStore } from "@/lib/grid/model";
import { runAllBlocks } from "@/lib/grid/run-block";
import { saveSettings } from "@/lib/storage/db";

/** 부록 P: 속성 창 열기/닫기 (셸 바 버튼·창의 ✕ 공용, 설정에 저장) */
export function toggleProps(next?: boolean): void {
  const st = useWorkbookStore.getState();
  const target = next ?? !st.propsOpen;
  st.setPropsOpen(target);
  void saveSettings({ propsOpen: target });
}

/**
 * 그리드·Python 패널 접기 토글 (툴바·세로 스트립·단축키 공용).
 * next 생략 시 현재 상태를 뒤집는다. 접힘 상태는 앱 설정에 저장 — 새로고침해도 유지.
 */
export function togglePanelCollapse(panel: "grid" | "python", next?: boolean): void {
  const st = useWorkbookStore.getState();
  const current = panel === "grid" ? st.gridCollapsed : st.pyCollapsed;
  const target = next ?? !current;
  if (target === current) return; // 이미 그 상태 — 불필요한 설정 쓰기 방지
  st.setPanelCollapsed(panel, target);
  const { gridCollapsed, pyCollapsed } = useWorkbookStore.getState();
  // 접힘은 세션 한정 — 앱을 새로 열면 항상 시트·Python 패널이 함께 보이는 것이 기본값이다
}

/** 아이콘 전용 툴 버튼 + 툴팁 — 셸 바·시트 편집 툴바·Python 패널 헤더 공용 */
export function ToolButton({
  label,
  onClick,
  disabled,
  soon,
  active,
  children,
}: {
  label: string;
  onClick?: () => void;
  disabled?: boolean;
  /** 후속 마일스톤 자리표시 */
  soon?: boolean;
  active?: boolean;
  children: ReactNode;
}) {
  return (
    <Tooltip>
      <TooltipTrigger asChild>
        <span tabIndex={disabled ? 0 : undefined}>
          <Button
            variant="ghost"
            size="icon"
            className={active ? "size-8 text-primary" : "size-8"}
            onClick={onClick}
            disabled={disabled}
            aria-label={label}
            aria-pressed={active}
          >
            {children}
          </Button>
        </span>
      </TooltipTrigger>
      <TooltipContent>{soon ? `${label} — 곧 제공` : label}</TooltipContent>
    </Tooltip>
  );
}

/** ? 단축키 안내 — 헤더에서 셸 바로 이동 */
function ShortcutHelp() {
  return (
    <Tooltip>
      <TooltipTrigger asChild>
        <button
          aria-label="키보드 단축키 안내"
          className="flex size-6 shrink-0 items-center justify-center rounded-full border text-xs text-muted-foreground hover:bg-muted"
        >
          ?
        </button>
      </TooltipTrigger>
      <TooltipContent className="max-w-64">
        <p className="mb-1 font-medium">키보드 단축키</p>
        <ul className="space-y-0.5 text-xs">
          <li>Ctrl+Shift+P — Python 블록 추가</li>
          <li>Ctrl+Enter — 블록 실행 (편집기)</li>
          <li>Ctrl+Z / Ctrl+Y — 실행 취소 / 다시 실행</li>
          <li>Ctrl(또는 Alt)+1 — 그리드로 포커스</li>
          <li>Ctrl(또는 Alt)+2 — Python 편집기로 포커스</li>
          <li>Ctrl(또는 Alt)+3 — 속성 창(변수·콘솔·진단·파일)으로 포커스</li>
          <li>Ctrl+Alt+1 — 스프레드시트 접기/펼치기</li>
          <li>Ctrl+Alt+2 — Python 패널 접기/펼치기</li>
          <li>Ctrl+Alt+3 — 속성 창 열기/닫기</li>
          <li>Esc — 스프레드시트 전체 화면 끝내기</li>
        </ul>
      </TooltipContent>
    </Tooltip>
  );
}

export default function ShellBar() {
  const gridCollapsed = useWorkbookStore((s) => s.gridCollapsed);
  const pyCollapsed = useWorkbookStore((s) => s.pyCollapsed);
  const propsOpen = useWorkbookStore((s) => s.propsOpen);

  return (
    <div
      data-testid="shell-bar"
      className="flex h-10 shrink-0 items-center gap-1 border-b bg-muted/40 px-2"
    >
      {/* 패널 접기 — 1024px 미만은 탭 전환 UI라 접기 개념이 없다 (lg:contents = 레이아웃 무영향) */}
      <span className="hidden lg:contents">
        <ToolButton
          label={gridCollapsed ? "스프레드시트 보이기 (Ctrl+Alt+1)" : "스프레드시트 감추기 (Ctrl+Alt+1)"}
          active={gridCollapsed}
          onClick={() => togglePanelCollapse("grid")}
        >
          <Table />
        </ToolButton>
        <Separator orientation="vertical" className="mx-1 h-5" />
      </span>

      {/* 파일 계열 — 감추기 버튼 오른쪽 (헤더에서 내려왔다) */}
      <FileMenu />
      <SampleWorkbookMenu />
      <RecentWorkbookMenu />
      <ShortcutHelp />

      {/* Python 패널 접기 — 셸 바 오른쪽 끝 (스프레드시트 접기와 좌·우로 마주본다) */}
      <span className="ml-auto hidden items-center gap-1 lg:flex">
        {/* 패널이 접혀 있으면 ▶ 전체 실행만 여기로 되돌아온다 — 접힌 채로도 재계산 가능 */}
        {pyCollapsed && (
          <Tooltip>
            <TooltipTrigger asChild>
              <Button
                variant="default"
                size="icon"
                className="size-8"
                onClick={() => void runAllBlocks()}
                aria-label="전체 실행"
              >
                <Play weight="fill" />
              </Button>
            </TooltipTrigger>
            <TooltipContent>전체 실행 — 계산 순서대로 모든 블록</TooltipContent>
          </Tooltip>
        )}
        <ToolButton
          label={pyCollapsed ? "Python 패널 보이기 (Ctrl+Alt+2)" : "Python 패널 감추기 (Ctrl+Alt+2)"}
          active={pyCollapsed}
          onClick={() => togglePanelCollapse("python")}
        >
          <Code />
        </ToolButton>
      </span>
      {/* 부록 P: 속성 창(변수·파일) — 창이 오른쪽에 열리므로 셸 바 오른쪽 끝 */}
      <span className="ml-auto flex items-center lg:ml-0">
        <Separator orientation="vertical" className="mx-1 hidden h-5 lg:block" />
        <ToolButton
          label={propsOpen ? "속성 창 닫기" : "속성 창 열기 (변수·파일)"}
          active={propsOpen}
          onClick={() => toggleProps()}
        >
          <SidebarSimple />
        </ToolButton>
      </span>
    </div>
  );
}
