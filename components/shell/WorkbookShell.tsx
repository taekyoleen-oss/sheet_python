"use client";

// 전체 레이아웃 — 설계서 §4.3: 헤더 / 툴바 / [그리드+시트 탭 | Python 패널] / 하단 패널 / 상태 바

import { useEffect, useRef, useState } from "react";
import { CaretLeft, CaretRight } from "@phosphor-icons/react";
import {
  ResizableHandle,
  ResizablePanel,
  ResizablePanelGroup,
} from "@/components/ui/resizable";
import { TooltipProvider } from "@/components/ui/tooltip";
import PasteImportDialog, {
  applyInternalPaste,
  fillSelection,
  startPasteFlow,
} from "@/components/grid/PasteImportDialog";
import FormulaBar from "@/components/grid/FormulaBar";
import SheetEditToolbar from "@/components/grid/SheetEditToolbar";
import SheetGrid from "@/components/grid/SheetGrid";
import SheetTabs from "@/components/grid/SheetTabs";
import AiChatPanel from "@/components/ai-chat/AiChatPanel";
import PropertiesPanel from "@/components/panels/PropertiesPanel";
import FitGuideDialog from "@/components/python/FitGuideDialog";
import PythonPanel from "@/components/python/PythonPanel";
import TocPanel from "@/components/python/TocPanel";
import {
  loadWorkbookData,
  openWorkbookFile,
  loadDefaultWorkbook,
} from "@/components/shell/FileMenu";
import Header from "@/components/shell/Header";
import { RuntimeStatus } from "@/components/shell/RuntimeStatus";
import ShellBar, { togglePanelCollapse, toggleProps } from "@/components/shell/ShellBar";
import StatusBar from "@/components/shell/StatusBar";
import dynamic from "next/dynamic";

// 참조 콘텐츠(~1MB 정적 데이터 + KaTeX)는 첫 전환 시에만 로드 — 워크북 첫 페인트 보호
const ReferenceView = dynamic(() => import("@/components/reference/ReferenceView"), {
  ssr: false,
  loading: () => (
    <p className="p-8 text-center text-sm text-muted-foreground">참조 콘텐츠 불러오는 중…</p>
  ),
});
import { internalCopyFor, rememberCopy } from "@/lib/grid/clipboard/internal";
import { parseClipboard } from "@/lib/grid/clipboard/parse";
import { serializeRange } from "@/lib/grid/clipboard/serialize";
import { useWorkbookStore } from "@/lib/grid/model";
import { addBlockAtSelection } from "@/lib/grid/run-block";
import { DEFAULT_INIT_SCRIPT, getRuntimeClient } from "@/lib/runtime/client";
import { startWorkDirSync } from "@/lib/runtime/workdir-sync";
import { useAutosave } from "@/lib/storage/autosave";
import { getWorkbook, loadSettings, saveSettings } from "@/lib/storage/db";

/** 텍스트 입력 요소 안이면 true — 전역 단축키·클립보드 가로채기 금지 */
const isTextInput = (target: EventTarget | null): boolean => {
  const el = target as HTMLElement | null;
  return (
    !!el &&
    (el.tagName === "INPUT" || el.tagName === "TEXTAREA" || el.isContentEditable)
  );
};

/** §4.7 반응형 구간 */
type Tier = "sm" | "md" | "lg" | "xl";
function useTier(): Tier {
  const [tier, setTier] = useState<Tier>("xl");
  useEffect(() => {
    const compute = () => {
      const w = window.innerWidth;
      setTier(w < 640 ? "sm" : w < 1024 ? "md" : w < 1280 ? "lg" : "xl");
    };
    compute();
    window.addEventListener("resize", compute);
    return () => window.removeEventListener("resize", compute);
  }, []);
  return tier;
}

type MobileView = "grid" | "python" | "toc" | "ai";

/** 접힌 패널 자리의 얇은 세로 스트립 — 클릭하면 다시 펼쳐진다 */
function CollapsedStrip({ panel, label }: { panel: "grid" | "python"; label: string }) {
  return (
    <button
      data-testid={`strip-${panel}`}
      aria-label={`${label} 패널 열기`}
      onClick={() => togglePanelCollapse(panel, false)}
      className={`flex w-7 shrink-0 flex-col items-center justify-start gap-1 bg-muted/40 py-2 text-xs text-muted-foreground hover:bg-muted hover:text-foreground ${
        panel === "grid" ? "border-r" : "border-l"
      }`}
    >
      {panel === "grid" ? <CaretRight className="size-3" /> : <CaretLeft className="size-3" />}
      <span className="[writing-mode:vertical-rl] tracking-wide">{label}</span>
    </button>
  );
}

export default function WorkbookShell() {
  const saveStatus = useAutosave();
  const [restored, setRestored] = useState(false);
  // 부록 P.6: 그리드 % — 기본 40(스프레드시트는 보조, Python 작업 중심)
  const [splitRatio, setSplitRatio] = useState(40);
  // 런타임 싱글턴 — 첫 클라이언트 렌더에서 생성 (ssr:false 페이지)
  const [runtime] = useState(() => getRuntimeClient());
  // §4.7 반응형: lg=Python 패널 접이식, md/sm=탭 전환
  const tier = useTier();
  const tierRef = useRef(tier);
  tierRef.current = tier;
  const [mobileView, setMobileView] = useState<MobileView>("grid");
  // 패널 접힘 — lg·xl 공통(스토어 + 설정 저장). 둘 다 접히지 않도록 스토어가 보장한다
  const gridCollapsed = useWorkbookStore((s) => s.gridCollapsed);
  const pyCollapsed = useWorkbookStore((s) => s.pyCollapsed);
  const tocOpen = useWorkbookStore((s) => s.tocOpen);
  // 부록 E R2: 뷰 전환 — 두 뷰 모두 마운트 유지(런타임·상태 보존), 비활성은 hidden
  const view = useWorkbookStore((s) => s.view);
  // 참조 뷰는 처음 활성화될 때 마운트하고 이후에는 hidden으로만 감춘다(상태 보존)
  const [refMounted, setRefMounted] = useState(false);
  useEffect(() => {
    if (view === "reference") setRefMounted(true);
  }, [view]);
  const aiChatOpen = useWorkbookStore((s) => s.aiChatOpen);
  // 부록 P: 속성 창 — 고정(옆에 붙어 화면을 나눔) / 겹침(화면 위)
  const propsOpen = useWorkbookStore((s) => s.propsOpen);
  const propsPinned = useWorkbookStore((s) => s.propsPinned);
  const gridMaximized = useWorkbookStore((s) => s.gridMaximized);
  const [propsWidth, setPropsWidth] = useState(360);
  const [propsResizing, setPropsResizing] = useState(false);
  const togglePropsPin = () => {
    const next = !useWorkbookStore.getState().propsPinned;
    useWorkbookStore.getState().setPropsPinned(next);
    void saveSettings({ propsPinned: next });
  };
  /** 왼쪽 가장자리 드래그로 너비 조절 (240~720px, 놓을 때 저장) */
  const startPropsResize = (e: React.PointerEvent) => {
    e.preventDefault();
    const x0 = e.clientX;
    const w0 = propsWidth;
    let w = w0;
    setPropsResizing(true); // 끄는 동안은 너비 애니메이션을 끈다
    const move = (ev: PointerEvent) => {
      w = Math.max(240, Math.min(720, w0 + x0 - ev.clientX));
      setPropsWidth(w);
    };
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      setPropsResizing(false);
      void saveSettings({ propsWidth: w });
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  };
  const closeToc = () => {
    useWorkbookStore.getState().setTocOpen(false);
    void saveSettings({ tocOpen: false });
  };
  const closeAiChat = () => {
    useWorkbookStore.getState().setAiChatOpen(false);
    void saveSettings({ aiChatOpen: false });
  };
  // 분할 크기 — 목차·AI 채팅 패널이 열리면 그리드·Python이 비율대로 줄어든다 (합계 100%)
  const TOC_SIZE = 16;
  const AI_SIZE = 20;
  const rest = 100 - (tocOpen ? TOC_SIZE : 0) - (aiChatOpen ? AI_SIZE : 0);
  const gridSize = pyCollapsed ? rest : Math.round((splitRatio / 100) * rest);
  const pySize = gridCollapsed ? rest : rest - gridSize;

  // 부록 P: 워크북 작업 폴더를 부트마다 os.chdir 코드로 적용
  useEffect(() => startWorkDirSync(runtime), [runtime]);

  // 런타임 백그라운드 부트 (멱등) — 첫 페인트와 CDN 다운로드가 경쟁하지 않게 유휴 시점으로 미룬다
  useEffect(() => {
    const start = () => void runtime.boot({ initScript: DEFAULT_INIT_SCRIPT });
    if ("requestIdleCallback" in window) requestIdleCallback(start, { timeout: 3000 });
    else setTimeout(start, 1500);
  }, [runtime]);

  // Ctrl+Z/Y·Ctrl+Shift+P + 패널 포커스 이동 Ctrl(또는 Alt)+1/2/3 (§1.6 접근성)
  useEffect(() => {
    const focusGrid = () => {
      const st = useWorkbookStore.getState();
      if (!st.selection) st.setSelection({ r0: 0, c0: 0, r1: 0, c1: 0 }); // 키보드 시작점
      if (tierRef.current === "md" || tierRef.current === "sm") setMobileView("grid");
      togglePanelCollapse("grid", false); // 접혀 있으면 펼치고 포커스
      requestAnimationFrame(() =>
        document.querySelector<HTMLElement>('[data-testid="data-grid-canvas"]')?.focus(),
      );
    };
    const focusPython = () => {
      useWorkbookStore.getState().setGridMaximized(false); // Python 작업 → 스프레드시트는 다시 보조 크기
      if (tierRef.current === "md" || tierRef.current === "sm") setMobileView("python");
      togglePanelCollapse("python", false);
      const st = useWorkbookStore.getState();
      const target = st.lastEditorBlockId ?? st.workbook.pyBlocks[0]?.id;
      if (target) requestAnimationFrame(() => useWorkbookStore.getState().setFocusBlock(target));
    };
    // 부록 P.7: 예전 하단 패널 자리 = 속성 창 — 열고 위쪽 활성 탭에 포커스
    const focusBottom = () => {
      toggleProps(true);
      requestAnimationFrame(() =>
        document
          .querySelector<HTMLElement>('[data-testid="properties-panel"] [role="tab"][aria-selected="true"]')
          ?.focus(),
      );
    };
    const onKeyDown = (e: KeyboardEvent) => {
      // 참조 뷰에서는 워크북 단축키(블록 추가·undo·패널 포커스)가 동작하지 않는다 (부록 E R2)
      if (useWorkbookStore.getState().view === "reference") return;
      // 출력 위치 지정 취소 (§ 앵커 재지정)
      if (e.key === "Escape" && useWorkbookStore.getState().anchorPicking) {
        useWorkbookStore.getState().setAnchorPicking(null);
        return;
      }
      // 부록 P.6: Esc — 스프레드시트 전체 화면 끝 (셀 편집 중이면 편집 취소가 먼저)
      if (e.key === "Escape" && useWorkbookStore.getState().gridMaximized && !isTextInput(e.target)) {
        useWorkbookStore.getState().setGridMaximized(false);
        return;
      }
      // Ctrl+Alt+3 — 속성 창 열기/닫기
      if ((e.ctrlKey || e.metaKey) && e.altKey && e.key === "3") {
        e.preventDefault();
        toggleProps();
        return;
      }
      // Ctrl+Alt+1/2 — 스프레드시트·Python 패널 접기/펼치기 (포커스 이동 단축키보다 먼저 판정)
      if ((e.ctrlKey || e.metaKey) && e.altKey && (e.key === "1" || e.key === "2")) {
        e.preventDefault();
        togglePanelCollapse(e.key === "1" ? "grid" : "python");
        return;
      }
      // 패널 포커스 이동은 텍스트 입력 중에도 동작 (Ctrl+숫자는 브라우저 탭 예약이라 Alt+숫자 병용)
      if ((e.ctrlKey || e.metaKey || e.altKey) && ["1", "2", "3"].includes(e.key)) {
        e.preventDefault();
        if (e.key === "1") focusGrid();
        else if (e.key === "2") focusPython();
        else focusBottom();
        return;
      }
      if (!(e.ctrlKey || e.metaKey)) return;
      if (isTextInput(e.target)) return; // 셀 편집기·제목 입력 등에서는 네이티브 텍스트 undo
      const key = e.key.toLowerCase();
      if (key === "p" && e.shiftKey) {
        e.preventDefault();
        addBlockAtSelection(); // ＋ Python 블록 (§2.3.1)
        return;
      }
      // 부록 O.3: 채우기 Ctrl+D(아래)·Ctrl+R(오른쪽) — 브라우저 북마크·새로고침보다 우선
      if ((key === "d" || key === "r") && !e.shiftKey && !e.altKey) {
        e.preventDefault();
        fillSelection(key === "d" ? "down" : "right");
        return;
      }
      if (key !== "z" && key !== "y") return;
      e.preventDefault();
      const temporal = useWorkbookStore.temporal.getState();
      if (key === "y" || e.shiftKey) temporal.redo();
      else temporal.undo();
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, []);

  // 그리드 포커스 상태의 붙여넣기/복사 — glide 내장 copy/paste는 SheetGrid에서 꺼 둠
  useEffect(() => {
    const onPaste = (e: ClipboardEvent) => {
      if (useWorkbookStore.getState().view === "reference") return; // 참조 뷰: 그리드 붙여넣기 금지
      if (isTextInput(e.target) || !e.clipboardData) return;
      const html = e.clipboardData.getData("text/html") || undefined;
      const text = e.clipboardData.getData("text/plain") || undefined;
      if (!html && !text) return;
      e.preventDefault();
      // 부록 O.3: 이 앱에서 복사한 범위면 수식째 붙인다 (상대 참조 이동)
      const internal = internalCopyFor(text);
      if (internal) {
        applyInternalPaste(internal);
        return;
      }
      void startPasteFlow(parseClipboard({ html, text }));
    };
    const onCopy = (e: ClipboardEvent) => {
      const cut = e.type === "cut"; // 부록 O.4: 잘라내기 — 값은 클립보드로, 이동은 붙여넣을 때
      if (useWorkbookStore.getState().view === "reference") return; // 참조 뷰: 페이지 텍스트 복사에 양보
      if (isTextInput(e.target) || !e.clipboardData) return;
      const domSelection = window.getSelection();
      if (domSelection && !domSelection.isCollapsed) return; // 페이지 텍스트 선택 중 → 기본 복사에 양보
      const { selection, workbook, activeSheetId } = useWorkbookStore.getState();
      if (!selection) return;
      const sheet = workbook.sheets.find((s) => s.id === activeSheetId);
      if (!sheet) return;
      const { text, html } = serializeRange(sheet, selection);
      e.clipboardData.setData("text/plain", text);
      e.clipboardData.setData("text/html", html);
      rememberCopy(sheet, selection, text, cut);
      e.preventDefault();
    };
    window.addEventListener("paste", onPaste);
    window.addEventListener("copy", onCopy);
    window.addEventListener("cut", onCopy);
    return () => {
      window.removeEventListener("paste", onPaste);
      window.removeEventListener("copy", onCopy);
      window.removeEventListener("cut", onCopy);
    };
  }, []);

  // 마운트 시: 설정 + 마지막 워크북 복원. 없거나 실패하면 임금 회귀 예측 샘플 (§2.2 첫 방문)
  useEffect(() => {
    (async () => {
      try {
        const settings = await loadSettings();
        if (settings?.gridSplit) setSplitRatio(settings.gridSplit);
        if (settings?.gridCompact === false) useWorkbookStore.getState().setGridCompact(false);
        if (settings?.tocOpen) useWorkbookStore.getState().setTocOpen(true);
        if (settings?.aiChatOpen) useWorkbookStore.getState().setAiChatOpen(true);
        if (settings?.propsOpen) useWorkbookStore.getState().setPropsOpen(true);
        if (settings?.propsPinned === false) useWorkbookStore.getState().setPropsPinned(false);
        if (settings?.propsWidth) setPropsWidth(settings.propsWidth);
        if (settings?.showRefs === false) useWorkbookStore.getState().setShowRefs(false);
        if (settings?.view === "reference") useWorkbookStore.getState().setView("reference");
        const wb = settings?.lastWorkbookId
          ? await getWorkbook(settings.lastWorkbookId)
          : undefined;
        if (wb) useWorkbookStore.getState().loadWorkbook(wb);
        else await loadDefaultWorkbook();
      } catch {
        await loadDefaultWorkbook().catch(() => undefined);
      } finally {
        setRestored(true);
        // e2e 테스트가 복원 완료를 기다릴 수 있게 신호
        (window as unknown as { __pygridReady?: boolean }).__pygridReady = true;
      }
    })();
  }, []);

  // 그리드 영역 드래그 앤 드롭 열기 (§1.5)
  const [dropActive, setDropActive] = useState(false);
  const dropHandlers = {
    onDragOver: (e: React.DragEvent) => {
      if (e.dataTransfer.types.includes("Files")) {
        e.preventDefault();
        setDropActive(true);
      }
    },
    onDragLeave: () => setDropActive(false),
    onDrop: (e: React.DragEvent) => {
      e.preventDefault();
      setDropActive(false);
      const file = e.dataTransfer.files?.[0];
      if (file) void openWorkbookFile(file);
    },
  };

  // 드래그 종료 후 한 번 호출됨 — 분할 비율·하단 패널 높이를 설정 스토어에 보존 (§3.2)
  // splitRatio는 그리드:Python 비율이다 — 목차 패널이 열려 있어도 같은 뜻이 되도록 정규화한다
  const onLayoutChanged = (layout: Record<string, number>) => {
    if (layout.grid && layout.python) {
      void saveSettings({
        gridSplit: Math.round((layout.grid / (layout.grid + layout.python)) * 100),
      });
    }
  };

  return (
    <TooltipProvider delayDuration={300}>
      <div className="flex h-screen flex-col bg-background">
        <Header>
          <RuntimeStatus client={runtime} />
        </Header>
        {/* 워크북 뷰 — 참조 뷰 활성 시에도 마운트 유지(hidden): 런타임·그리드 상태 보존 */}
        <div className={view === "workbook" ? "flex min-h-0 flex-1 flex-col" : "hidden"}>
        <ShellBar />
        <main className="relative flex min-h-0 flex-1">
        <div className="flex min-h-0 min-w-0 flex-1 flex-col">
        {tier === "md" || tier === "sm" ? (
          /* §4.7 640–1023(및 <640 열람 우선): 그리드 ↔ Python ↔ 결과 탭 전환 */
          <div className="flex min-h-0 flex-1 flex-col">
            <div role="tablist" aria-label="화면 전환" className="flex shrink-0 border-b bg-muted/40">
              {(
                [
                  ["grid", "그리드"],
                  ["python", "Python"],
                  ["toc", "목차"],
                  ["ai", "AI 채팅"],
                ] as const
              ).map(([id, label]) => (
                <button
                  key={id}
                  role="tab"
                  aria-selected={mobileView === id}
                  onClick={() => setMobileView(id)}
                  className={`h-8 border-b-2 px-4 text-xs ${
                    mobileView === id
                      ? "border-primary font-medium text-foreground" // 대비 4.5:1 — 인디케이터만 primary
                      : "border-transparent text-muted-foreground"
                  }`}
                >
                  {label}
                </button>
              ))}
            </div>
            <div className="min-h-0 flex-1">
              {mobileView === "grid" && (
                <div
                  {...dropHandlers}
                  className={`flex h-full min-w-0 flex-col ${dropActive ? "ring-2 ring-inset ring-primary" : ""}`}
                >
                  <SheetEditToolbar />
                  <FormulaBar />
                  <SheetGrid />
                  <SheetTabs />
                </div>
              )}
              {mobileView === "python" && <PythonPanel />}
              {mobileView === "toc" && <TocPanel />}
              {mobileView === "ai" && <AiChatPanel />}

            </div>
          </div>
        ) : gridMaximized ? (
          /* 부록 P.6: 스프레드시트 전체 화면 — Esc·오른쪽 아래 버튼·Python 포커스(Ctrl+2)로 돌아온다 */
          <div
            {...dropHandlers}
            data-testid="grid-maximized"
            className={`flex min-h-0 flex-1 flex-col ${dropActive ? "ring-2 ring-inset ring-primary" : ""}`}
          >
            <SheetEditToolbar />
            <FormulaBar />
            <SheetGrid />
            <SheetTabs />
          </div>
        ) : (
          /* 부록 P.7: 하단 패널 없음 — 진단·미리보기·콘솔은 오른쪽 속성 창 탭.
             접힌 패널은 그룹 밖 세로 스트립으로 대체 — Panel은 Group의 직계 자식이어야 한다 */
          <div className="flex min-h-0 flex-1">
            {gridCollapsed && <CollapsedStrip panel="grid" label="시트" />}
            <ResizablePanelGroup
              key={`${restored ? "r" : "i"}-${gridCollapsed ? "no-grid" : "grid"}-${pyCollapsed ? "no-py" : "py"}-${tocOpen ? "toc" : "no-toc"}-${aiChatOpen ? "ai" : "no-ai"}`}
              orientation="horizontal"
              className="min-h-0 min-w-0 flex-1"
              onLayoutChanged={onLayoutChanged}
            >
              {!gridCollapsed && (
                <ResizablePanel id="grid" defaultSize={`${gridSize}%`} minSize="15%">
                  <div
                    {...dropHandlers}
                    className={`flex h-full min-w-0 flex-col ${dropActive ? "ring-2 ring-inset ring-primary" : ""}`}
                  >
                    <SheetEditToolbar />
                    <FormulaBar />
                    <SheetGrid />
                    <SheetTabs />
                  </div>
                </ResizablePanel>
              )}
              {!pyCollapsed && (
                <>
                  {!gridCollapsed && <ResizableHandle withHandle />}
                  <ResizablePanel
                    id="python"
                    defaultSize={`${pySize}%`}
                    minSize="15%"
                  >
                    <PythonPanel />
                  </ResizablePanel>
                </>
              )}
              {/* 부록 D.2: 목차 전용 패널 (툴바 토글·자체 ✕, 상태는 설정에 저장) */}
              {tocOpen && (
                <>
                  <ResizableHandle withHandle />
                  <ResizablePanel id="toc" defaultSize={`${TOC_SIZE}%`} minSize="10%">
                    <TocPanel onClose={closeToc} />
                  </ResizablePanel>
                </>
              )}
              {/* 부록 G.2: AI 채팅 패널 (최우측, 같은 패턴) */}
              {aiChatOpen && (
                <>
                  <ResizableHandle withHandle />
                  <ResizablePanel id="aichat" defaultSize={`${AI_SIZE}%`} minSize="12%">
                    <AiChatPanel onClose={closeAiChat} />
                  </ResizablePanel>
                </>
              )}
            </ResizablePanelGroup>
            {pyCollapsed && <CollapsedStrip panel="python" label="Python" />}
          </div>
        )}
        </div>
        {/* 부록 P·P.7: 속성 창 — 항상 마운트(콘솔 기록 보존)하고 오른쪽에서 밀려 나오고 들어간다.
            고정: 옆 칸 너비가 0 ↔ W로 변하며 화면을 나눈다 / 겹침(좁은 화면은 항상): 화면 위로 미끄러져 덮는다 */}
        {(() => {
          const overlay = !propsPinned || tier === "md" || tier === "sm";
          return (
            <aside
              aria-label="속성 창"
              aria-hidden={!propsOpen}
              inert={!propsOpen}
              data-open={propsOpen}
              style={{ width: overlay ? `min(${propsWidth}px, 100%)` : propsOpen ? propsWidth : 0 }}
              className={[
                overlay ? "absolute inset-y-0 right-0 z-30" : "relative shrink-0",
                "overflow-hidden bg-background",
                propsOpen ? "border-l" : "",
                overlay && propsOpen ? "shadow-2xl" : "",
                overlay && !propsOpen ? "translate-x-full" : "translate-x-0",
                propsResizing ? "" : "transition-[width,transform] duration-200 ease-out motion-reduce:transition-none",
              ].join(" ")}
            >
              <div
                role="separator"
                aria-orientation="vertical"
                aria-label="속성 창 너비 조절"
                onPointerDown={startPropsResize}
                className="absolute inset-y-0 left-0 z-10 w-1.5 cursor-col-resize hover:bg-primary/20"
              />
              <div className="h-full" style={{ width: overlay ? "100%" : propsWidth }}>
                <PropertiesPanel
                  open={propsOpen}
                  pinned={!overlay}
                  onTogglePin={tier === "md" || tier === "sm" ? undefined : togglePropsPin}
                  onClose={() => toggleProps(false)}
                />
              </div>
            </aside>
          );
        })()}
        </main>
        </div>
        {/* 참조 뷰(데이터 예제/분석) — 항상 마운트, 비활성 시 hidden (부록 E R2) */}
        <div
          data-testid="reference-view"
          className={view === "reference" ? "min-h-0 flex-1 overflow-hidden" : "hidden"}
        >
          {refMounted ? <ReferenceView /> : null}
        </div>
        <StatusBar saveStatus={saveStatus} />
        <PasteImportDialog />
        <FitGuideDialog />
      </div>
    </TooltipProvider>
  );
}
