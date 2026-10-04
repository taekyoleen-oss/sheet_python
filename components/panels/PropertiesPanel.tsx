"use client";

// 부록 P: 속성 창 (RStudio 오른쪽 창) — 오른쪽에서 열고 닫으며, 화면에 고정하거나 위에 겹친다.
//  탭: 위 [변수 · 콘솔 · 진단], 아래 [파일 · 출력 미리보기] — 하단 패널을 없애 코딩 공간을 넓혔다(P.7).
//  위: 변수 — 실행으로 만든 전역 변수(이름·타입·크기). 클릭 = 세부 팝업, 시트에 보이기·모델 결과·이름 복사
//  아래: 파일 — 탐색기처럼 폴더 이동, 작업 폴더 지정(os.chdir 코드), 파일 불러오기(코드 블록 생성)

import { useCallback, useEffect, useRef, useState } from "react";
import {
  ArrowsClockwise,
  ArrowUp,
  ChartLineUp,
  Copy,
  DotsThreeVertical,
  FileText,
  FolderOpen,
  FolderPlus,
  FolderSimple,
  House,
  PushPin,
  PushPinSlash,
  Table,
  UploadSimple,
  X,
} from "@phosphor-icons/react";
import { toast } from "sonner";
import { ConsoleTab } from "@/components/panels/ConsoleTab";
import DiagnosticsTab from "@/components/panels/DiagnosticsTab";
import OutputPreviewTab from "@/components/panels/OutputPreviewTab";
import ShowVariableDialog from "@/components/panels/ShowVariableDialog";
import VariableDetailDialog from "@/components/panels/VariableDetailDialog";
import ModelResultDialog from "@/components/python/ModelResultDialog";
import { copyText } from "@/components/reference/code-popup";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { ResizableHandle, ResizablePanel, ResizablePanelGroup } from "@/components/ui/resizable";
import {
  breadcrumbs,
  fmtSize,
  DEFAULT_WORK_DIR,
  HOME_DIR,
  joinPath,
  loadCode,
  parentPath,
  workDirCode,
} from "@/lib/grid/files";
import { insertSnippetAsBlock } from "@/lib/grid/insert-snippet";
import { blocksInOrder, useWorkbookStore } from "@/lib/grid/model";
import { definingBlock } from "@/lib/grid/model-output";
import { outputsOf } from "@/lib/grid/outputs";
import { getRuntimeClient, type RuntimeClient } from "@/lib/runtime/client";
import type { DirEntry, VariableInfo } from "@/lib/runtime/protocol";
import { applyWorkDir } from "@/lib/runtime/workdir-sync";
import { cn } from "@/lib/utils";
import type { PyBlock } from "@/types/workbook";

const store = () => useWorkbookStore.getState();

const copy = async (text: string, what = "복사했습니다") => {
  if (await copyText(text)) toast.success(what);
};

/**
 * 보이는 동안만 런타임을 조회한다 — 보이기 시작할 때 한 번, 그 뒤 실행 완료(ready)마다.
 * 창이 닫혀 있거나 다른 탭이면 inspect·listDir을 보내지 않는다(모델 카탈로그 계산 비용).
 */
function useOnReady(client: RuntimeClient, active: boolean, fn: () => void) {
  const ref = useRef(fn);
  ref.current = fn;
  useEffect(() => {
    if (!active) return;
    ref.current();
    return client.on("status", (s) => {
      if (s === "ready") ref.current();
    });
  }, [client, active]);
}

// ── 변수 ────────────────────────────────────────────────

function sizeText(v: VariableInfo): string {
  if (v.shape) return `${v.shape[0]}×${v.shape[1]}`;
  return "";
}

function EnvironmentSection({ client, active }: { client: RuntimeClient; active: boolean }) {
  const [vars, setVars] = useState<VariableInfo[]>([]);
  const [detail, setDetail] = useState<VariableInfo | null>(null);
  const [show, setShow] = useState<VariableInfo | null>(null);
  const [modelBlock, setModelBlock] = useState<PyBlock | null>(null);

  const refresh = useCallback(async () => {
    if (client.getStatus() !== "ready") return;
    try {
      setVars(await client.inspect());
    } catch {
      /* 재부트 중 — 다음 ready에서 회복 */
    }
  }, [client]);
  useOnReady(client, active, () => void refresh());

  const openModel = (v: VariableInfo) => {
    const b = definingBlock(store().workbook, v.name);
    setDetail(null);
    if (b) setModelBlock(b);
    else toast.error("모델을 만든 코드 블록을 찾을 수 없습니다");
  };

  return (
    <div className="flex h-full flex-col">
      <div className="flex h-8 shrink-0 items-center gap-1 border-b bg-muted/40 px-2">
        <span className="text-xs font-semibold">변수</span>
        <span className="text-xs text-muted-foreground">실행으로 만든 {vars.length}개</span>
        <button
          className="ml-auto rounded p-1 hover:bg-accent"
          aria-label="변수 목록 새로고침"
          title="새로고침"
          onClick={() => void refresh()}
        >
          <ArrowsClockwise className="size-3.5" />
        </button>
      </div>
      <div className="min-h-0 flex-1 overflow-y-auto text-xs" data-testid="props-variables">
        {vars.length === 0 ? (
          <p className="px-3 py-4 text-center text-muted-foreground">
            {client.getStatus() === "ready"
              ? "아직 실행으로 만든 변수가 없습니다 — 블록을 실행하거나 아래 파일을 불러오세요"
              : "런타임 준비 후 표시됩니다"}
          </p>
        ) : (
          <table className="w-full border-collapse">
            <thead className="sticky top-0 bg-background text-left text-[11px] text-muted-foreground">
              <tr className="border-b">
                <th className="px-2 py-1 font-normal">이름</th>
                <th className="px-1 py-1 font-normal">타입</th>
                <th className="px-1 py-1 font-normal">크기</th>
                <th className="w-6" />
              </tr>
            </thead>
            <tbody>
              {vars.map((v) => (
                <VarRow
                  key={v.name}
                  info={v}
                  onDetail={() => setDetail(v)}
                  onShow={() => setShow(v)}
                  onModel={() => openModel(v)}
                />
              ))}
            </tbody>
          </table>
        )}
      </div>
      <VariableDetailDialog
        info={detail}
        onClose={() => setDetail(null)}
        onShow={(v) => {
          setDetail(null);
          setShow(v);
        }}
        onModel={openModel}
      />
      <ShowVariableDialog info={show} onClose={() => setShow(null)} />
      {modelBlock && (
        <ModelResultDialog block={modelBlock} open onOpenChange={(o) => !o && setModelBlock(null)} />
      )}
    </div>
  );
}

function VarRow({
  info,
  onDetail,
  onShow,
  onModel,
}: {
  info: VariableInfo;
  onDetail: () => void;
  onShow: () => void;
  onModel: () => void;
}) {
  return (
    <tr
      className="group cursor-pointer border-b hover:bg-accent/40"
      onClick={onDetail}
      title={`${info.summary ?? ""}
(클릭: 세부 내용)`}
      data-var={info.name}
    >
      <td className="max-w-28 truncate px-2 py-1 font-mono font-medium">{info.name}</td>
      <td className="max-w-24 truncate px-1 py-1 font-mono text-muted-foreground">
        {info.model ? <span className="text-primary">{info.type}</span> : info.type}
      </td>
      <td className="px-1 py-1 font-mono text-muted-foreground">{sizeText(info)}</td>
      <td className="px-1 py-0.5 text-right" onClick={(e) => e.stopPropagation()}>
        <DropdownMenu>
          <DropdownMenuTrigger
            aria-label={`${info.name} 작업`}
            className="rounded p-0.5 opacity-60 hover:bg-accent group-hover:opacity-100"
          >
            <DotsThreeVertical className="size-3.5" />
          </DropdownMenuTrigger>
          <DropdownMenuContent align="end" className="text-xs">
            <DropdownMenuItem onSelect={onDetail}>세부 내용 보기</DropdownMenuItem>
            <DropdownMenuItem onSelect={onShow}>
              <Table className="size-3.5" /> 스프레드시트에 보이기…
            </DropdownMenuItem>
            {info.model && (
              <DropdownMenuItem onSelect={onModel}>
                <ChartLineUp className="size-3.5" /> 모델 결과 → 시트…
              </DropdownMenuItem>
            )}
            <DropdownMenuSeparator />
            <DropdownMenuItem onSelect={() => void copy(info.name, "이름을 복사했습니다")}>
              <Copy className="size-3.5" /> 이름 복사
            </DropdownMenuItem>
          </DropdownMenuContent>
        </DropdownMenu>
      </td>
    </tr>
  );
}

// ── 파일 ────────────────────────────────────────────────

/** 불러오기 코드를 새 블록으로 — 기준은 마지막으로 편집한 블록(없으면 계산 순서 마지막) 아래 */
function loadIntoBlock(filePath: string, cwd: string) {
  const res = loadCode(filePath, cwd);
  if (!res) {
    toast.error("이 형식은 불러오기 코드를 만들 수 없습니다 (csv·xlsx·json·parquet·pkl·txt)");
    return;
  }
  const st = store();
  const ref = st.lastEditorBlockId ?? blocksInOrder(st.workbook).filter((b) => b.kind !== "markdown").pop()?.id ?? null;
  const name = filePath.split("/").pop();
  const made = insertSnippetAsBlock(ref, "below", `데이터 불러오기: ${name}`, res.code);
  if (!made) return;
  st.setFocusBlock(made.id);
  toast.success(`불러오기 블록을 추가했습니다 — ▶ 실행하면 변수 ${res.varName}가 만들어지고 미리보기가 블록 아래에 보입니다`);
}

function FilesSection({ client, active }: { client: RuntimeClient; active: boolean }) {
  const workDir = useWorkbookStore((s) => s.workbook.workDir);
  const [cwd, setCwd] = useState<string>(workDir ?? DEFAULT_WORK_DIR);
  const [path, setPath] = useState<string | null>(null); // null = 작업 폴더
  const [entries, setEntries] = useState<DirEntry[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [selected, setSelected] = useState<string | null>(null);
  const [folderName, setFolderName] = useState<string | null>(null); // null = 입력 줄 닫힘
  const uploadRef = useRef<HTMLInputElement>(null);

  const list = useCallback(
    async (target: string | null) => {
      try {
        const res = await client.listDir(target ?? undefined);
        setCwd(res.cwd);
        setPath(res.path);
        setEntries(res.entries);
        setError(null);
      } catch (e) {
        setError((e as Error).message);
      }
    },
    [client],
  );
  const pathRef = useRef(path);
  pathRef.current = path;
  useOnReady(client, active, () => void list(pathRef.current));
  // 작업 폴더가 바뀌면 그 폴더로 이동
  useEffect(() => {
    if (active) void list(null);
  }, [workDir, list, active]);

  const here = path ?? cwd;

  const setAsWorkDir = async (dir: string) => {
    store().setWorkDir(dir === DEFAULT_WORK_DIR ? null : dir);
    const err = await applyWorkDir(client, dir);
    if (err) toast.error(`작업 폴더 지정 실패: ${err}`);
    else toast.success(`작업 폴더: ${dir} — 코드의 상대 경로가 이 폴더 기준이 됩니다`);
    void list(dir);
  };

  const newFolder = async (name: string) => {
    setFolderName(null);
    if (!name.trim() || name.includes("/")) return;
    const dir = joinPath(here, name.trim());
    try {
      const res = await client.repl(`import os\nos.makedirs(${JSON.stringify(dir)}, exist_ok=True)`);
      if (res.traceback) toast.error(`폴더를 만들 수 없습니다: ${res.traceback.trim().split("\n").pop()}`);
    } catch (e) {
      toast.error(`폴더를 만들 수 없습니다: ${(e as Error).message}`);
    }
    void list(here);
  };

  const upload = async (files: FileList | null) => {
    if (!files) return;
    for (const f of Array.from(files)) {
      try {
        // writeFile은 작업 폴더에 쓴다 (재부트 뒤에도 클라이언트 사본으로 복원)
        await client.writeFile(f.name, new Uint8Array(await f.arrayBuffer()));
      } catch (e) {
        toast.error(`${f.name}: ${(e as Error).message}`);
      }
    }
    void list(null);
  };

  return (
    <div className="flex h-full flex-col">
      <div className="flex h-8 shrink-0 items-center gap-0.5 border-b bg-muted/40 px-2">
        <span className="mr-1 text-xs font-semibold">파일</span>
        <IconBtn label="사용자 폴더 (tklee)" onClick={() => void list(HOME_DIR)}>
          <House className="size-3.5" />
        </IconBtn>
        <IconBtn label="상위 폴더" onClick={() => void list(parentPath(here))} disabled={here === "/" || here === HOME_DIR}>
          <ArrowUp className="size-3.5" />
        </IconBtn>
        <IconBtn label="작업 폴더로 이동" onClick={() => void list(null)}>
          <FolderOpen className="size-3.5" />
        </IconBtn>
        <IconBtn label="새 폴더" onClick={() => setFolderName("")}>
          <FolderPlus className="size-3.5" />
        </IconBtn>
        <IconBtn label="작업 폴더에 파일 올리기" onClick={() => uploadRef.current?.click()}>
          <UploadSimple className="size-3.5" />
        </IconBtn>
        <IconBtn label="새로고침" onClick={() => void list(here)}>
          <ArrowsClockwise className="size-3.5" />
        </IconBtn>
        <input
          ref={uploadRef}
          type="file"
          multiple
          hidden
          onChange={(e) => {
            void upload(e.target.files);
            e.target.value = "";
          }}
        />
      </div>

      {/* 작업 폴더 + 그 코드 */}
      <div className="shrink-0 space-y-1 border-b px-2 py-1.5 text-[11px]">
        <div className="flex items-center gap-1">
          <span className="text-muted-foreground">작업 폴더</span>
          <code className="min-w-0 flex-1 truncate font-mono text-primary" title={cwd} data-testid="props-cwd">
            {cwd}
          </code>
          <IconBtn label="작업 폴더 지정 코드 복사" onClick={() => void copy(workDirCode(cwd), "코드를 복사했습니다")}>
            <Copy className="size-3" />
          </IconBtn>
        </div>
        <div className="flex flex-wrap items-center gap-0.5 font-mono">
          {breadcrumbs(here, HOME_DIR).map((b, i, arr) => (
            <span key={b.path} className="flex items-center">
              <button
                className={cn("rounded px-0.5 hover:bg-accent", i === arr.length - 1 && "font-semibold")}
                onClick={() => void list(b.path)}
              >
                {b.name}
              </button>
              {i > 0 && i < arr.length - 1 && <span className="text-muted-foreground">/</span>}
            </span>
          ))}
          {here !== cwd && (
            <button
              className="ml-auto rounded border border-primary px-1.5 py-0.5 text-[11px] text-primary hover:bg-primary/10"
              onClick={() => void setAsWorkDir(here)}
              title={workDirCode(here)}
            >
              이 폴더를 작업 폴더로
            </button>
          )}
        </div>
      </div>

      {folderName !== null && (
        <form
          className="flex shrink-0 items-center gap-1 border-b px-2 py-1"
          onSubmit={(e) => {
            e.preventDefault();
            void newFolder(folderName);
          }}
        >
          <FolderPlus className="size-3.5 shrink-0 text-muted-foreground" />
          <input
            autoFocus
            aria-label="새 폴더 이름"
            placeholder="새 폴더 이름 (Enter)"
            value={folderName}
            onChange={(e) => setFolderName(e.target.value)}
            onKeyDown={(e) => e.key === "Escape" && setFolderName(null)}
            className="h-6 min-w-0 flex-1 rounded border bg-background px-1.5 font-mono text-xs"
          />
        </form>
      )}
      <div className="min-h-0 flex-1 overflow-y-auto text-xs" data-testid="props-files">
        {error ? (
          <p className="px-3 py-3 text-destructive">{error}</p>
        ) : entries.length === 0 ? (
          <p className="px-3 py-4 text-center text-muted-foreground">
            빈 폴더입니다 — ⤒ 버튼으로 파일을 올리거나 파일 메뉴의 &lsquo;데이터 불러오기&rsquo;를 쓰세요
          </p>
        ) : (
          <ul>
            {entries.map((e) => {
              const full = joinPath(here, e.name);
              return (
                <li
                  key={e.name}
                  className={cn(
                    "group flex cursor-pointer items-center gap-1.5 px-2 py-1 hover:bg-accent/40",
                    selected === full && "bg-accent",
                  )}
                  onClick={() => (e.dir ? void list(full) : setSelected(full))}
                  onDoubleClick={() => !e.dir && loadIntoBlock(full, cwd)}
                  data-entry={e.name}
                >
                  {e.dir ? (
                    <FolderSimple weight="fill" className="size-3.5 shrink-0 text-amber-500" />
                  ) : (
                    <FileText className="size-3.5 shrink-0 text-muted-foreground" />
                  )}
                  <span className="min-w-0 flex-1 truncate font-mono">{e.name}</span>
                  {!e.dir && <span className="shrink-0 text-muted-foreground">{fmtSize(e.size)}</span>}
                  <span onClick={(ev) => ev.stopPropagation()}>
                    <DropdownMenu>
                      <DropdownMenuTrigger
                        aria-label={`${e.name} 작업`}
                        className="rounded p-0.5 opacity-60 hover:bg-accent group-hover:opacity-100"
                      >
                        <DotsThreeVertical className="size-3.5" />
                      </DropdownMenuTrigger>
                      <DropdownMenuContent align="end" className="text-xs">
                        {e.dir ? (
                          <>
                            <DropdownMenuItem onSelect={() => void list(full)}>열기</DropdownMenuItem>
                            <DropdownMenuItem onSelect={() => void setAsWorkDir(full)}>작업 폴더로 지정</DropdownMenuItem>
                          </>
                        ) : (
                          <DropdownMenuItem onSelect={() => loadIntoBlock(full, cwd)}>
                            불러오기 (Python 코드 블록 추가)
                          </DropdownMenuItem>
                        )}
                        <DropdownMenuSeparator />
                        <DropdownMenuItem onSelect={() => void copy(full, "경로를 복사했습니다")}>
                          <Copy className="size-3.5" /> 경로 복사
                        </DropdownMenuItem>
                      </DropdownMenuContent>
                    </DropdownMenu>
                  </span>
                </li>
              );
            })}
          </ul>
        )}
      </div>
      {selected && (
        <div className="flex shrink-0 items-center gap-1 border-t bg-muted/30 px-2 py-1 text-[11px]">
          <span className="min-w-0 flex-1 truncate font-mono">{selected.split("/").pop()}</span>
          <button
            className="rounded bg-primary px-2 py-0.5 text-primary-foreground hover:bg-primary/90"
            onClick={() => loadIntoBlock(selected, cwd)}
          >
            불러오기
          </button>
        </div>
      )}
    </div>
  );
}

function IconBtn({
  label,
  onClick,
  disabled,
  children,
}: {
  label: string;
  onClick: () => void;
  disabled?: boolean;
  children: React.ReactNode;
}) {
  return (
    <button
      type="button"
      aria-label={label}
      title={label}
      disabled={disabled}
      onClick={onClick}
      className="rounded p-1 hover:bg-accent disabled:opacity-40"
    >
      {children}
    </button>
  );
}

// ── 창 ──────────────────────────────────────────────────

const TAB =
  "h-7 border-b-2 px-2 text-xs data-[on=true]:border-primary data-[on=true]:font-medium data-[on=false]:border-transparent data-[on=false]:text-muted-foreground hover:text-foreground";

/** 탭 머리 — 위·아래 칸 공용 */
function TabBar<T extends string>({
  label,
  tabs,
  value,
  onChange,
}: {
  label: string;
  tabs: { id: T; name: string; badge?: number }[];
  value: T;
  onChange: (v: T) => void;
}) {
  return (
    <div role="tablist" aria-label={label} className="flex h-8 shrink-0 items-end gap-0.5 border-b bg-muted/40 px-1">
      {tabs.map((t) => (
        <button
          key={t.id}
          role="tab"
          aria-selected={value === t.id}
          data-on={value === t.id}
          className={TAB}
          onClick={() => onChange(t.id)}
        >
          {t.name}
          {t.badge ? (
            <span className="ml-1 rounded bg-destructive px-1 text-[10px] text-white">{t.badge}</span>
          ) : null}
        </button>
      ))}
    </div>
  );
}

type PanelTab = Parameters<ReturnType<typeof store>["showPanelTab"]>[0];

/**
 * 부록 P.7: 속성 창 = 오른쪽 작업대.
 *  위 칸 탭 [변수 · 콘솔 · 진단], 아래 칸 탭 [파일 · 출력 미리보기] — 예전 하단 패널을 모두 여기로 옮겼다.
 *  콘솔이 실행 출력을 놓치지 않도록 창은 닫혀 있어도 마운트된 채 숨는다(WorkbookShell).
 */
export default function PropertiesPanel({
  open = true,
  pinned,
  onTogglePin,
  onClose,
}: {
  open?: boolean;
  pinned: boolean;
  onTogglePin?: () => void;
  onClose: () => void;
}) {
  const [client] = useState(() => getRuntimeClient());
  const top = useWorkbookStore((s) => s.propsTopTab);
  const bottom = useWorkbookStore((s) => s.propsBottomTab);
  const errorCount = useWorkbookStore((s) =>
    s.workbook.pyBlocks
      .filter((b) => b.kind !== "markdown")
      .reduce(
        (n, b) => n + outputsOf(b).filter((o) => o.last?.status === "error" || o.last?.status === "spill").length,
        0,
      ),
  );
  const show = (t: PanelTab) => store().showPanelTab(t);

  return (
    <div className="flex h-full flex-col bg-background" data-testid="properties-panel">
      <div className="flex h-8 shrink-0 items-center gap-1 border-b px-2">
        <span className="text-xs font-semibold">속성</span>
        <span className="text-[11px] text-muted-foreground">Ctrl+Alt+3</span>
        {onTogglePin && (
          <button
            className="ml-auto rounded p-1 hover:bg-accent"
            onClick={onTogglePin}
            aria-label={pinned ? "화면 위에 겹치기" : "화면에 고정"}
            aria-pressed={pinned}
            title={pinned ? "고정됨 — 누르면 화면 위에 겹칩니다" : "겹침 — 누르면 옆에 고정됩니다"}
          >
            {pinned ? <PushPin weight="fill" className="size-3.5 text-primary" /> : <PushPinSlash className="size-3.5" />}
          </button>
        )}
        <button
          className={cn("rounded p-1 hover:bg-accent", !onTogglePin && "ml-auto")}
          onClick={onClose}
          aria-label="속성 창 닫기"
        >
          <X className="size-3.5" />
        </button>
      </div>
      <ResizablePanelGroup orientation="vertical" className="min-h-0 flex-1">
        <ResizablePanel id="props-top" defaultSize="55%" minSize="15%">
          <div className="flex h-full flex-col">
            <TabBar<PanelTab>
              label="속성 위쪽 탭"
              value={top}
              onChange={show}
              tabs={[
                { id: "variables", name: "변수" },
                { id: "console", name: "콘솔" },
                { id: "diagnostics", name: "진단", badge: errorCount },
              ]}
            />
            <div className="min-h-0 flex-1">
              <div className={top === "variables" ? "h-full" : "hidden"}>
                <EnvironmentSection client={client} active={open && top === "variables"} />
              </div>
              {/* 콘솔은 히스토리 보존을 위해 항상 마운트 */}
              <div className={top === "console" ? "h-full" : "hidden"}>
                <ConsoleTab client={client} className="h-full" />
              </div>
              {top === "diagnostics" && (
                <div className="h-full overflow-y-auto">
                  <DiagnosticsTab />
                </div>
              )}
            </div>
          </div>
        </ResizablePanel>
        <ResizableHandle withHandle />
        <ResizablePanel id="props-bottom" defaultSize="45%" minSize="15%">
          <div className="flex h-full flex-col">
            <TabBar<PanelTab>
              label="속성 아래쪽 탭"
              value={bottom}
              onChange={show}
              tabs={[
                { id: "files", name: "파일" },
                { id: "preview", name: "출력 미리보기" },
              ]}
            />
            <div className="min-h-0 flex-1">
              <div className={bottom === "files" ? "h-full" : "hidden"}>
                <FilesSection client={client} active={open && bottom === "files"} />
              </div>
              {bottom === "preview" && (
                <div className="h-full overflow-y-auto">
                  <OutputPreviewTab />
                </div>
              )}
            </div>
          </div>
        </ResizablePanel>
      </ResizablePanelGroup>
    </div>
  );
}
