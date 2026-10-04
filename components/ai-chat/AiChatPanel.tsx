"use client";

// 부록 G.2·G.3 + L — 워크북 에이전트 채팅 패널 (최우측, TocPanel 패턴: 툴바 토글 + 자체 ✕).
// 앱의 유일한 AI 진입점(L.5): 블록 ✦ 4액션·패널 상단 생성 바가 모두 여기로 들어온다.
// 이력·지침은 IndexedDB 앱 설정 로컬 전용(워크북 파일 미포함).
// 변경은 항상 제안 → [적용] 클릭(L.0). 자동 적용·자동 실행은 없다.

import { useCallback, useEffect, useRef, useState } from "react";
import { ArrowUp, Broom, Gear, PaperclipHorizontal, X } from "@phosphor-icons/react";
import { toast } from "sonner";
import { CopyButton } from "@/components/reference/code-popup";
import { getApiKey, openApiKeyDialog } from "@/components/shell/ApiKeyDialog";
import { editorRegistry } from "@/components/python/CodeEditor";
import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { runAgent, type RefChip, type ToolLog } from "@/lib/ai/agent";
import {
  buildChatSystem,
  capHistory,
  chatMessages,
  DEFAULT_CHAT_INSTRUCTIONS,
  parseInstructionFence,
  splitCodeBlocks,
  withAttachment,
  type ChatAsk,
  type ChatAttachment,
  type ChatMessage,
} from "@/lib/ai/chat";
import { sheetOverview } from "@/lib/ai/schema";
import { applyCellProposal, type BlockProposal, type CellProposal, type Proposal } from "@/lib/ai/tools";
import { notifyWorkbookEdit } from "@/lib/grid/calc-host";
import { codeTitle } from "@/lib/grid/code-sections";
import { appendSnippetToBlock, insertSnippetAsBlock } from "@/lib/grid/insert-snippet";
import { renderMarkdown } from "@/lib/grid/markdown";
import { useWorkbookStore } from "@/lib/grid/model";
import { loadSettings, saveSettings } from "@/lib/storage/db";
import { cn } from "@/lib/utils";

// 패널이 아직 마운트되기 전에 온 요청을 들고 있다가 마운트 시 소비한다
const ASK_EVENT = "pygrid:ai-chat-ask";
let pendingAsk: ChatAsk | null = null;

/** 어디서든 채팅 패널 열기 — question이 있으면 즉시 전송(블록 ✦ 액션·생성 바, L.5) */
export function askAiChat(ask: ChatAsk): void {
  pendingAsk = ask;
  useWorkbookStore.getState().setAiChatOpen(true);
  void saveSettings({ aiChatOpen: true });
  window.dispatchEvent(new Event(ASK_EVENT));
}

/** 그리드에서 근거 범위로 이동 + 하이라이트 유지 */
function gotoRef(ref: RefChip, all: RefChip[]): void {
  const st = useWorkbookStore.getState();
  st.setActiveSheet(ref.sheetId);
  st.setSelection({ r0: ref.r0, c0: ref.c0, r1: ref.r1, c1: ref.c1 });
  st.setChatRefs(all);
}

/** 대화 속 코드 카드 → 블록 반영 (부록 F.1 삽입 경로 재사용, 자동 실행 없음) */
function insertChatCode(
  code: string,
  placement: "append" | "below" | "above",
  opts: { targetBlockId?: string; title?: string; note?: string; replace?: boolean } = {},
): boolean {
  const st = useWorkbookStore.getState();
  const usable = (id: string | null | undefined): boolean =>
    !!id && st.workbook.pyBlocks.some((b) => b.id === id && b.kind !== "markdown");
  const refId = usable(opts.targetBlockId)
    ? (opts.targetBlockId as string)
    : usable(st.lastEditorBlockId)
      ? (st.lastEditorBlockId as string)
      : null;

  let targetId: string;
  if (placement === "append") {
    if (!refId) {
      toast("대상 블록이 없습니다 — 블록 편집기를 먼저 클릭하세요");
      return false;
    }
    if (opts.replace) st.setBlockCode(refId, code); // 제안은 블록 전체 코드 = 교체 (1 undo)
    else if (!appendSnippetToBlock(refId, code)) return false;
    notifyWorkbookEdit([], [refId]);
    targetId = refId;
  } else {
    const res = insertSnippetAsBlock(
      refId,
      placement,
      opts.title || codeTitle(code) || "AI 제안",
      code,
      opts.note,
    );
    if (!res) {
      toast.error("블록을 만들 수 없습니다 (활성 시트 없음)");
      return false;
    }
    if (refId && !res.ordered) {
      toast("순서를 보장할 빈 위치가 없어 빈 영역에 배치했습니다 — ↑↓로 순서를 조정하세요");
    }
    targetId = res.id;
  }
  st.setFocusBlock(targetId);
  st.setSelectedBlock(targetId);
  toast("블록에 반영했습니다 — 확인 후 실행하세요");
  return true;
}

function CodeCard({ code }: { code: string }) {
  return (
    <div className="rounded border bg-code-bg" data-testid="chat-code-card">
      <pre className="max-h-56 overflow-auto p-2 font-mono text-xs">{code}</pre>
      <div className="flex flex-wrap items-center gap-1 border-t px-1 py-0.5">
        <CopyButton text={code} />
        <Button
          variant="ghost"
          size="sm"
          className="h-6 px-2 text-xs"
          onClick={() => insertChatCode(code, "append")}
        >
          현재 블록에 추가
        </Button>
        <Button
          variant="ghost"
          size="sm"
          className="h-6 px-2 text-xs text-primary"
          onClick={() => insertChatCode(code, "below")}
        >
          아래 새 블록으로
        </Button>
        <Button
          variant="ghost"
          size="sm"
          className="h-6 px-2 text-xs"
          onClick={() => insertChatCode(code, "above")}
        >
          위 새 블록으로
        </Button>
      </div>
    </div>
  );
}

/** 도구 호출 한 줄 로그 (L.2 — 접이식) */
function ToolLogRow({ log }: { log: ToolLog }) {
  return (
    <details className="rounded border bg-muted/30 px-1.5 py-0.5" data-testid="tool-log">
      <summary className="cursor-pointer truncate text-[11px] text-muted-foreground">
        {log.summary}
      </summary>
      <pre className="mt-1 max-h-24 overflow-auto font-mono text-[10px] text-muted-foreground">
        {log.name}({JSON.stringify(log.input)})
      </pre>
    </details>
  );
}

function RefChips({ refs }: { refs: RefChip[] }) {
  return (
    <div className="flex flex-wrap gap-1">
      {refs.map((r, i) => (
        <button
          key={i}
          data-testid="ref-chip"
          onClick={() => gotoRef(r, refs)}
          title="그리드에서 이 범위로 이동"
          className="rounded bg-[var(--chip-amber-bg)] px-1.5 py-0.5 font-mono text-[11px] text-[var(--chip-amber-fg)] hover:underline"
        >
          {r.label}
        </button>
      ))}
    </div>
  );
}

function CellProposalCard({
  proposal: p,
  state,
  onResolve,
}: {
  proposal: CellProposal;
  state?: "applied" | "ignored";
  onResolve: (result: "applied" | "ignored") => void;
}) {
  const cells = p.values.reduce((n, row) => n + row.length, 0);
  const preview = p.values.slice(0, 10);
  return (
    <div className="rounded border border-primary/40 bg-accent/30" data-testid="cell-proposal">
      <div className="flex flex-wrap items-center gap-1 border-b px-2 py-1">
        <span className="text-xs font-medium">셀 제안</span>
        <span className="rounded bg-[var(--chip-teal-bg)] px-1.5 font-mono text-[11px] text-[var(--chip-teal-fg)]">
          {p.label}
        </span>
        <span className="text-[11px] text-muted-foreground">{cells}셀</span>
      </div>
      {p.reason && <p className="px-2 pt-1.5 text-xs">{p.reason}</p>}
      {p.formulaCells > 0 && (
        <p className="px-2 pt-1 text-xs text-[var(--warning-text)]">
          수식 셀 {p.formulaCells}개를 덮어씁니다 — 적용 후 되돌리려면 Ctrl+Z
        </p>
      )}
      <div className="max-h-40 overflow-auto p-2">
        <table className="w-full border-collapse text-[11px]">
          <tbody>
            {preview.map((row, i) => (
              <tr key={i}>
                {row.map((v, j) => (
                  <td key={j} className="border px-1 py-0.5 font-mono">
                    {v === null ? "" : String(v)}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
        {p.values.length > preview.length && (
          <p className="pt-1 text-[11px] text-muted-foreground">
            … 외 {p.values.length - preview.length}행
          </p>
        )}
      </div>
      <div className="flex items-center justify-end gap-1 border-t px-1 py-1">
        {state ? (
          <span className="px-2 text-[11px] text-muted-foreground">
            {state === "applied" ? "적용됨" : "무시함"}
          </span>
        ) : (
          <>
            <Button
              size="sm"
              className="h-6 px-2 text-xs"
              onClick={() => {
                const res = applyCellProposal(p);
                if ("error" in res) {
                  toast.error(res.error);
                  return;
                }
                notifyWorkbookEdit([res.range]);
                onResolve("applied");
                toast("셀에 적용했습니다 — 되돌리려면 Ctrl+Z");
              }}
            >
              적용
            </Button>
            <Button
              variant="ghost"
              size="sm"
              className="h-6 px-2 text-xs"
              onClick={() => onResolve("ignored")}
            >
              무시
            </Button>
          </>
        )}
      </div>
    </div>
  );
}

function BlockProposalCard({
  proposal: p,
  state,
  onResolve,
}: {
  proposal: BlockProposal;
  state?: "applied" | "ignored";
  onResolve: (result: "applied" | "ignored") => void;
}) {
  const apply = (placement: "append" | "below" | "above") => {
    if (
      insertChatCode(p.code, placement, {
        targetBlockId: p.targetBlockId,
        title: p.title,
        note: p.note,
        replace: placement === "append",
      })
    ) {
      onResolve("applied");
    }
  };
  return (
    <div className="rounded border border-primary/40 bg-code-bg" data-testid="block-proposal">
      <div className="flex flex-wrap items-center gap-1 border-b px-2 py-1">
        <span className="text-xs font-medium">코드 제안</span>
        {p.title && <span className="truncate text-[11px] text-muted-foreground">{p.title}</span>}
        {p.targetLabel && (
          <span className="rounded bg-[var(--chip-teal-bg)] px-1.5 font-mono text-[11px] text-[var(--chip-teal-fg)]">
            {p.targetLabel}
          </span>
        )}
      </div>
      {p.note && <p className="px-2 pt-1.5 text-xs text-muted-foreground">{p.note}</p>}
      <pre className="max-h-56 overflow-auto p-2 font-mono text-xs">{p.code}</pre>
      <div className="flex flex-wrap items-center gap-1 border-t px-1 py-0.5">
        <CopyButton text={p.code} />
        {state ? (
          <span className="px-2 text-[11px] text-muted-foreground">
            {state === "applied" ? "적용됨" : "무시함"}
          </span>
        ) : (
          <>
            <Button
              variant="ghost"
              size="sm"
              className="h-6 px-2 text-xs text-primary"
              onClick={() => apply("append")}
            >
              현재 블록에 적용
            </Button>
            <Button
              variant="ghost"
              size="sm"
              className="h-6 px-2 text-xs"
              onClick={() => apply("below")}
            >
              아래 새 블록
            </Button>
            <Button
              variant="ghost"
              size="sm"
              className="h-6 px-2 text-xs"
              onClick={() => apply("above")}
            >
              위 새 블록
            </Button>
            <Button
              variant="ghost"
              size="sm"
              className="h-6 px-2 text-xs"
              onClick={() => onResolve("ignored")}
            >
              무시
            </Button>
          </>
        )}
      </div>
    </div>
  );
}

export default function AiChatPanel({ onClose }: { onClose?: () => void }) {
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [input, setInput] = useState("");
  const [attachment, setAttachment] = useState<ChatAttachment | null>(null);
  const [busy, setBusy] = useState(false);
  const [liveTools, setLiveTools] = useState<ToolLog[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [instructions, setInstructions] = useState(DEFAULT_CHAT_INSTRUCTIONS);
  const [ready, setReady] = useState(false);
  const [autoSend, setAutoSend] = useState<string | null>(null);
  const [instrOpen, setInstrOpen] = useState(false);
  const [instrDraft, setInstrDraft] = useState("");
  const listRef = useRef<HTMLDivElement>(null);

  // 마운트: 저장된 이력·지침 복원 + 대기 중 요청 소비
  useEffect(() => {
    void loadSettings().then((s) => {
      if (s?.aiChatHistory) setMessages(s.aiChatHistory);
      if (s?.aiChatInstructions !== undefined) setInstructions(s.aiChatInstructions);
      setReady(true);
    });
    const take = () => {
      if (!pendingAsk) return;
      const ask = pendingAsk;
      pendingAsk = null;
      if (ask.attachment) setAttachment(ask.attachment);
      if (ask.question?.trim()) setAutoSend(ask.question.trim());
      else if (ask.question !== undefined) setInput("");
    };
    take();
    window.addEventListener(ASK_EVENT, take);
    return () => window.removeEventListener(ASK_EVENT, take);
  }, []);

  // 새 메시지 → 맨 아래로
  useEffect(() => {
    listRef.current?.scrollTo({ top: listRef.current.scrollHeight });
  }, [messages, busy, liveTools]);

  const persist = useCallback((next: ChatMessage[]) => {
    setMessages(next);
    void saveSettings({ aiChatHistory: capHistory(next) });
  }, []);

  const replaceMessage = (index: number, patch: Partial<ChatMessage>) => {
    persist(messages.map((m, i) => (i === index ? { ...m, ...patch } : m)));
  };

  const saveInstructions = (value: string) => {
    setInstructions(value);
    void saveSettings({ aiChatInstructions: value });
  };

  const send = useCallback(
    async (question?: string) => {
      const q = (question ?? input).trim();
      if (!q || busy) return;
      const key = await getApiKey();
      if (!key) {
        openApiKeyDialog();
        return;
      }
      const user: ChatMessage = { role: "user", content: withAttachment(q, attachment) };
      const base = [...messages, user];
      setBusy(true);
      setError(null);
      setLiveTools([]);
      setInput("");
      setAttachment(null);
      persist(base);
      try {
        // 워크북 개요(값 미포함)는 전송할 마지막 메시지에만 — 이력에는 남기지 않는다
        const sent = chatMessages(base);
        const last = sent[sent.length - 1];
        if (last && typeof last.content === "string") {
          last.content = `${last.content}\n\n${sheetOverview(useWorkbookStore.getState().workbook)}`;
        }
        const res = await runAgent({
          apiKey: key,
          system: buildChatSystem(instructions),
          messages: sent,
          onToolCall: (log) => setLiveTools((t) => [...t, log]),
        });
        if (res.refs.length > 0) useWorkbookStore.getState().setChatRefs(res.refs);
        const text =
          res.text ||
          (res.proposals.length > 0 ? "제안을 만들었습니다 — 아래 카드에서 확인하세요." : "(빈 응답)");
        persist([
          ...base,
          {
            role: "assistant",
            content: text,
            ...(res.tools.length > 0 ? { tools: res.tools } : {}),
            ...(res.proposals.length > 0 ? { proposals: res.proposals } : {}),
            ...(res.refs.length > 0 ? { refs: res.refs } : {}),
          },
        ]);
      } catch (e) {
        setError((e as Error).message);
      } finally {
        setBusy(false);
        setLiveTools([]);
      }
    },
    [attachment, busy, input, instructions, messages, persist],
  );

  // 블록 ✦ 액션·생성 바에서 온 요청 — 설정(지침) 복원 후 자동 전송
  useEffect(() => {
    if (!ready || autoSend === null || busy) return;
    const q = autoSend;
    setAutoSend(null);
    void send(q);
  }, [ready, autoSend, busy, send]);

  /** 편집기 드래그 선택(없으면 커서 줄)을 첨부로 */
  const attachSelection = () => {
    const st = useWorkbookStore.getState();
    const id = st.lastEditorBlockId;
    const handle = id ? editorRegistry.get(id) : undefined;
    if (!handle) {
      toast("블록 편집기를 먼저 클릭한 뒤 선택하세요");
      return;
    }
    const text = handle.getSelection();
    if (!text.trim()) {
      toast("선택된 코드가 없습니다");
      return;
    }
    setAttachment({ label: "선택 코드", text });
  };

  const renderProposal = (p: Proposal, msgIndex: number, msg: ChatMessage) => {
    const onResolve = (result: "applied" | "ignored") =>
      replaceMessage(msgIndex, { resolved: { ...(msg.resolved ?? {}), [p.id]: result } });
    const state = msg.resolved?.[p.id];
    return p.kind === "cells" ? (
      <CellProposalCard key={p.id} proposal={p} state={state} onResolve={onResolve} />
    ) : (
      <BlockProposalCard key={p.id} proposal={p} state={state} onResolve={onResolve} />
    );
  };

  return (
    <div className="flex h-full flex-col border-l bg-card" data-testid="ai-chat-panel">
      <div className="flex h-8 shrink-0 items-center gap-1 border-b px-2">
        <span className="text-xs font-medium">AI 채팅</span>
        <Button
          variant="ghost"
          size="icon-xs"
          className="ml-auto"
          onClick={() => {
            setInstrDraft(instructions);
            setInstrOpen(true);
          }}
          aria-label="채팅 지침 편집"
          title="채팅 지침"
        >
          <Gear />
        </Button>
        <Button
          variant="ghost"
          size="icon-xs"
          disabled={messages.length === 0}
          onClick={() => {
            persist([]);
            useWorkbookStore.getState().setChatRefs([]);
          }}
          aria-label="대화 지우기"
          title="대화 지우기"
        >
          <Broom />
        </Button>
        {onClose && (
          <Button
            variant="ghost"
            size="icon-xs"
            onClick={onClose}
            aria-label="AI 채팅 패널 닫기"
            title="닫기"
          >
            <X />
          </Button>
        )}
      </div>

      {/* 메시지 목록 */}
      <div ref={listRef} className="min-h-0 flex-1 space-y-2 overflow-y-auto p-2">
        {messages.length === 0 && (
          <p className="px-2 py-8 text-center text-xs text-muted-foreground">
            시트 데이터·코드·오류에 대해 무엇이든 물어보세요.
            <br />
            필요하면 AI가 시트를 직접 읽고, 셀·코드 변경을 제안합니다.
            <br />
            제안은 [적용]을 눌러야 반영되고 자동 실행은 없습니다.
          </p>
        )}
        {messages.map((m, i) => {
          if (m.role === "user") {
            return (
              <div
                key={i}
                className="ml-6 whitespace-pre-wrap rounded border bg-accent/40 px-2 py-1.5 text-xs"
                data-role="user"
              >
                {m.content}
              </div>
            );
          }
          const { body, instruction } = parseInstructionFence(m.content);
          return (
            <div key={i} className="mr-2 space-y-1.5 text-xs" data-role="assistant">
              {m.tools?.map((log, j) => <ToolLogRow key={j} log={log} />)}
              {m.refs && m.refs.length > 0 && <RefChips refs={m.refs} />}
              {splitCodeBlocks(body).map((seg, j) =>
                seg.type === "code" ? (
                  <CodeCard key={j} code={seg.content} />
                ) : (
                  <div key={j} className="space-y-1 px-1">
                    {renderMarkdown(seg.content)}
                  </div>
                ),
              )}
              {m.proposals?.map((p) => renderProposal(p, i, m))}
              {instruction && (
                <div
                  className="rounded border border-[var(--warning)]/50 bg-[var(--chip-amber-bg)]/50 px-2 py-1.5"
                  data-testid="instruction-proposal"
                >
                  <p className="text-xs">
                    지침에 추가: <span className="font-medium">{instruction}</span>
                  </p>
                  <div className="mt-1 flex justify-end gap-1">
                    <Button
                      size="sm"
                      className="h-6 px-2 text-xs"
                      onClick={() => {
                        saveInstructions(`${instructions.trim()}\n${instruction}`);
                        replaceMessage(i, { content: body }); // 펜스 제거 — 카드 재표시 방지
                        toast("지침에 추가했습니다 — 다음 대화부터 적용됩니다");
                      }}
                    >
                      반영
                    </Button>
                    <Button
                      variant="ghost"
                      size="sm"
                      className="h-6 px-2 text-xs"
                      onClick={() => replaceMessage(i, { content: body })}
                    >
                      무시
                    </Button>
                  </div>
                </div>
              )}
            </div>
          );
        })}
        {busy && (
          <div className="space-y-1">
            {liveTools.map((log, j) => (
              <ToolLogRow key={j} log={log} />
            ))}
            <p className="px-2 text-xs text-muted-foreground">응답 생성 중…</p>
          </div>
        )}
        {error && <p className="px-2 text-xs text-destructive">{error}</p>}
      </div>

      {/* 첨부 카드 + 입력 */}
      <div className="shrink-0 border-t p-2">
        {attachment && (
          <div className="mb-1 flex items-center gap-1 rounded border bg-muted/40 px-2 py-1">
            <span className="min-w-0 flex-1 truncate font-mono text-[11px] text-muted-foreground">
              [{attachment.label}] {attachment.text.split("\n")[0]}
            </span>
            <Button
              variant="ghost"
              size="icon-xs"
              onClick={() => setAttachment(null)}
              aria-label="첨부 제거"
            >
              <X />
            </Button>
          </div>
        )}
        <div className="flex items-end gap-1">
          <Button
            variant="ghost"
            size="icon-xs"
            className="mb-1"
            onClick={attachSelection}
            aria-label="선택 코드 질문"
            title="편집기에서 선택한 코드(없으면 커서 줄)를 첨부"
          >
            <PaperclipHorizontal />
          </Button>
          <textarea
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter" && !e.shiftKey && !e.nativeEvent.isComposing) {
                e.preventDefault();
                void send();
              }
            }}
            rows={2}
            placeholder="질문 입력 (Enter=전송, Shift+Enter=줄바꿈)"
            aria-label="AI 채팅 입력"
            disabled={busy}
            className="min-h-0 flex-1 resize-none rounded border bg-background p-1.5 text-xs outline-none placeholder:text-muted-foreground focus:border-primary"
          />
          <Button
            size="icon-sm"
            className={cn("mb-0.5 shrink-0")}
            disabled={busy || !input.trim()}
            onClick={() => void send()}
            aria-label="전송"
          >
            <ArrowUp />
          </Button>
        </div>
      </div>

      {/* 지침 편집 다이얼로그 (G.3) */}
      <Dialog open={instrOpen} onOpenChange={setInstrOpen}>
        <DialogContent className="max-w-[560px]">
          <DialogHeader>
            <DialogTitle>채팅 지침</DialogTitle>
            <DialogDescription>
              이 브라우저에만 저장되며 워크북 파일에는 포함되지 않습니다. 앱 규칙(sheet()·안전
              규칙)과 충돌하면 앱 규칙이 우선합니다.
            </DialogDescription>
          </DialogHeader>
          <textarea
            value={instrDraft}
            onChange={(e) => setInstrDraft(e.target.value)}
            rows={8}
            aria-label="채팅 지침"
            className="w-full resize-y rounded border bg-background p-2 text-xs outline-none focus:border-primary"
          />
          <DialogFooter>
            <Button
              variant="ghost"
              className="mr-auto"
              onClick={() => setInstrDraft(DEFAULT_CHAT_INSTRUCTIONS)}
            >
              기본값 복원
            </Button>
            <Button variant="ghost" onClick={() => setInstrOpen(false)}>
              닫기
            </Button>
            <Button
              onClick={() => {
                saveInstructions(instrDraft.trim());
                setInstrOpen(false);
                toast("지침을 저장했습니다 — 다음 대화부터 적용됩니다");
              }}
            >
              저장
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}
