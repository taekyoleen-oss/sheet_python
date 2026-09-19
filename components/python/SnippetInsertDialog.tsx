"use client";

// 부록 F.1 — 코드 삽입 팝업: 그룹(예제 코드 9 + 핸들링 12 + 그래프 3) → 스니펫(그래프는 SVG 미리보기)
// → 삽입될 코드 전체 미리보기 → [현재 블록에 추가 | 아래 새 블록으로 | 위 새 블록으로].
// 기준 블록 = 마지막으로 편집기 포커스를 받은 코드 블록. Enter = 마지막 사용 위치로 삽입.
// 새 블록은 자동 실행하지 않는다.
// 부록 G.1 — 자리표시자(df·"…열"·{{range}}) amber 표시 + 드롭다운 치환(이름만, 로직 불변).

import { Fragment, useEffect, useMemo, useState, type ReactNode } from "react";
import { toast } from "sonner";
import { openFitGuideDialog } from "@/components/python/FitGuideDialog";
import { CopyButton, highlightPython } from "@/components/reference/code-popup";
import { PLOT_META, PlotSampleSvg } from "@/components/reference/PlotSampleSvg";
import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
  DialogTrigger,
} from "@/components/ui/dialog";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { sheetSchemas } from "@/lib/ai/schema";
import { formatA1 } from "@/lib/grid/a1";
import { notifyWorkbookEdit } from "@/lib/grid/calc-host";
import { codeTitle } from "@/lib/grid/code-sections";
import { appendSnippetToBlock, insertSnippetAsBlock } from "@/lib/grid/insert-snippet";
import { useWorkbookStore } from "@/lib/grid/model";
import { outputsOf } from "@/lib/grid/outputs";
import {
  detectPlaceholders,
  PLACEHOLDER_RE,
  substitutePlaceholders,
} from "@/lib/grid/snippet-placeholders";
import { xlRefForSelection } from "@/lib/grid/xl-ref";
import { getRuntimeClient } from "@/lib/runtime/client";
import { EXAMPLE_SNIPPET_FAMILIES } from "@/lib/reference/exampleSnippets";
import { PLOT_SNIPPET_GROUPS, plotInsertCode } from "@/lib/reference/plotSnippets";
import { WRANGLE_SNIPPET_GROUPS, snippetInsertCode } from "@/lib/reference/wrangleSnippets";
import { cn } from "@/lib/utils";

type Placement = "append" | "below" | "above";

interface Group {
  key: string;
  /** 코드 조립 방식 — plot만 다르다(예제 조각은 wrangle과 같은 규칙) */
  kind: "wrangle" | "plot";
  /** 좌측 목록의 상위 카테고리 제목 */
  section: string;
  label: string;
  snippets: { id: string; label: string; desc: string; code: string }[];
}

// 좌측 목록 순서 = 예제 코드 상위 카테고리(통계분석 · 위험률 산출) → 핸들링 → 그래프.
// 통계분석과 위험률 산출은 절차가 달라 한 목록에 섞지 않는다(부록 N).
const GROUPS: Group[] = [
  ...EXAMPLE_SNIPPET_FAMILIES.flatMap((fam) =>
    fam.groups.map((g) => ({
      key: `e:${fam.id}:${g.id}`,
      kind: "wrangle" as const,
      section: fam.label,
      label: g.label,
      snippets: g.snippets,
    })),
  ),
  ...WRANGLE_SNIPPET_GROUPS.map((g) => ({
    key: `w:${g.id}`,
    kind: "wrangle" as const,
    section: "핸들링",
    label: g.label,
    snippets: g.snippets,
  })),
  ...PLOT_SNIPPET_GROUPS.map((g) => ({
    key: `p:${g.id}`,
    kind: "plot" as const,
    section: "그래프",
    label: g.label,
    snippets: g.snippets,
  })),
];

/** 좌측 목록에 보이는 섹션 순서 (중복 제거) */
const SECTIONS: string[] = [...new Set(GROUPS.map((g) => g.section))];

/** 자리표시자를 amber로 감싼 코드 미리보기 (경량 하이라이트 + 복사) */
function PreviewCode({ code }: { code: string }) {
  const nodes = useMemo(() => {
    const out: ReactNode[] = [];
    let last = 0;
    let key = 0;
    for (const m of code.matchAll(new RegExp(PLACEHOLDER_RE.source, "g"))) {
      const idx = m.index ?? 0;
      out.push(<Fragment key={key++}>{highlightPython(code.slice(last, idx))}</Fragment>);
      out.push(
        <span
          key={key++}
          data-ph={m[2] ?? m[0]}
          className="rounded bg-[var(--chip-amber-bg)] px-0.5 text-[var(--chip-amber-fg)]"
        >
          {m[0]}
        </span>,
      );
      last = idx + m[0].length;
    }
    out.push(<Fragment key={key++}>{highlightPython(code.slice(last))}</Fragment>);
    return out;
  }, [code]);
  return (
    <div className="relative mt-2">
      <CopyButton text={code} className="absolute right-2 top-2 z-10" />
      <pre
        className="overflow-x-auto rounded border bg-muted px-4 py-3.5 font-mono leading-[1.75] text-foreground"
        style={{ fontSize: 12 }}
      >
        <code>{nodes}</code>
      </pre>
    </div>
  );
}

/** 마지막으로 실행된 블록의 출력 변수 — dfVars 안에 있으면 기본 선택 후보 (G.1) */
function lastRunVariable(dfVars: string[]): string | undefined {
  const blocks = useWorkbookStore.getState().workbook.pyBlocks;
  let best: { ranAt: string; name: string } | undefined;
  for (const b of blocks) {
    for (const o of outputsOf(b)) {
      const name = o.selection?.variable;
      const ranAt = o.last?.ranAt;
      if (!name || !ranAt || !dfVars.includes(name)) continue;
      if (!best || ranAt > best.ranAt) best = { ranAt, name };
    }
  }
  return best?.name;
}

function GroupList({
  active,
  onSelect,
  onFitGuide,
}: {
  active: string;
  onSelect: (key: string) => void;
  onFitGuide: () => void;
}) {
  const section = (title: string) => (
    <Fragment key={title}>
      <p className="px-2 pb-0.5 pt-2 text-[11px] font-semibold text-muted-foreground">
        {title}
      </p>
      {GROUPS.filter((g) => g.section === title).map((g) => (
        <button
          key={g.key}
          onClick={() => onSelect(g.key)}
          aria-pressed={g.key === active}
          className={cn(
            "block w-full truncate px-2 py-1 text-left text-xs hover:bg-accent",
            g.key === active ? "bg-accent font-medium text-primary" : "text-foreground/80",
          )}
        >
          {g.label}
        </button>
      ))}
    </Fragment>
  );
  return (
    <div className="w-44 shrink-0 overflow-y-auto rounded border py-1" data-testid="snippet-groups">
      {SECTIONS.map(section)}
      <p className="px-2 pb-0.5 pt-2 text-[11px] font-semibold text-muted-foreground">가이드</p>
      <button
        onClick={onFitGuide}
        className="block w-full truncate px-2 py-1 text-left text-xs text-foreground/80 hover:bg-accent"
      >
        모델적합 가이드…
      </button>
    </div>
  );
}

export default function SnippetInsertDialog() {
  const [open, setOpen] = useState(false);
  const [groupKey, setGroupKey] = useState(GROUPS[0].key);
  const [snippetId, setSnippetId] = useState(GROUPS[0].snippets[0]?.id ?? "");
  const [lastPlacement, setLastPlacement] = useState<Placement>("below");

  // 기준 블록 = 마지막으로 편집기 포커스를 받은 코드 블록 (v1.2 lastEditorBlockId)
  const refBlock = useWorkbookStore((s) =>
    s.workbook.pyBlocks.find((b) => b.id === s.lastEditorBlockId && b.kind !== "markdown"),
  );
  const refSheetName = useWorkbookStore((s) =>
    refBlock ? s.workbook.sheets.find((sh) => sh.id === refBlock.sheetId)?.name : undefined,
  );

  const group = GROUPS.find((g) => g.key === groupKey) ?? GROUPS[0];
  const snippet = group.snippets.find((s) => s.id === snippetId) ?? group.snippets[0];
  const insertText = useMemo(() => {
    if (!snippet) return "";
    return group.kind === "plot" ? plotInsertCode(snippet) : snippetInsertCode(snippet);
  }, [group.kind, snippet]);

  // ── 자리표시자 치환 (부록 G.1) — 이름만 바꾼다, 로직 불변 ──
  const placeholders = useMemo(() => detectPlaceholders(insertText), [insertText]);
  const [subs, setSubs] = useState<Record<string, string>>({});
  const [dfVars, setDfVars] = useState<string[]>([]);
  useEffect(() => setSubs({}), [groupKey, snippetId]); // 스니펫이 바뀌면 선택 초기화

  // 열릴 때 런타임 DataFrame 변수 수집 (준비 전·실패는 빈 목록)
  useEffect(() => {
    if (!open) return;
    void (async () => {
      try {
        const client = getRuntimeClient();
        if (client.getStatus() !== "ready") return;
        setDfVars(
          (await client.inspect()).filter((v) => v.type === "DataFrame").map((v) => v.name),
        );
      } catch {
        /* 런타임 미준비 — 드롭다운은 '치환 안 함'만 */
      }
    })();
  }, [open]);

  // 자동 선택: DataFrame이 정확히 1개면 그것, 여럿이면 마지막 실행 블록의 출력 변수
  useEffect(() => {
    if (dfVars.length === 0) return;
    const auto = dfVars.length === 1 ? dfVars[0] : lastRunVariable(dfVars);
    if (!auto) return;
    setSubs((cur) => {
      const next = { ...cur };
      for (const p of placeholders) {
        if (p.kind === "variable" && !(p.token in next)) next[p.token] = auto;
      }
      return next;
    });
  }, [dfVars, placeholders]);

  // 열 후보 = 활성 시트 헤더 행 + 블록 표 미리보기의 열 이름
  const columnOptions = useMemo(() => {
    if (!open) return [] as string[];
    const st = useWorkbookStore.getState();
    const set = new Set<string>();
    const schema = sheetSchemas(st.workbook).find(
      (s, i) => st.workbook.sheets[i]?.id === st.activeSheetId,
    );
    for (const h of schema?.headers ?? []) if (h.trim()) set.add(h.trim());
    for (const b of st.workbook.pyBlocks) {
      for (const o of outputsOf(b)) {
        const preview = o.last?.preview as { columns?: string[] } | undefined;
        for (const c of preview?.columns ?? []) set.add(c);
      }
    }
    return [...set];
  }, [open]);

  const previewCode = useMemo(
    () => substitutePlaceholders(insertText, subs),
    [insertText, subs],
  );

  const selectGroup = (key: string) => {
    setGroupKey(key);
    const g = GROUPS.find((x) => x.key === key);
    setSnippetId(g?.snippets[0]?.id ?? "");
  };

  const doInsert = (placement: Placement) => {
    if (!snippet) return;
    const text = previewCode; // 치환 적용본 (미치환 토큰은 편집기에서 amber 표시)
    let targetId: string;
    if (placement === "append") {
      if (!refBlock || !appendSnippetToBlock(refBlock.id, text)) return;
      notifyWorkbookEdit([], [refBlock.id]); // 타이핑 커밋과 같은 통지 경로 (§2.3.3)
      targetId = refBlock.id;
    } else {
      const res = insertSnippetAsBlock(refBlock?.id ?? null, placement, snippet.label, text);
      if (!res) {
        toast.error("블록을 만들 수 없습니다 (활성 시트 없음)");
        return;
      }
      if (refBlock && !res.ordered) {
        toast("순서를 보장할 빈 위치가 없어 빈 영역에 배치했습니다 — ↑↓로 순서를 조정하세요");
      }
      targetId = res.id;
    }
    setLastPlacement(placement);
    setOpen(false);
    // 닫힘 뒤 대상 블록 포커스 (focusBlock 메커니즘 — CodeEditor가 이어받는다)
    requestAnimationFrame(() => {
      const st = useWorkbookStore.getState();
      st.setFocusBlock(targetId);
      st.setSelectedBlock(targetId);
    });
  };

  const refLabel = refBlock
    ? refBlock.title?.trim() ||
      codeTitle(refBlock.code) ||
      `${refSheetName ?? "?"}!${formatA1({
        r0: refBlock.anchor.r,
        c0: refBlock.anchor.c,
        r1: refBlock.anchor.r,
        c1: refBlock.anchor.c,
      })}`
    : null;

  return (
    <Dialog open={open} onOpenChange={setOpen}>
      <DialogTrigger asChild>
        <Button variant="outline" size="sm" className="h-7 shrink-0 px-2 text-xs">
          코드 삽입
        </Button>
      </DialogTrigger>
      <DialogContent
        className="flex h-[85vh] max-w-[960px] flex-col gap-2 sm:max-w-[960px]"
        onCloseAutoFocus={(e) => e.preventDefault()}
        onKeyDown={(e) => {
          if (e.key !== "Enter" || e.nativeEvent.isComposing) return;
          e.preventDefault();
          doInsert(lastPlacement === "append" && !refBlock ? "below" : lastPlacement);
        }}
      >
        <DialogHeader>
          <DialogTitle className="text-sm">
            코드 삽입 — 예제 코드(통계분석·위험률 산출)·핸들링·그래프
          </DialogTitle>
          <DialogDescription className="text-xs">
            스니펫을 고르면 삽입될 코드가 그대로 보입니다. Enter = 마지막 사용 위치로 삽입.
          </DialogDescription>
        </DialogHeader>

        <div className="flex min-h-0 flex-1 gap-2">
          <GroupList
            active={group.key}
            onSelect={selectGroup}
            onFitGuide={() => {
              // 팝업을 닫고 모델적합 가이드 마법사로 전환 (부록 H.3 진입점)
              setOpen(false);
              openFitGuideDialog();
            }}
          />

          {/* 스니펫 목록 — 그래프는 SVG 미리보기 썸네일 */}
          <div className="w-64 shrink-0 overflow-y-auto rounded border py-1">
            {group.snippets.map((s) => {
              const activeSnip = s.id === snippet?.id;
              const meta = group.kind === "plot" ? PLOT_META[s.id] : undefined;
              return (
                <button
                  key={s.id}
                  onClick={() => setSnippetId(s.id)}
                  aria-pressed={activeSnip}
                  className={cn(
                    "flex w-full items-center gap-2 px-2 py-1.5 text-left hover:bg-accent",
                    activeSnip && "bg-accent",
                  )}
                >
                  {meta && (
                    <span className="w-16 shrink-0 overflow-hidden rounded border bg-card">
                      <PlotSampleSvg shape={meta.shape} />
                    </span>
                  )}
                  <span className="min-w-0 flex-1">
                    <span
                      className={cn(
                        "block truncate text-xs",
                        activeSnip ? "font-medium text-primary" : "font-medium",
                      )}
                    >
                      {s.label}
                    </span>
                    <span className="block truncate text-[11px] text-muted-foreground">
                      {s.desc}
                    </span>
                  </span>
                </button>
              );
            })}
          </div>

          {/* 삽입될 코드 전체 미리보기 — 자리표시자 amber + 치환 드롭다운 (부록 G.1) */}
          <div
            className="min-w-0 flex-1 overflow-y-auto rounded border px-2 pb-2"
            data-testid="snippet-code-preview"
          >
            {snippet ? (
              <>
                {placeholders.length > 0 && (
                  <div
                    className="mt-2 flex flex-wrap items-center gap-x-3 gap-y-1 rounded border bg-[var(--chip-amber-bg)]/40 px-2 py-1.5"
                    data-testid="placeholder-controls"
                  >
                    <span className="text-[11px] text-muted-foreground">
                      자리표시자 치환(이름만) —
                    </span>
                    {placeholders.map((p) => {
                      const options =
                        p.kind === "variable"
                          ? dfVars
                          : p.kind === "column"
                            ? columnOptions
                            : [xlRefForSelection(refBlock?.sheetId) ?? 'xl("A1")'];
                      const current = subs[p.token] ?? p.token;
                      return (
                        <span key={p.token} className="flex items-center gap-1">
                          <code className="rounded bg-[var(--chip-amber-bg)] px-1 font-mono text-[11px] text-[var(--chip-amber-fg)]">
                            {p.token}
                          </code>
                          <span aria-hidden className="text-[11px] text-muted-foreground">
                            →
                          </span>
                          <Select
                            value={current}
                            onValueChange={(v) =>
                              setSubs((cur) => {
                                const next = { ...cur };
                                if (v === p.token) delete next[p.token];
                                else next[p.token] = v;
                                return next;
                              })
                            }
                          >
                            <SelectTrigger
                              className="h-6 w-36 text-xs"
                              aria-label={`자리표시자 ${p.token}`}
                            >
                              <SelectValue />
                            </SelectTrigger>
                            <SelectContent>
                              <SelectItem value={p.token}>치환 안 함</SelectItem>
                              {options
                                .filter((o) => o !== p.token)
                                .map((o) => (
                                  <SelectItem key={o} value={o} className="font-mono">
                                    {o}
                                  </SelectItem>
                                ))}
                            </SelectContent>
                          </Select>
                        </span>
                      );
                    })}
                  </div>
                )}
                <PreviewCode code={previewCode} />
              </>
            ) : (
              <p className="p-4 text-xs text-muted-foreground">스니펫을 선택하세요</p>
            )}
          </div>
        </div>

        {/* 삽입 위치 — 정확히 세 버튼 (기준 블록 없으면 새 블록 하나로 축약) */}
        <div className="flex shrink-0 items-center gap-1.5 border-t pt-2">
          <span className="mr-auto min-w-0 truncate text-xs text-muted-foreground">
            {refLabel ? `기준 블록: ${refLabel}` : "기준 블록 없음 — 새 블록은 빈 영역에 만듭니다"}
          </span>
          <Button
            variant="outline"
            size="sm"
            className="h-7 px-2 text-xs"
            disabled={!refBlock || !snippet}
            title={refBlock ? "기준 블록 코드 끝에 덧붙입니다" : "블록 편집기를 먼저 클릭하세요"}
            onClick={() => doInsert("append")}
          >
            현재 블록에 추가
          </Button>
          {refBlock ? (
            <>
              <Button
                size="sm"
                className="h-7 px-2 text-xs"
                disabled={!snippet}
                title="계산 순서상 기준 블록 바로 다음에 새 블록을 만듭니다"
                onClick={() => doInsert("below")}
              >
                아래 새 블록으로
              </Button>
              <Button
                variant="outline"
                size="sm"
                className="h-7 px-2 text-xs"
                disabled={!snippet}
                title="계산 순서상 기준 블록 바로 앞에 새 블록을 만듭니다"
                onClick={() => doInsert("above")}
              >
                위 새 블록으로
              </Button>
            </>
          ) : (
            <Button
              size="sm"
              className="h-7 px-2 text-xs"
              disabled={!snippet}
              onClick={() => doInsert("below")}
            >
              새 블록으로
            </Button>
          )}
        </div>
      </DialogContent>
    </Dialog>
  );
}
