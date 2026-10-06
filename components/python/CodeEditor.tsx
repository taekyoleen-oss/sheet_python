"use client";

// CodeMirror 6 파이썬 편집기 — xl("...") 리터럴 하이라이트 + 커서 시 그리드 범위 하이라이트(§4.8)

import {
  acceptCompletion,
  completionStatus,
  startCompletion,
  type Completion,
  type CompletionContext,
  type CompletionResult,
} from "@codemirror/autocomplete";
import { indentLess, indentMore } from "@codemirror/commands";
import { python, pythonLanguage } from "@codemirror/lang-python";
import { indentUnit } from "@codemirror/language";
import { Prec } from "@codemirror/state";
import {
  Decoration,
  MatchDecorator,
  ViewPlugin,
  keymap,
  placeholder as cmPlaceholder,
  type DecorationSet,
  type EditorView as EditorViewType,
  type ViewUpdate,
} from "@codemirror/view";
import { EditorView, basicSetup } from "codemirror";
import { useEffect, useRef } from "react";
import { parseA1 } from "@/lib/grid/a1";
import { useWorkbookStore } from "@/lib/grid/model";
import { getRuntimeClient } from "@/lib/runtime/client";
import { PLACEHOLDER_RE } from "@/lib/grid/snippet-placeholders";
import { cn } from "@/lib/utils";

/** 마운트된 블록 편집기 핸들 (참조 삽입 바·스니펫 메뉴·목차 서브 항목·AI 채팅이 사용) */
export interface EditorHandle {
  /** 커서 위치에 텍스트 삽입 */
  insert: (text: string) => void;
  /** 해당 줄(0-기반)로 커서 이동 + 스크롤 (부록 F.2 목차 서브 항목) */
  scrollToLine: (line: number) => void;
  /** 드래그 선택 텍스트, 없으면 커서 줄 (부록 G.2 "선택 코드 질문") */
  getSelection: () => string;
}
export const editorRegistry = new Map<string, EditorHandle>();

const XL_RE = /\b(?:sheet|xl)\(\s*(["'])([^"']+)\1/g;

const xlDecorator = new MatchDecorator({
  regexp: XL_RE,
  decoration: Decoration.mark({ class: "cm-xlref" }),
});

const xlHighlighter = ViewPlugin.fromClass(
  class {
    decorations: DecorationSet;
    constructor(view: EditorViewType) {
      this.decorations = xlDecorator.createDeco(view);
    }
    update(update: ViewUpdate) {
      this.decorations = xlDecorator.updateDeco(update, this.decorations);
    }
  },
  { decorations: (v) => v.decorations },
);

// 부록 G.1: 스니펫 자리표시자(df·"…열"·{{range}}) 표시 전용 amber 하이라이트.
// 토큰이 실제 이름으로 바뀌면 더 이상 매치되지 않아 자연 소멸한다.
const placeholderDecorator = new MatchDecorator({
  regexp: new RegExp(PLACEHOLDER_RE.source, "g"),
  decoration: Decoration.mark({ class: "cm-placeholder" }),
});

const placeholderHighlighter = ViewPlugin.fromClass(
  class {
    decorations: DecorationSet;
    constructor(view: EditorViewType) {
      this.decorations = placeholderDecorator.createDeco(view);
    }
    update(update: ViewUpdate) {
      this.decorations = placeholderDecorator.updateDeco(update, this.decorations);
    }
  },
  { decorations: (v) => v.decorations },
);

// ── 런타임 자동완성: 실행으로 만든 변수·속성(df.)·열 이름(df[") ─────────────────────────
// 런타임이 한가할 때만 묻는다(실행 중이면 기본 완성만). 결과는 잠깐 캐시한다.
const DOTTED = /[A-Za-z_]\w*(?:\.[A-Za-z_]\w*)*/.source;
const cache = new Map<string, { at: number; names: string[] }>();

/** 파이썬 식을 평가해 문자열 목록을 받는다 — repr('a', "b") 목록을 그대로 파싱 */
async function runtimeNames(expr: string): Promise<string[]> {
  const hit = cache.get(expr);
  if (hit && performance.now() - hit.at < 3000) return hit.names;
  const client = getRuntimeClient();
  if (client.getStatus() !== "ready") return [];
  const { repr } = await client.repl(expr, 5);
  const names: string[] = [];
  for (const m of (repr ?? "").matchAll(/'((?:[^'\\]|\\.)*)'|"((?:[^"\\]|\\.)*)"/g)) names.push(m[1] ?? m[2]);
  cache.set(expr, { at: performance.now(), names });
  return names;
}

async function runtimeCompletions(ctx: CompletionContext): Promise<CompletionResult | null> {
  // df["… → 열 이름
  const key = ctx.matchBefore(new RegExp(`(${DOTTED})\\[(["'])[^"']*$`));
  if (key) {
    const [, obj, q] = /^(.*?)\[(["'])/.exec(key.text) ?? [];
    const cols = await runtimeNames(`[str(c) for c in getattr(${obj}, "columns", [])]`).catch(() => []);
    const from = key.from + key.text.indexOf(q) + 1;
    return cols.length ? { from, options: cols.map((c) => ({ label: c, type: "property", apply: c })), validFor: /^[^"']*$/ } : null;
  }
  // obj.attr → dir(obj)
  const dot = ctx.matchBefore(new RegExp(`${DOTTED}\\.\\w*$`));
  if (dot) {
    const at = dot.text.lastIndexOf(".");
    const obj = dot.text.slice(0, at);
    const names = await runtimeNames(`[n for n in dir(${obj}) if not n.startswith("_")]`).catch(() => []);
    return names.length
      ? { from: dot.from + at + 1, options: names.map((n): Completion => ({ label: n, type: "method" })), validFor: /^\w*$/ }
      : null;
  }
  // 이름 → 실행으로 만든 전역 변수 + sheet()
  const word = ctx.matchBefore(/[A-Za-z_]\w*$/);
  if (!word && !ctx.explicit) return null;
  const vars = await runtimeNames(
    `[n for n, v in globals().items() if not n.startswith("_") and type(v).__name__ != "module"]`,
  ).catch(() => []);
  return {
    from: word?.from ?? ctx.pos,
    options: [
      { label: "sheet", type: "function", detail: '("A1:C10", headers=True)', apply: 'sheet("' },
      // 이 코드에 나오는 이름은 파이썬 기본 완성이 이미 낸다 — 겹치지 않게 뺀다
      ...vars
        .filter((n) => !new RegExp(`\\b${n}\\b`).test(ctx.state.doc.toString()))
        .map((n): Completion => ({ label: n, type: "variable", boost: 1 })),
    ],
    validFor: /^\w*$/,
  };
}

/** Tab: 자동완성 수락 → 줄 중간이면 자동완성 열기 → 그 밖엔 들여쓰기(4칸). Shift+Tab 내어쓰기 */
const tabKeymap = keymap.of([
  {
    key: "Tab",
    run: (view) => {
      if (acceptCompletion(view)) return true;
      const sel = view.state.selection.main;
      const before = view.state.doc.lineAt(sel.head).text.slice(0, sel.head - view.state.doc.lineAt(sel.head).from);
      if (sel.empty && /[\w.\["']$/.test(before) && completionStatus(view.state) === null) return startCompletion(view);
      if (sel.empty && before.trim() !== "") {
        view.dispatch(view.state.replaceSelection("    "));
        return true;
      }
      return indentMore(view);
    },
    shift: indentLess,
  },
]);

const editorTheme = EditorView.theme({
  "&": { fontSize: "11px", backgroundColor: "var(--code-bg)" },
  ".cm-content": { fontFamily: "var(--font-jetbrains), monospace", padding: "6px 8px" },
  ".cm-gutters": { display: "none" },
  ".cm-xlref": { color: "#4A90C2", backgroundColor: "#EAF3FA", borderRadius: "2px" },
  // 자리표시자 — amber(경고 계열), Python 관여 색(Sky Blue)과 구분 (부록 G.1)
  ".cm-placeholder": {
    backgroundColor: "var(--chip-amber-bg)",
    color: "var(--chip-amber-fg)",
    borderBottom: "1px dotted var(--warning)",
    borderRadius: "2px",
  },
  "&.cm-focused": { outline: "none" },
});

export default function CodeEditor({
  blockId,
  sheetId,
  value,
  onChange,
  onRun,
  placeholder,
  className,
}: {
  /** 있으면 registry 등록 + lastEditorBlock 추적 + focusBlockId 반응 */
  blockId?: string;
  /** xl() 참조의 기본 시트 (hover 하이라이트용) */
  sheetId?: string;
  value: string;
  onChange: (value: string) => void;
  onRun?: () => void;
  placeholder?: string;
  className?: string;
}) {
  const hostRef = useRef<HTMLDivElement>(null);
  const viewRef = useRef<EditorViewType | null>(null);
  const cbRef = useRef({ onChange, onRun, sheetId, blockId });
  cbRef.current = { onChange, onRun, sheetId, blockId };

  useEffect(() => {
    if (!hostRef.current) return;

    /** 커서가 xl("...") 안이면 해당 범위를 그리드에 하이라이트 */
    const syncHoverRange = (view: EditorViewType) => {
      const st = useWorkbookStore.getState();
      if (!view.hasFocus) {
        if (st.hoverRange) st.setHoverRange(null);
        return;
      }
      const pos = view.state.selection.main.head;
      const text = view.state.doc.toString();
      const re = new RegExp(XL_RE.source, "g");
      let match: RegExpExecArray | null;
      while ((match = re.exec(text)) !== null) {
        if (pos >= match.index && pos <= match.index + match[0].length) {
          try {
            const parsed = parseA1(match[2]);
            const targetSheetId =
              parsed.sheetName === undefined
                ? cbRef.current.sheetId
                : st.workbook.sheets.find((s) => s.name === parsed.sheetName)?.id;
            if (targetSheetId) {
              st.setHoverRange({ sheetId: targetSheetId, range: parsed.range });
              return;
            }
          } catch {
            /* 잘못된 참조는 무시 */
          }
        }
      }
      if (st.hoverRange) st.setHoverRange(null);
    };

    const view = new EditorView({
      parent: hostRef.current,
      doc: value,
      extensions: [
        basicSetup,
        python(),
        indentUnit.of("    "), // PEP 8 — 4칸
        pythonLanguage.data.of({ autocomplete: runtimeCompletions }),
        Prec.high(tabKeymap),
        xlHighlighter,
        placeholderHighlighter,
        editorTheme,
        // 접근성 + 테스트: textarea 시절과 같은 레이블 유지
        EditorView.contentAttributes.of({ "aria-label": "Python 코드" }),
        ...(placeholder ? [cmPlaceholder(placeholder)] : []),
        Prec.highest(
          keymap.of([
            {
              key: "Ctrl-Enter",
              mac: "Cmd-Enter",
              run: () => {
                cbRef.current.onRun?.();
                return true;
              },
            },
            {
              // 부록 F.4: Shift+Enter도 실행 — 실행기 없는 편집기(초기화 스크립트)는 줄바꿈 유지
              key: "Shift-Enter",
              run: () => {
                if (!cbRef.current.onRun) return false;
                cbRef.current.onRun();
                return true;
              },
            },
          ]),
        ),
        EditorView.updateListener.of((update) => {
          if (update.docChanged) cbRef.current.onChange(update.state.doc.toString());
          if (update.docChanged || update.selectionSet || update.focusChanged) {
            syncHoverRange(update.view);
          }
        }),
        EditorView.domEventHandlers({
          focus: () => {
            const id = cbRef.current.blockId;
            if (id) useWorkbookStore.getState().setLastEditorBlock(id);
          },
        }),
      ],
    });
    viewRef.current = view;

    const id = cbRef.current.blockId;
    if (id) {
      editorRegistry.set(id, {
        insert: (text) => {
          view.dispatch(view.state.replaceSelection(text));
          view.focus();
        },
        scrollToLine: (line) => {
          const ln = view.state.doc.line(
            Math.max(1, Math.min(line + 1, view.state.doc.lines)),
          );
          view.dispatch({
            selection: { anchor: ln.from },
            effects: EditorView.scrollIntoView(ln.from, { y: "start" }),
          });
          view.focus();
        },
        getSelection: () => {
          const sel = view.state.selection.main;
          if (!sel.empty) return view.state.doc.sliceString(sel.from, sel.to);
          return view.state.doc.lineAt(sel.head).text; // 선택 없으면 커서 줄
        },
      });
    }
    return () => {
      if (id) editorRegistry.delete(id);
      view.destroy();
      viewRef.current = null;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // 외부 변경(undo 등) 반영 — 편집 중이 아닐 때만
  useEffect(() => {
    const view = viewRef.current;
    if (view && !view.hasFocus && view.state.doc.toString() !== value) {
      view.dispatch({ changes: { from: 0, to: view.state.doc.length, insert: value } });
    }
  }, [value]);

  // 블록 추가 직후 포커스 (focusBlockId 신호)
  const focusRequested = useWorkbookStore((s) => blockId !== undefined && s.focusBlockId === blockId);
  useEffect(() => {
    if (focusRequested) {
      // 스크롤 없이 포커스 — 카드(PyBlockCard)가 상단 기준으로 정렬한다
      viewRef.current?.contentDOM.focus({ preventScroll: true });
      useWorkbookStore.getState().setFocusBlock(null);
    }
  }, [focusRequested]);

  return <div ref={hostRef} className={cn("max-h-72 overflow-auto text-xs", className)} />;
}
