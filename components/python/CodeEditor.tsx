"use client";

// CodeMirror 6 파이썬 편집기 — xl("...") 리터럴 하이라이트 + 커서 시 그리드 범위 하이라이트(§4.8)

import { python } from "@codemirror/lang-python";
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

const XL_RE = /xl\(\s*(["'])([^"']+)\1/g;

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
      viewRef.current?.focus();
      hostRef.current?.scrollIntoView({ block: "nearest" });
      useWorkbookStore.getState().setFocusBlock(null);
    }
  }, [focusRequested]);

  return <div ref={hostRef} className={cn("max-h-72 overflow-auto text-xs", className)} />;
}
