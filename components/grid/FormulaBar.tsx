"use client";

// 수식 입력줄 + 이름 상자 — 부록 O.4 (엑셀 상단 막대)
// 이름 상자: 선택 주소 표시. 주소·이름 입력 → 이동, 새 이름 입력 → 선택 범위에 이름 정의.
// 입력줄: 활성 셀 원문(수식이면 =…) 편집, Enter 확정·Esc 취소. `=` 뒤 함수·이름 자동완성.

import { useEffect, useMemo, useRef, useState } from "react";
import { CaretDown, TrashSimple } from "@phosphor-icons/react";
import { toast } from "sonner";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { A1Error, formatA1, parseA1 } from "@/lib/grid/a1";
import { notifyWorkbookEdit } from "@/lib/grid/calc-host";
import {
  cellFromInput,
  FORMULA_FUNCTIONS,
  FUNCTION_HELP,
  formatRef,
  parseRefText,
} from "@/lib/grid/formula";
import { useWorkbookStore } from "@/lib/grid/model";
import { cellKey, type Cell, type CellRange } from "@/types/workbook";

const store = () => useWorkbookStore.getState();

/** 셀 → 입력줄 원문 */
const rawOf = (cell: Cell | undefined): string =>
  !cell
    ? ""
    : (cell.fx ??
      (typeof cell.v === "boolean" ? (cell.v ? "TRUE" : "FALSE") : String(cell.v ?? "")));

/** 이름 정의 참조 → 시트·범위로 이동 */
function goTo(sheetName: string | undefined, range: CellRange): boolean {
  const st = store();
  const sheet = sheetName
    ? st.workbook.sheets.find((s) => s.name === sheetName)
    : st.workbook.sheets.find((s) => s.id === st.activeSheetId);
  if (!sheet) {
    toast.error(`시트를 찾을 수 없습니다: ${sheetName}`);
    return false;
  }
  if (sheet.id !== st.activeSheetId) st.setActiveSheet(sheet.id);
  // 전체 열/행 이름은 시트 크기로 자른다
  st.setSelection({
    r0: range.r0,
    c0: range.c0,
    r1: Math.min(range.r1, sheet.rowCount - 1),
    c1: Math.min(range.c1, sheet.colCount - 1),
  });
  return true;
}

interface Suggestion {
  text: string;
  insert: string;
  help: string;
}

export default function FormulaBar() {
  const selection = useWorkbookStore((s) => s.selection);
  const sheet = useWorkbookStore((s) => s.workbook.sheets.find((sh) => sh.id === s.activeSheetId));
  const names = useWorkbookStore((s) => s.workbook.names);
  const cell = selection && sheet ? sheet.cells[cellKey(selection.r0, selection.c0)] : undefined;
  const raw = rawOf(cell);

  const [draft, setDraft] = useState<string | null>(null); // null = 편집 중 아님
  const [caret, setCaret] = useState(0);
  const [pick, setPick] = useState(0);
  const [nameText, setNameText] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  // 선택이 바뀌면 편집 취소
  useEffect(() => {
    setDraft(null);
    setNameText(null);
  }, [selection?.r0, selection?.c0, sheet?.id]);

  // 이름 상자 표시: 선택이 이름 범위와 같으면 이름, 아니면 주소
  const address = useMemo(() => {
    if (!selection || !sheet) return "";
    for (const n of names ?? []) {
      const p = parseRefText(n.ref);
      if (!p || p.sheetName !== sheet.name) continue;
      const b = p.b ?? p.a;
      if (p.a.r === selection.r0 && p.a.c === selection.c0 && b.r === selection.r1 && b.c === selection.c1)
        return n.name;
    }
    return formatA1(selection);
  }, [selection, sheet, names]);

  const value = draft ?? raw;

  // `=` 수식에서 커서 앞 식별자 → 함수·이름 후보
  const suggestions = useMemo<Suggestion[]>(() => {
    if (draft === null || !draft.startsWith("=")) return [];
    const m = /([A-Za-z가-힣_][A-Za-z0-9가-힣_.]*)$/.exec(draft.slice(0, caret));
    if (!m || /\d$/.test(m[1]) && /^[A-Za-z]{1,3}\d+$/.test(m[1])) return [];
    const q = m[1].toUpperCase();
    const fns = FORMULA_FUNCTIONS.filter((f) => f.startsWith(q)).map((f) => ({
      text: f,
      insert: `${f}(`,
      help: FUNCTION_HELP[f] ?? "",
    }));
    const ns = (names ?? [])
      .filter((n) => n.name.toUpperCase().startsWith(q))
      .map((n) => ({ text: n.name, insert: n.name, help: `이름 → ${n.ref}` }));
    return [...ns, ...fns].slice(0, 8);
  }, [draft, caret, names]);

  // 커서가 함수 괄호 안이면 그 함수 도움말
  const argHelp = useMemo(() => {
    if (draft === null || !draft.startsWith("=")) return null;
    let depth = 0;
    const before = draft.slice(0, caret);
    for (let i = before.length - 1; i >= 0; i--) {
      const ch = before[i];
      if (ch === ")") depth++;
      else if (ch === "(") {
        if (depth === 0) {
          const m = /([A-Za-z][A-Za-z0-9.]*)$/.exec(before.slice(0, i));
          return m ? (FUNCTION_HELP[m[1].toUpperCase()] ?? null) : null;
        }
        depth--;
      }
    }
    return null;
  }, [draft, caret]);

  const accept = (sg: Suggestion) => {
    if (draft === null) return;
    const before = draft.slice(0, caret).replace(/[A-Za-z가-힣_][A-Za-z0-9가-힣_.]*$/, "");
    const next = before + sg.insert + draft.slice(caret);
    const pos = before.length + sg.insert.length;
    setDraft(next);
    setCaret(pos);
    requestAnimationFrame(() => inputRef.current?.setSelectionRange(pos, pos));
  };

  const commit = () => {
    if (draft === null || !selection || !sheet) return;
    const { r0, c0 } = selection;
    if (!store().setCellValue(sheet.id, r0, c0, cellFromInput(draft))) {
      toast.error("Python 출력(spill) 셀은 직접 편집할 수 없습니다");
      return;
    }
    notifyWorkbookEdit([{ sheetId: sheet.id, r0, c0, r1: r0, c1: c0 }]);
    setDraft(null);
    // 엑셀처럼 그리드로 포커스 복귀
    document.querySelector<HTMLElement>('[data-testid="data-grid-canvas"]')?.focus();
  };

  const onNameEnter = () => {
    const text = (nameText ?? "").trim();
    setNameText(null);
    if (!text || !sheet) return;
    const named = (names ?? []).find((n) => n.name.toUpperCase() === text.toUpperCase());
    if (named) {
      const p = parseRefText(named.ref);
      if (p) goTo(p.sheetName, { r0: p.a.r, c0: p.a.c, r1: (p.b ?? p.a).r, c1: (p.b ?? p.a).c });
      return;
    }
    try {
      const p = parseA1(text);
      goTo(p.sheetName, p.range);
      return;
    } catch (e) {
      if (!(e instanceof A1Error)) throw e;
    }
    // 새 이름 → 현재 선택을 절대 참조로 정의
    if (!selection) return;
    const ref = formatRef({
      sheetName: sheet.name,
      a: { r: selection.r0, c: selection.c0, ar: true, ac: true },
      b:
        selection.r0 === selection.r1 && selection.c0 === selection.c1
          ? undefined
          : { r: selection.r1, c: selection.c1, ar: true, ac: true },
    });
    const problem = store().defineName(text, ref);
    if (problem) toast.error(problem);
    else toast.success(`이름 '${text}' = ${ref}`);
  };

  const locked = !!cell?.src;

  return (
    <div
      data-testid="formula-bar"
      className="relative flex h-8 shrink-0 items-center gap-1 border-b bg-background px-2 text-xs"
    >
      <div className="flex items-center">
        <input
          aria-label="이름 상자"
          title="주소·이름을 입력하면 이동, 새 이름을 입력하면 선택 범위에 이름을 정의합니다"
          className="h-6 w-28 rounded-l border px-1.5 font-mono"
          value={nameText ?? address}
          onFocus={(e) => {
            setNameText(address);
            requestAnimationFrame(() => e.target.select());
          }}
          onChange={(e) => setNameText(e.target.value)}
          onBlur={() => setNameText(null)}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              e.preventDefault();
              onNameEnter();
              (e.target as HTMLInputElement).blur();
            } else if (e.key === "Escape") {
              setNameText(null);
              (e.target as HTMLInputElement).blur();
            }
          }}
        />
        <DropdownMenu>
          <DropdownMenuTrigger
            aria-label="정의된 이름 목록"
            className="flex h-6 w-5 items-center justify-center rounded-r border border-l-0 hover:bg-accent"
          >
            <CaretDown className="size-3" />
          </DropdownMenuTrigger>
          <DropdownMenuContent align="start" className="max-h-72 min-w-64 overflow-y-auto">
            <DropdownMenuLabel className="text-xs">정의된 이름</DropdownMenuLabel>
            <DropdownMenuSeparator />
            {(names ?? []).length === 0 ? (
              <p className="px-2 py-1.5 text-xs text-muted-foreground">
                범위를 선택하고 이름 상자에 새 이름을 입력하세요.
              </p>
            ) : (
              (names ?? []).map((n) => (
                <DropdownMenuItem
                  key={n.name}
                  className="flex items-center justify-between gap-3 text-xs"
                  onSelect={() => {
                    const p = parseRefText(n.ref);
                    if (p) goTo(p.sheetName, { r0: p.a.r, c0: p.a.c, r1: (p.b ?? p.a).r, c1: (p.b ?? p.a).c });
                  }}
                >
                  <span className="font-mono">{n.name}</span>
                  <span className="truncate font-mono text-muted-foreground">{n.ref}</span>
                  <button
                    aria-label={`이름 ${n.name} 삭제`}
                    className="rounded p-0.5 hover:bg-destructive/10 hover:text-destructive"
                    onClick={(e) => {
                      e.stopPropagation();
                      store().removeName(n.name);
                    }}
                  >
                    <TrashSimple className="size-3" />
                  </button>
                </DropdownMenuItem>
              ))
            )}
          </DropdownMenuContent>
        </DropdownMenu>
      </div>
      <span className="px-1 font-serif italic text-muted-foreground">fx</span>
      <input
        ref={inputRef}
        aria-label="수식 입력줄"
        className="h-6 min-w-0 flex-1 rounded border px-1.5 font-mono disabled:bg-muted/50"
        disabled={!selection}
        readOnly={locked}
        title={locked ? "Python 출력(spill) 셀 — 블록 코드에서 바꾸세요" : undefined}
        value={value}
        onChange={(e) => {
          setDraft(e.target.value);
          setCaret(e.target.selectionStart ?? e.target.value.length);
          setPick(0);
        }}
        onSelect={(e) => setCaret((e.target as HTMLInputElement).selectionStart ?? 0)}
        onBlur={() => {
          // 후보 클릭(mousedown)이 blur보다 먼저 처리되도록 지연
          setTimeout(() => setDraft((d) => (document.activeElement === inputRef.current ? d : null)), 150);
        }}
        onKeyDown={(e) => {
          if (suggestions.length > 0 && (e.key === "ArrowDown" || e.key === "ArrowUp")) {
            e.preventDefault();
            setPick((p) => (p + (e.key === "ArrowDown" ? 1 : -1) + suggestions.length) % suggestions.length);
            return;
          }
          if (suggestions.length > 0 && e.key === "Tab") {
            e.preventDefault();
            accept(suggestions[pick] ?? suggestions[0]);
            return;
          }
          if (e.key === "Enter") {
            e.preventDefault();
            commit();
          } else if (e.key === "Escape") {
            setDraft(null);
          }
        }}
      />
      {draft !== null && (suggestions.length > 0 || argHelp) && (
        <div className="absolute left-40 top-8 z-30 min-w-72 rounded border bg-popover text-popover-foreground shadow-md">
          {argHelp && suggestions.length === 0 && (
            <p className="px-2 py-1 font-mono text-[11px] text-muted-foreground">{argHelp}</p>
          )}
          {suggestions.map((sg, i) => (
            <button
              key={sg.text}
              type="button"
              className={`flex w-full items-baseline gap-2 px-2 py-1 text-left ${i === pick ? "bg-accent" : ""}`}
              onMouseDown={(e) => {
                e.preventDefault();
                accept(sg);
              }}
            >
              <span className="font-mono font-semibold">{sg.text}</span>
              <span className="truncate text-[11px] text-muted-foreground">{sg.help}</span>
            </button>
          ))}
          {suggestions.length > 0 && (
            <p className="border-t px-2 py-0.5 text-[10px] text-muted-foreground">Tab 선택 · ↑↓ 이동</p>
          )}
        </div>
      )}
    </div>
  );
}
