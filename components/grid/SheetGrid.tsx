"use client";

// Glide Data Grid 래퍼 — 프로젝트에서 glide-data-grid를 import하는 유일한 파일

import "@glideapps/glide-data-grid/dist/index.css";
import {
  CompactSelection,
  DataEditor,
  GridCellKind,
  type CellClickedEventArgs,
  type DataEditorRef,
  type DrawCellCallback,
  type EditableGridCell,
  type FillPatternEventArgs,
  type GridCell,
  type GridColumn,
  type GridMouseEventArgs,
  type GridSelection,
  type Item,
  type Theme,
} from "@glideapps/glide-data-grid";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { toast } from "sonner";
import { applyFill } from "@/components/grid/PasteImportDialog";
import {
  ContextMenu,
  ContextMenuContent,
  ContextMenuItem,
  ContextMenuSeparator,
  ContextMenuTrigger,
} from "@/components/ui/context-menu";
import { colToLetter, formatA1 } from "@/lib/grid/a1";
import { notifyWorkbookEdit } from "@/lib/grid/calc-host";
import { formatCellDisplay } from "@/lib/grid/format";
import { cellFromInput, FORMULA_ERROR_KO, isFormulaError } from "@/lib/grid/formula";
import { useWorkbookStore } from "@/lib/grid/model";
import { dataEdge, type Dir } from "@/lib/grid/navigate";
import { outputsOf, srcBlockId } from "@/lib/grid/outputs";
import { applyAnchorPick } from "@/lib/grid/run-block";
import {
  cellKey,
  type CellRange,
  type OutputBinding,
  type PyBlock,
} from "@/types/workbook";

/** 앵커 셀에 놓인 출력 (배지·객체 카드·컨텍스트 메뉴 대상) */
interface AnchorEntry {
  block: PyBlock;
  output?: OutputBinding;
}

/** §3.4: spill 잠금 안내 문구 */
function spillLockMessage(src: string): string {
  const st = useWorkbookStore.getState();
  const block = st.workbook.pyBlocks.find((b) => b.id === srcBlockId(src));
  if (!block) return "Python 블록의 결과입니다. 코드를 수정하거나 블록을 삭제하세요";
  const sheetName = st.workbook.sheets.find((s) => s.id === block.sheetId)?.name ?? "?";
  const addr = formatA1({
    r0: block.anchor.r,
    c0: block.anchor.c,
    r1: block.anchor.r,
    c1: block.anchor.c,
  });
  return `블록 ${sheetName}!${addr}의 결과입니다. 코드를 수정하거나 블록을 삭제하세요`;
}

const EMPTY_SELECTION: GridSelection = {
  columns: CompactSelection.empty(),
  rows: CompactSelection.empty(),
  current: undefined,
};

const DEFAULT_COL_WIDTH = 88;
/** 부록 P.6: 작게 보기(보조 화면) / 기본 크기 — 전체 화면에서는 항상 기본 크기 */
const DENSITY = {
  compact: { colWidth: 72, rowHeight: 24, headerHeight: 26, font: 12, headerFont: "600 11px", pad: 2 },
  normal: { colWidth: DEFAULT_COL_WIDTH, rowHeight: 34, headerHeight: 36, font: 12, headerFont: "600 11px", pad: 3 },
} as const;
const PRIMARY = "#4A90C2";
const WARNING = "#D9A441";
const DESTRUCTIVE = "#C2504A";
const MUTED = "#6B7280";

const inRange = (rg: CellRange, r: number, c: number): boolean =>
  r >= rg.r0 && r <= rg.r1 && c >= rg.c0 && c <= rg.c1;

export default function SheetGrid() {
  const sheet = useWorkbookStore(
    (s) => s.workbook.sheets.find((sh) => sh.id === s.activeSheetId) ?? s.workbook.sheets[0],
  );
  const pyBlocks = useWorkbookStore((s) => s.workbook.pyBlocks);
  const runningBlocks = useWorkbookStore((s) => s.runningBlocks);
  const dens = useWorkbookStore((s) => (s.gridCompact && !s.gridMaximized ? DENSITY.compact : DENSITY.normal));
  const flash = useWorkbookStore((s) => s.flash);
  const [gridSelection, setGridSelection] = useState<GridSelection>(EMPTY_SELECTION);
  const editorRef = useRef<DataEditorRef>(null);
  const [menuCell, setMenuCell] = useState<{ r: number; c: number } | null>(null);
  const [hoverTip, setHoverTip] = useState<{ x: number; y: number; text: string } | null>(null);

  // 시트 전환 시 선택 초기화
  useEffect(() => {
    setGridSelection(EMPTY_SELECTION);
  }, [sheet.id]);

  // 스토어 → glide 선택 동기화 (앵커로 이동·Ctrl+1 기본 선택 등 외부 변경 반영)
  const storeSelection = useWorkbookStore((s) => s.selection);
  useEffect(() => {
    if (!storeSelection) return;
    const cur = gridSelection.current?.range;
    const same =
      cur &&
      cur.x === storeSelection.c0 &&
      cur.y === storeSelection.r0 &&
      cur.width === storeSelection.c1 - storeSelection.c0 + 1 &&
      cur.height === storeSelection.r1 - storeSelection.r0 + 1;
    if (!same) {
      setGridSelection({
        columns: CompactSelection.empty(),
        rows: CompactSelection.empty(),
        current: {
          cell: [storeSelection.c0, storeSelection.r0],
          range: {
            x: storeSelection.c0,
            y: storeSelection.r0,
            width: storeSelection.c1 - storeSelection.c0 + 1,
            height: storeSelection.r1 - storeSelection.r0 + 1,
          },
          rangeStack: [],
        },
      });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [storeSelection]); // gridSelection 의도적 제외 — glide발 변경은 same으로 걸러져 루프 없음

  // 이름 상자 이동만 그 셀로 스크롤한다 — 다른 경로(블록 실행·앵커 이동)의 선택 변경은 화면을 움직이지 않는다
  useEffect(() => {
    const onGo = (e: Event) => {
      const { r, c } = (e as CustomEvent<{ r: number; c: number }>).detail;
      editorRef.current?.scrollTo(c, r);
    };
    window.addEventListener("pygrid:goto", onGo);
    return () => window.removeEventListener("pygrid:goto", onGo);
  }, []);

  // canvas는 CSS 변수를 못 쓰므로 body에서 실제 폰트 패밀리를 읽는다
  const [fontFamily, setFontFamily] = useState("Pretendard, sans-serif");
  useEffect(() => {
    const f = getComputedStyle(document.body).fontFamily;
    if (f) setFontFamily(f);
  }, []);

  /** 활성 시트의 앵커 → 블록·출력 (출력마다 앵커가 있다 — 부록 D.1) */
  const anchorMap = useMemo(() => {
    const map = new Map<string, AnchorEntry>();
    for (const b of pyBlocks) {
      if (b.kind === "markdown") {
        if (b.sheetId === sheet.id) map.set(cellKey(b.anchor.r, b.anchor.c), { block: b });
        continue;
      }
      for (const o of outputsOf(b)) {
        if ((o.sheetId ?? b.sheetId) !== sheet.id) continue;
        map.set(cellKey(o.anchor.r, o.anchor.c), { block: b, output: o });
      }
    }
    return map;
  }, [pyBlocks, sheet.id]);

  const blockById = useMemo(
    () => new Map(pyBlocks.map((b) => [b.id, b])),
    [pyBlocks],
  );

  /** 활성 시트의 spill 테두리 범위 — 블록마다 출력 수만큼 */
  const spillRanges = useMemo(() => {
    const out: CellRange[] = [];
    for (const b of pyBlocks) {
      for (const o of outputsOf(b)) {
        if ((o.sheetId ?? b.sheetId) !== sheet.id) continue;
        if (o.last?.status === "ok" && o.last.spillRange) out.push(o.last.spillRange);
      }
    }
    return out;
  }, [pyBlocks, sheet.id]);

  // 부록 J.3: 성공 실행이 읽은 xl() 참조 범위 (teal tint, 토글 가능)
  const executedRefs = useWorkbookStore((s) => s.executedRefs);
  const showRefs = useWorkbookStore((s) => s.showRefs);
  const refRanges = useMemo(() => {
    if (!showRefs) return [];
    const out: CellRange[] = [];
    for (const ranges of Object.values(executedRefs)) {
      for (const rg of ranges) if (rg.sheetId === sheet.id) out.push(rg);
    }
    return out;
  }, [executedRefs, showRefs, sheet.id]);

  // 부록 L.2: 채팅 에이전트가 읽은 근거 범위 (amber tint)
  const chatRefs = useWorkbookStore((s) => s.chatRefs);
  const chatRanges = useMemo(
    () => chatRefs.filter((rg) => rg.sheetId === sheet.id) as CellRange[],
    [chatRefs, sheet.id],
  );

  // drawCell만 바뀌면 glide가 다시 그리지 않는다 — 참조 tint 영역(이전+현재)을 직접 damage
  const prevRefRanges = useRef<CellRange[]>([]);
  useEffect(() => {
    const cells: { cell: Item }[] = [];
    for (const rg of [...prevRefRanges.current, ...refRanges, ...chatRanges]) {
      for (let r = rg.r0; r <= rg.r1 && cells.length < 20_000; r++) {
        for (let c = rg.c0; c <= rg.c1; c++) cells.push({ cell: [c, r] });
      }
    }
    // ponytail: 20k 셀 상한 — 초대형 참조는 스크롤 시 자연 리페인트로 마저 그린다
    prevRefRanges.current = [...refRanges, ...chatRanges];
    if (cells.length > 0) editorRef.current?.updateCells(cells);
  }, [refRanges, chatRanges]);

  const flashRange = flash && flash.sheetId === sheet.id ? flash.range : null;
  const hoverRange = useWorkbookStore((s) => s.hoverRange);
  const editorHover = hoverRange && hoverRange.sheetId === sheet.id ? hoverRange.range : null;

  const theme = useMemo<Partial<Theme>>(
    () => ({
      accentColor: PRIMARY,
      accentLight: "#EAF3FA",
      borderColor: "#E2E5E9",
      horizontalBorderColor: "#E2E5E9",
      bgCell: "#FFFFFF",
      bgHeader: "#F1F3F5",
      bgHeaderHovered: "#E9ECEF",
      bgHeaderHasFocus: "#EAF3FA",
      textDark: "#23272E",
      textHeader: "#5A6472",
      textLight: "#6B7280",
      fontFamily,
      baseFontStyle: `${dens.font}px`,
      headerFontStyle: dens.headerFont,
      cellHorizontalPadding: 8,
      cellVerticalPadding: dens.pad,
    }),
    [fontFamily, dens],
  );

  const columns = useMemo<GridColumn[]>(
    () =>
      Array.from({ length: sheet.colCount }, (_, i) => ({
        id: String(i),
        title: colToLetter(i),
        width: sheet.colWidths?.[i] ?? dens.colWidth,
      })),
    [sheet.colCount, sheet.colWidths, dens],
  );

  const getCellContent = useCallback(
    (item: Item): GridCell => {
      const [col, row] = item;
      const key = cellKey(row, col);
      const cell = sheet.cells[key];
      const anchor = anchorMap.get(key);
      const anchorBlock = anchor?.block;

      // 실행 중 앵커: #BUSY! (렌더 전용 — 스토어 셀은 건드리지 않는다)
      if (anchorBlock && runningBlocks[anchorBlock.id]) {
        return {
          kind: GridCellKind.Text,
          data: "#BUSY!",
          displayData: "#BUSY!",
          allowOverlay: false,
          readonly: true,
          contentAlign: "center",
          themeOverride: { textDark: WARNING },
        };
      }
      if (!cell) {
        return { kind: GridCellKind.Text, data: "", displayData: "", allowOverlay: true };
      }
      const locked = !!cell.src;
      // 객체 카드 앵커: drawCell이 카드를 그린다 (기본 텍스트는 비움)
      if (
        anchor?.output?.mode === "object" &&
        anchor.output.last?.status === "ok" &&
        cell.src === `${anchor.block.id}:${anchor.output.id}`
      ) {
        return {
          kind: GridCellKind.Text,
          data: String(cell.v ?? ""),
          displayData: "",
          allowOverlay: false,
          readonly: true,
        };
      }
      const contentAlign =
        cell.t === "n" ? "right" : cell.t === "b" || cell.t === "d" ? "center" : "left";
      let themeOverride: Partial<Theme> | undefined =
        cell.t === "e"
          ? { textDark: DESTRUCTIVE }
          : locked
            ? { bgCell: "#EAF3FA" }
            : undefined;
      // 부록 J.2: 셀 서식 (굵게·크기) — 폰트 패밀리는 theme이 유지
      if (cell.st) {
        themeOverride = {
          ...themeOverride,
          baseFontStyle: `${cell.st.b ? "bold " : ""}${cell.st.fs ?? dens.font}px`,
        };
      }
      return {
        kind: GridCellKind.Text,
        // 수식 셀은 편집 시드가 수식 원문 (부록 I.3 — 편집 재진입 시 원문 표시)
        data:
          cell.fx ??
          (typeof cell.v === "boolean" ? (cell.v ? "TRUE" : "FALSE") : String(cell.v ?? "")),
        displayData: formatCellDisplay(cell),
        // 잠긴(src) 셀도 편집 시도는 허용 — 커밋 시 스토어가 거부하고 안내 toast (§3.4)
        allowOverlay: true,
        contentAlign,
        themeOverride,
      };
    },
    [sheet, anchorMap, runningBlocks, dens],
  );

  /** 오버레이: [PY] 배지, spill 테두리, 성공 플래시, 객체 카드 */
  const drawCell = useCallback<DrawCellCallback>(
    (args, drawContent) => {
      const { ctx, rect, col, row } = args;
      const key = cellKey(row, col);
      const anchor = anchorMap.get(key);
      const anchorBlock = anchor?.block;
      const storeCell = sheet.cells[key];
      const isCard =
        anchor?.output?.mode === "object" &&
        anchor.output.last?.status === "ok" &&
        storeCell?.src === `${anchor.block.id}:${anchor.output.id}` &&
        !runningBlocks[anchor.block.id];

      drawContent();

      if (isCard && storeCell) {
        // 객체 카드: 연한 배경 + 아이콘 + 모노 라벨
        ctx.save();
        ctx.fillStyle = "#EAF3FA";
        ctx.fillRect(rect.x + 1, rect.y + 1, rect.width - 2, rect.height - 2);
        ctx.strokeStyle = PRIMARY;
        ctx.lineWidth = 1;
        ctx.strokeRect(rect.x + 2.5, rect.y + 2.5, rect.width - 5, rect.height - 5);
        ctx.fillStyle = "#23272E";
        ctx.font = `11px "JetBrains Mono", monospace`;
        ctx.textBaseline = "middle";
        ctx.fillText(
          `▤ ${String(storeCell.v ?? "")}`,
          rect.x + 6,
          rect.y + rect.height / 2,
          rect.width - 12,
        );
        ctx.restore();
      }

      // spill 파란 테두리 (범위 경계 변만)
      for (const rg of spillRanges) {
        if (!inRange(rg, row, col)) continue;
        ctx.save();
        ctx.strokeStyle = PRIMARY;
        ctx.lineWidth = 1.5;
        ctx.beginPath();
        if (row === rg.r0) {
          ctx.moveTo(rect.x, rect.y + 1);
          ctx.lineTo(rect.x + rect.width, rect.y + 1);
        }
        if (row === rg.r1) {
          ctx.moveTo(rect.x, rect.y + rect.height - 1);
          ctx.lineTo(rect.x + rect.width, rect.y + rect.height - 1);
        }
        if (col === rg.c0) {
          ctx.moveTo(rect.x + 1, rect.y);
          ctx.lineTo(rect.x + 1, rect.y + rect.height);
        }
        if (col === rg.c1) {
          ctx.moveTo(rect.x + rect.width - 1, rect.y);
          ctx.lineTo(rect.x + rect.width - 1, rect.y + rect.height);
        }
        ctx.stroke();
        ctx.restore();
      }

      // 부록 J.3: 실행 참조 tint(teal) + 부록 L.2: 채팅이 읽은 근거 tint(amber)
      // — 둘 다 배경 + 가장자리 얇은 점선 (spill Sky Blue와 구별)
      for (const [ranges, fill, stroke] of [
        [refRanges, "rgba(31, 110, 100, 0.08)", "#1F6E64"], // --chip-teal-fg 8%
        [chatRanges, "rgba(125, 90, 20, 0.12)", "#7D5A14"], // --chip-amber-fg 12%
      ] as const) {
      for (const rg of ranges) {
        if (!inRange(rg, row, col)) continue;
        ctx.save();
        ctx.fillStyle = fill;
        ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
        ctx.strokeStyle = stroke;
        ctx.lineWidth = 1;
        ctx.setLineDash([3, 3]);
        ctx.beginPath();
        if (row === rg.r0) {
          ctx.moveTo(rect.x, rect.y + 0.5);
          ctx.lineTo(rect.x + rect.width, rect.y + 0.5);
        }
        if (row === rg.r1) {
          ctx.moveTo(rect.x, rect.y + rect.height - 0.5);
          ctx.lineTo(rect.x + rect.width, rect.y + rect.height - 0.5);
        }
        if (col === rg.c0) {
          ctx.moveTo(rect.x + 0.5, rect.y);
          ctx.lineTo(rect.x + 0.5, rect.y + rect.height);
        }
        if (col === rg.c1) {
          ctx.moveTo(rect.x + rect.width - 0.5, rect.y);
          ctx.lineTo(rect.x + rect.width - 0.5, rect.y + rect.height);
        }
        ctx.stroke();
        ctx.restore();
      }
      }

      // 성공 플래시 (400ms)
      if (flashRange && inRange(flashRange, row, col)) {
        ctx.save();
        ctx.fillStyle = "rgba(74, 144, 194, 0.18)";
        ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
        ctx.restore();
      }

      // 편집기 xl() 커서 → 점선 하이라이트 (§4.8)
      if (editorHover && inRange(editorHover, row, col)) {
        ctx.save();
        ctx.strokeStyle = PRIMARY;
        ctx.lineWidth = 1.5;
        ctx.setLineDash([4, 3]);
        ctx.beginPath();
        if (row === editorHover.r0) {
          ctx.moveTo(rect.x, rect.y + 1);
          ctx.lineTo(rect.x + rect.width, rect.y + 1);
        }
        if (row === editorHover.r1) {
          ctx.moveTo(rect.x, rect.y + rect.height - 1);
          ctx.lineTo(rect.x + rect.width, rect.y + rect.height - 1);
        }
        if (col === editorHover.c0) {
          ctx.moveTo(rect.x + 1, rect.y);
          ctx.lineTo(rect.x + 1, rect.y + rect.height);
        }
        if (col === editorHover.c1) {
          ctx.moveTo(rect.x + rect.width - 1, rect.y);
          ctx.lineTo(rect.x + rect.width - 1, rect.y + rect.height);
        }
        ctx.stroke();
        ctx.restore();
      }

      // 앵커 배지 (우상단 칩) — 코드는 [PY], 마크다운은 [§](Python 관여가 아니라 muted)
      if (anchorBlock) {
        const isMd = anchorBlock.kind === "markdown";
        ctx.save();
        const w = isMd ? 12 : 22;
        const h = 11;
        const x = rect.x + rect.width - w - 2;
        const y = rect.y + 2;
        ctx.fillStyle = isMd ? MUTED : PRIMARY;
        ctx.beginPath();
        ctx.roundRect(x, y, w, h, 2);
        ctx.fill();
        ctx.fillStyle = "#FFFFFF";
        ctx.font = `700 ${isMd ? 9 : 7}px "JetBrains Mono", monospace`;
        ctx.textAlign = "center";
        ctx.textBaseline = "middle";
        ctx.fillText(isMd ? "§" : "PY", x + w / 2, y + h / 2 + 0.5);
        ctx.restore();
      }
    },
    [
      anchorMap,
      sheet.cells,
      spillRanges,
      refRanges,
      chatRanges,
      flashRange,
      editorHover,
      runningBlocks,
    ],
  );

  const onCellEdited = useCallback(
    (item: Item, newValue: EditableGridCell) => {
      if (newValue.kind !== GridCellKind.Text) return;
      const [col, row] = item;
      const store = useWorkbookStore.getState();
      const ok = store.setCellValue(sheet.id, row, col, cellFromInput(newValue.data));
      if (!ok) {
        const src = sheet.cells[cellKey(row, col)]?.src;
        if (src) toast.error(spillLockMessage(src), { id: "spill-lock" });
        return;
      }
      notifyWorkbookEdit([{ sheetId: sheet.id, r0: row, c0: col, r1: row, c1: col }]);
    },
    [sheet],
  );

  const onGridSelectionChange = useCallback(
    (sel: GridSelection) => {
      setGridSelection(sel);
      const store = useWorkbookStore.getState();
      if (sel.current) {
        const r = sel.current.range;
        store.setSelection({ r0: r.y, c0: r.x, r1: r.y + r.height - 1, c1: r.x + r.width - 1 });
      } else if (sel.columns.length > 0) {
        const first = sel.columns.first();
        const last = sel.columns.last();
        if (first !== undefined && last !== undefined) {
          store.setSelection({ r0: 0, c0: first, r1: sheet.rowCount - 1, c1: last });
        }
      } else if (sel.rows.length > 0) {
        const first = sel.rows.first();
        const last = sel.rows.last();
        if (first !== undefined && last !== undefined) {
          store.setSelection({ r0: first, c0: 0, r1: last, c1: sheet.colCount - 1 });
        }
      } else {
        store.setSelection(null);
      }
    },
    [sheet.rowCount, sheet.colCount],
  );

  /** 앵커(활성 셀) 고정 + 반대 코너로 사각 선택 — glide·스토어 동시 갱신 (Ctrl+방향키·가장자리 점프 공용) */
  const applySelection = useCallback(
    (anchor: { r: number; c: number }, corner: { r: number; c: number }) => {
      const r0 = Math.min(anchor.r, corner.r);
      const r1 = Math.max(anchor.r, corner.r);
      const c0 = Math.min(anchor.c, corner.c);
      const c1 = Math.max(anchor.c, corner.c);
      setGridSelection({
        columns: CompactSelection.empty(),
        rows: CompactSelection.empty(),
        current: {
          cell: [anchor.c, anchor.r],
          range: { x: c0, y: r0, width: c1 - c0 + 1, height: r1 - r0 + 1 },
          rangeStack: [],
        },
      });
      useWorkbookStore.getState().setSelection({ r0, c0, r1, c1 });
      editorRef.current?.scrollTo(corner.c, corner.r);
    },
    [],
  );

  /** 현재 선택에서 앵커·이동 코너를 dir 방향 dataEdge로 확장한 코너 계산 (엑셀 Ctrl+Shift+방향키) */
  const extendCorner = useCallback(
    (
      anchor: { r: number; c: number },
      rg: { x: number; y: number; width: number; height: number },
      dir: Dir,
    ): { r: number; c: number } => {
      // 이동 코너 = 앵커의 반대쪽 코너
      const corner = {
        r: anchor.r === rg.y ? rg.y + rg.height - 1 : rg.y,
        c: anchor.c === rg.x ? rg.x + rg.width - 1 : rg.x,
      };
      // 엑셀처럼 점프 기준은 앵커의 행/열에서 계산한다
      return dir === "up" || dir === "down"
        ? { r: dataEdge(sheet, { r: corner.r, c: anchor.c }, dir).r, c: corner.c }
        : { r: corner.r, c: dataEdge(sheet, { r: anchor.r, c: corner.c }, dir).c };
    },
    [sheet],
  );

  const onDelete = useCallback(
    (sel: GridSelection): boolean => {
      const store = useWorkbookStore.getState();
      // ponytail: 행/열 헤더 선택은 행·열마다 clearRange 한 번 — 대량 선택이 느려지면 단일 패스로
      const cleared: { r0: number; c0: number; r1: number; c1: number }[] = [];
      if (sel.current) {
        for (const r of [sel.current.range, ...sel.current.rangeStack]) {
          const range = { r0: r.y, c0: r.x, r1: r.y + r.height - 1, c1: r.x + r.width - 1 };
          store.clearRange(sheet.id, range);
          cleared.push(range);
        }
      }
      for (const r of sel.rows) {
        store.clearRange(sheet.id, { r0: r, c0: 0, r1: r, c1: sheet.colCount - 1 });
        cleared.push({ r0: r, c0: 0, r1: r, c1: sheet.colCount - 1 });
      }
      for (const c of sel.columns) {
        store.clearRange(sheet.id, { r0: 0, c0: c, r1: sheet.rowCount - 1, c1: c });
        cleared.push({ r0: 0, c0: c, r1: sheet.rowCount - 1, c1: c });
      }
      if (cleared.length > 0) {
        notifyWorkbookEdit(cleared.map((r) => ({ ...r, sheetId: sheet.id })));
      }
      return false; // 삭제는 스토어에서 처리했으므로 glide 기본 동작 차단
    },
    [sheet.id, sheet.rowCount, sheet.colCount],
  );

  /** 부록 O.4 채우기 핸들: glide 기본 복사 대신 연속 패턴(수식 이동·숫자 추세·끝 숫자 증가) */
  const onFillPattern = useCallback((e: FillPatternEventArgs) => {
    e.preventDefault();
    const rg = (x: { x: number; y: number; width: number; height: number }) => ({
      r0: x.y,
      c0: x.x,
      r1: x.y + x.height - 1,
      c1: x.x + x.width - 1,
    });
    applyFill(rg(e.patternSource), rg(e.fillDestination));
  }, []);

  const onColumnResize = useCallback(
    (_column: GridColumn, newSize: number, colIndex: number) => {
      useWorkbookStore.getState().setColWidth(sheet.id, colIndex, newSize);
    },
    [sheet.id],
  );

  // glide의 isDoubleClick은 시간만 본다(500ms 내 아무 클릭) — 같은 셀 재클릭인지 직접 확인
  const lastClickCell = useRef<{ r: number; c: number } | null>(null);

  /** 클릭 라우팅: 출력 위치 지정 → 앵커 이동, Shift+더블클릭 → 데이터 끝 확장,
   *  오류 셀 → 진단 탭(§2.3.7), 객체 카드 → 출력 미리보기 탭(§2.3.4) */
  const onCellClicked = useCallback(
    (item: Item, ev: CellClickedEventArgs) => {
      const [col, row] = item;
      const prevClick = lastClickCell.current;
      lastClickCell.current = { r: row, c: col };
      const picking = useWorkbookStore.getState().anchorPicking;
      if (picking) {
        applyAnchorPick(picking, sheet.id, { r: row, c: col });
        return;
      }
      const sameAsPrev = prevClick?.r === row && prevClick?.c === col;
      // glide isDoubleClick은 시간만 판정(500ms 내 아무 클릭) — 직전 클릭이 다른 셀이면
      // 진짜 더블클릭이 아니므로 편집기 활성화를 막는다. Enter로 선택이 이동한 직후
      // 다음 셀을 빠르게 클릭+타이핑할 때 오탐 활성화가 첫 입력을 유실시키는 것 방지.
      if (ev.isDoubleClick === true && !sameAsPrev) {
        ev.preventDefault(); // 활성화만 차단 — 아래 일반 클릭 라우팅은 그대로 수행
      }
      // 엑셀 선택 가장자리 더블클릭 점프의 셀 그리드 대응:
      // 수식어 없는 더블클릭은 항상 편집(glide 기본), Shift+더블클릭은 선택을 그 방향 데이터 끝까지 확장.
      if (ev.isDoubleClick === true && ev.shiftKey && sameAsPrev) {
        const cur = gridSelection.current;
        if (cur) {
          const anchor = { r: cur.cell[1], c: cur.cell[0] };
          const rg = cur.range;
          // 클릭 셀이 앵커 반대쪽 가장자리에 있는 방향 (세로·가로 둘 다면 앵커에서 먼 축)
          const cand: Dir[] = [];
          if (row === rg.y + rg.height - 1 && row > anchor.r) cand.push("down");
          if (row === rg.y && row < anchor.r) cand.push("up");
          if (col === rg.x + rg.width - 1 && col > anchor.c) cand.push("right");
          if (col === rg.x && col < anchor.c) cand.push("left");
          const dir =
            cand.length > 1
              ? Math.abs(row - anchor.r) >= Math.abs(col - anchor.c)
                ? cand.find((d) => d === "up" || d === "down")
                : cand.find((d) => d === "left" || d === "right")
              : cand[0];
          if (dir) {
            ev.preventDefault(); // 편집기 오픈 차단
            applySelection(anchor, extendCorner(anchor, rg, dir));
            return;
          }
        }
      }
      const key = cellKey(row, col);
      const cell = sheet.cells[key];
      if (!cell?.src) return;
      const store = useWorkbookStore.getState();
      if (cell.t === "e") {
        store.setSelectedBlock(srcBlockId(cell.src));
        store.showPanelTab("diagnostics");
        return;
      }
      const anchor = anchorMap.get(key);
      if (anchor?.output?.mode === "object" && anchor.output.last?.status === "ok") {
        store.setSelectedBlock(anchor.block.id);
        store.showPanelTab("preview");
      }
    },
    [sheet.cells, sheet.id, anchorMap, gridSelection, applySelection, extendCorner],
  );

  /** 우클릭: 셀 기록(+선택 이동). radix ContextMenu가 메뉴를 연다 */
  const onCellContextMenu = useCallback(
    (item: Item) => {
      const [col, row] = item;
      setMenuCell({ r: row, c: col });
      const store = useWorkbookStore.getState();
      if (!store.selection || !inRange(store.selection, row, col)) {
        store.setSelection({ r0: row, c0: col, r1: row, c1: col });
        setGridSelection({
          columns: CompactSelection.empty(),
          rows: CompactSelection.empty(),
          current: {
            cell: [col, row],
            range: { x: col, y: row, width: 1, height: 1 },
            rangeStack: [],
          },
        });
      }
    },
    [],
  );

  /** 오류 셀 hover → 한국어 요약 툴팁 · spill hover → 블록 카드 강조(§4.8) */
  const onItemHovered = useCallback(
    (args: GridMouseEventArgs) => {
      if (args.kind !== "cell") {
        setHoverTip(null);
        useWorkbookStore.getState().setHoverBlock(null);
        return;
      }
      const [col, row] = args.location;
      const cell = sheet.cells[cellKey(row, col)];
      useWorkbookStore.getState().setHoverBlock(cell?.src ? srcBlockId(cell.src) : null);
      // 수식 오류 셀 → 한국어 설명 (부록 I.3)
      if (cell?.fx && cell.t === "e" && isFormulaError(cell.v)) {
        setHoverTip({
          x: args.bounds.x,
          y: args.bounds.y + args.bounds.height + 4,
          text: FORMULA_ERROR_KO[cell.v],
        });
        return;
      }
      if (!cell?.src) {
        setHoverTip(null);
        return;
      }
      // 오류 셀 → 한국어 요약, 그 외 src(spill 잠금) 셀 → 잠금 안내 (§3.4)
      const text =
        (cell.t === "e"
          ? blockById.get(srcBlockId(cell.src))?.last?.summaryKo
          : undefined) ?? spillLockMessage(cell.src);
      setHoverTip({
        x: args.bounds.x,
        y: args.bounds.y + args.bounds.height + 4,
        text,
      });
    },
    [sheet.cells, blockById],
  );

  // 컨텍스트 메뉴 항목용 정보 (앵커면 그 출력, spill 셀이면 그 셀을 쓴 출력)
  const menuEntry: AnchorEntry | undefined = (() => {
    if (!menuCell) return undefined;
    const key = cellKey(menuCell.r, menuCell.c);
    const hit = anchorMap.get(key);
    if (hit) return hit;
    const src = sheet.cells[key]?.src;
    if (!src) return undefined;
    const block = blockById.get(srcBlockId(src));
    if (!block) return undefined;
    return { block, output: outputsOf(block).find((o) => `${block.id}:${o.id}` === src) };
  })();
  const menuBlock = menuEntry?.block;

  const copyRef = () => {
    const sel = useWorkbookStore.getState().selection;
    const range = sel ?? (menuCell ? { r0: menuCell.r, c0: menuCell.c, r1: menuCell.r, c1: menuCell.c } : null);
    if (range) void navigator.clipboard.writeText(formatA1(range));
  };

  // QA 발견(첫 글자 유실): 셀 클릭 직후 초고속 타이핑 시 오버레이 편집기 마운트 전에
  // 캔버스로 떨어진 키가 유실된다(glide는 마지막 키 하나로만 편집기를 시드).
  // 캡처 단계에서 인쇄 가능 키를 400ms 창으로 버퍼링했다가, 오버레이 textarea가 나타나
  // "마지막 키 하나"로 시드된 것이 확인되면 전체 버퍼로 보정한다.
  // IME(한글) 조합(isComposing·key.length>1)과 단축키는 관여하지 않는다.
  const typeBuf = useRef<{ chars: string[]; ts: number }>({ chars: [], ts: 0 });

  const bufferCanvasKey = useCallback((e: React.KeyboardEvent) => {
    const t = e.target as HTMLElement;
    if (t.tagName === "TEXTAREA" || t.tagName === "INPUT" || t.isContentEditable) return;
    if (e.key.length !== 1 || e.ctrlKey || e.metaKey || e.altKey || e.nativeEvent.isComposing)
      return;
    const buf = typeBuf.current;
    const now = performance.now();
    if (now - buf.ts > 400) buf.chars = [];
    buf.chars.push(e.key);
    buf.ts = now;
  }, []);

  /** Ctrl+방향키(데이터 끝 이동) / Ctrl+Shift+방향키(데이터 끝까지 확장) — glide로 전파 차단.
   *  나머지 키는 초고속 타이핑 버퍼로 넘긴다. 편집기(textarea·input) 오픈 중에는 불개입. */
  const onWrapperKeyDown = useCallback(
    (e: React.KeyboardEvent) => {
      const dir = (
        { ArrowUp: "up", ArrowDown: "down", ArrowLeft: "left", ArrowRight: "right" } as Record<
          string,
          Dir | undefined
        >
      )[e.key];
      if (dir && (e.ctrlKey || e.metaKey) && !e.altKey) {
        const t = e.target as HTMLElement;
        if (t.tagName === "TEXTAREA" || t.tagName === "INPUT" || t.isContentEditable) return;
        const cur = gridSelection.current;
        if (!cur) return;
        e.preventDefault();
        e.stopPropagation();
        const anchor = { r: cur.cell[1], c: cur.cell[0] };
        if (e.shiftKey) {
          applySelection(anchor, extendCorner(anchor, cur.range, dir));
        } else {
          const dest = dataEdge(sheet, anchor, dir);
          applySelection(dest, dest);
        }
        return;
      }
      bufferCanvasKey(e);
    },
    [gridSelection, sheet, applySelection, extendCorner, bufferCanvasKey],
  );

  useEffect(() => {
    const portal = document.getElementById("portal");
    if (!portal) return;
    const obs = new MutationObserver(() => {
      const buf = typeBuf.current;
      if (buf.chars.length < 2 || performance.now() - buf.ts > 400) return;
      const ta = portal.querySelector("textarea, input") as
        | HTMLTextAreaElement
        | HTMLInputElement
        | null;
      if (!ta) return;
      const want = buf.chars.join("");
      if (ta.value !== buf.chars[buf.chars.length - 1]) return; // 시드 형태가 아니면 불개입
      const proto =
        ta.tagName === "TEXTAREA" ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
      Object.getOwnPropertyDescriptor(proto, "value")?.set?.call(ta, want);
      ta.dispatchEvent(new Event("input", { bubbles: true }));
      ta.setSelectionRange(want.length, want.length);
      buf.chars = [];
    });
    obs.observe(portal, { childList: true, subtree: true });
    return () => obs.disconnect();
  }, []);

  return (
    <ContextMenu>
      <ContextMenuTrigger asChild>
        <div className="relative min-h-0 min-w-0 flex-1" onKeyDownCapture={onWrapperKeyDown}>
          <DataEditor
            ref={editorRef}
            columns={columns}
            rows={sheet.rowCount}
            rowHeight={dens.rowHeight}
            headerHeight={dens.headerHeight}
            getCellContent={getCellContent}
            onCellEdited={onCellEdited}
            gridSelection={gridSelection}
            onGridSelectionChange={onGridSelectionChange}
            onDelete={onDelete}
            fillHandle={true}
            onFillPattern={onFillPattern}
            onColumnResize={onColumnResize}
            onCellContextMenu={onCellContextMenu}
            onCellClicked={onCellClicked}
            onItemHovered={onItemHovered}
            drawCell={drawCell}
            freezeColumns={sheet.frozenCols ?? 0}
            rowMarkers="number"
            getCellsForSelection={true}
            keybindings={{ copy: false, cut: false, paste: false }} // 복사·붙여넣기는 WorkbookShell이 소유
            smoothScrollX
            smoothScrollY
            width="100%"
            height="100%"
            theme={theme}
          />
          {hoverTip && (
            <div
              className="pointer-events-none absolute z-20 max-w-72 rounded border bg-popover px-2 py-1 text-xs text-popover-foreground shadow-md"
              style={{ left: hoverTip.x, top: hoverTip.y }}
            >
              {hoverTip.text}
            </div>
          )}
        </div>
      </ContextMenuTrigger>
      <ContextMenuContent>
        {menuBlock && (
          <>
            {menuBlock.kind !== "markdown" && menuEntry?.output && (
              <ContextMenuItem
                onClick={() =>
                  useWorkbookStore
                    .getState()
                    .setOutputMode(
                      menuBlock.id,
                      menuEntry.output!.id,
                      menuEntry.output!.mode === "values" ? "object" : "values",
                    )
                }
              >
                Python 출력 →{" "}
                {menuEntry.output.mode === "values" ? "Python 객체" : "Excel 값"}
              </ContextMenuItem>
            )}
            <ContextMenuItem
              onClick={() => useWorkbookStore.getState().setFocusBlock(menuBlock.id)}
            >
              블록으로 이동
            </ContextMenuItem>
            <ContextMenuSeparator />
          </>
        )}
        <ContextMenuItem onClick={copyRef}>참조 복사</ContextMenuItem>
      </ContextMenuContent>
    </ContextMenu>
  );
}
