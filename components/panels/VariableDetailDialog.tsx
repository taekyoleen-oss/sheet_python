"use client";

// 부록 P: 변수 세부 내용 팝업 — 속성 창에서 변수를 누르면 열린다.
//  표(DataFrame·Series·배열): 열·형 + 상위 100행, 요약 통계(describe) 탭
//  모델: 시트로 보낼 수 있는 속성(현재 값·코드) 목록
//  그 밖의 값: repr 원문
// 런타임 객체 모드 실행(client.run)으로 미리보기를 받는다 — 셀에는 아무것도 쓰지 않는다.

import { useEffect, useLayoutEffect, useRef, useState, type ReactNode } from "react";
import { ChartLineUp, Table } from "@phosphor-icons/react";
import { Button } from "@/components/ui/button";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { getRuntimeClient } from "@/lib/runtime/client";
import type { PreviewPayload, VariableInfo } from "@/lib/runtime/protocol";
import { cn } from "@/lib/utils";

type Tab = "data" | "stats";

/** 객체 모드로 식 하나를 평가해 미리보기를 받는다 */
async function previewOf(code: string): Promise<PreviewPayload | { kind: "error"; message: string }> {
  const payload = await getRuntimeClient().run("__var_detail", code, {}, "object", "auto");
  if (!payload.ok) return { kind: "error", message: payload.message };
  return payload.preview ?? { kind: "repr", repr: "(미리보기 없음)" };
}

/** 보기 전용 소수 자리 — 값(데이터)은 그대로, 화면 표시만 바꾼다 */
const DIGITS = [0, 1, 2, 3, 4, 5, 6, 7, 8, 10] as const;
const DEFAULT_DIGITS = 5;

function DigitsSelect({ value, onChange, label, inherit }: { value: number | null; onChange: (d: number | null) => void; label: string; inherit?: boolean }) {
  return (
    <select
      aria-label={label}
      title={label}
      value={value ?? ""}
      onChange={(e) => onChange(e.target.value === "" ? null : Number(e.target.value))}
      className="rounded border bg-background px-1 py-0.5 font-sans text-[11px] font-normal"
    >
      {inherit && <option value="">전체</option>}
      {DIGITS.map((d) => (
        <option key={d} value={d}>
          {d}자리
        </option>
      ))}
    </select>
  );
}

/** 열 너비를 마우스로 — 처음 그려진 너비를 재서 고정 레이아웃으로 바꾸고, 머리 오른쪽 선을 끌어 조절 */
function ResizableTable({ head, children, className }: { head: ReactNode[]; children: ReactNode; className?: string }) {
  const ref = useRef<HTMLTableElement>(null);
  const [widths, setWidths] = useState<number[] | null>(null);
  useLayoutEffect(() => {
    if (widths || !ref.current) return;
    setWidths([...ref.current.querySelectorAll("thead th")].map((th) => (th as HTMLElement).offsetWidth));
  }, [widths]);
  const drag = (i: number) => (e: React.PointerEvent) => {
    if (!widths) return;
    e.preventDefault();
    const x0 = e.clientX;
    const w0 = widths[i];
    const move = (ev: PointerEvent) =>
      setWidths((w) => w && w.map((x, k) => (k === i ? Math.max(40, w0 + ev.clientX - x0) : x)));
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  };
  return (
    <table
      ref={ref}
      className={cn("border-collapse [&_td]:border-r [&_td]:border-border/60 [&_th]:border-r [&_th]:border-border/60", className)}
      style={widths ? { tableLayout: "fixed", width: widths.reduce((a, b) => a + b, 0) } : { width: "100%" }}
    >
      {widths && (
        <colgroup>
          {widths.map((w, i) => (
            <col key={i} style={{ width: w }} />
          ))}
        </colgroup>
      )}
      <thead className="sticky top-0 z-10 bg-muted">
        <tr>
          {head.map((h, i) => (
            <th key={i} className="relative border-b px-2 py-1 text-left align-top">
              {h}
              <span
                role="separator"
                aria-orientation="vertical"
                aria-label="열 너비 조절"
                onPointerDown={drag(i)}
                className="absolute top-0 -right-1 z-10 h-full w-2 cursor-col-resize hover:bg-primary/30"
              />
            </th>
          ))}
        </tr>
      </thead>
      {children}
    </table>
  );
}

function PreviewView({ p }: { p: PreviewPayload | { kind: "error"; message: string } | null }) {
  // 전체 소수 자리 + 열마다 덮어쓰기(null = 전체를 따름)
  const [digits, setDigits] = useState(DEFAULT_DIGITS);
  const [colDigits, setColDigits] = useState<Record<number, number | null>>({});
  if (!p) return <p className="py-6 text-center text-sm text-muted-foreground">불러오는 중…</p>;
  if (p.kind === "error") return <p className="text-sm text-destructive">{p.message}</p>;
  if (p.kind === "image") return <p className="text-sm text-muted-foreground">이미지 값입니다 — 시트에 객체로 놓아 보세요.</p>;
  const toolbar = (
    <label className="flex items-center gap-1.5 text-xs text-muted-foreground">
      소수 자리 <DigitsSelect value={digits} onChange={(d) => setDigits(d ?? DEFAULT_DIGITS)} label="소수 자리 (전체)" />
      <span>— 보기에만 적용, 데이터는 그대로</span>
    </label>
  );
  if (p.kind === "repr") {
    // 숫자 하나(float 스칼라)만 전체 소수 자리로 — 그 밖의 repr 원문은 그대로
    const n = /^-?\d+\.\d+(e[-+]?\d+)?$/i.test(p.repr.trim()) ? Number(p.repr) : null;
    return (
      <div className="grid gap-2">
        {n !== null && toolbar}
        <pre className="max-h-[55vh] overflow-auto rounded bg-muted/40 p-3 font-mono text-xs whitespace-pre">
          {n !== null ? n.toFixed(digits) : p.repr}
        </pre>
      </div>
    );
  }
  // 소수가 있는 열(float dtype 또는 정수가 아닌 숫자)만 자리수를 고를 수 있다
  const decimalCol = p.columns.map(
    (_, i) => /float|complex/.test(p.dtypes[i] ?? "") || p.rows.some((r) => typeof r[i] === "number" && !Number.isInteger(r[i])),
  );
  const show = (v: string | number | boolean | null, c: number) =>
    typeof v === "number" && decimalCol[c] && Number.isFinite(v) ? v.toFixed(colDigits[c] ?? digits) : String(v);
  return (
    <div className="grid min-w-0 gap-2">
      {decimalCol.some(Boolean) && toolbar}
      <div className="max-h-[55vh] overflow-auto rounded border">
        <ResizableTable
          className="font-mono text-xs tabular-nums [&_td]:overflow-hidden [&_td]:text-ellipsis [&_td]:whitespace-nowrap"
          head={[
            <span key="#" className="block text-right font-normal text-muted-foreground">#</span>,
            ...p.columns.map((c, i) => (
              <div key={i} className="overflow-hidden">
                <div className="truncate" title={c}>{c}</div>
                <div className="truncate font-normal text-muted-foreground">{p.dtypes[i]}</div>
                {decimalCol[i] && (
                  <DigitsSelect
                    inherit
                    value={colDigits[i] ?? null}
                    onChange={(d) => setColDigits((m) => ({ ...m, [i]: d }))}
                    label={`${c} 소수 자리`}
                  />
                )}
              </div>
            )),
          ]}
        >
          <tbody>
            {p.rows.map((row, r) => (
              <tr key={r} className="border-b last:border-0">
                <td className="px-2 py-0.5 text-right text-muted-foreground">{r}</td>
                {row.map((v, c) => (
                  <td key={c} className={cn("px-2 py-0.5", typeof v === "number" ? "text-right" : "text-left")}>
                    {v === null ? <span className="text-muted-foreground">NaN</span> : show(v, c)}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </ResizableTable>
      </div>
    </div>
  );
}

export default function VariableDetailDialog({
  info,
  onClose,
  onShow,
  onModel,
}: {
  info: VariableInfo | null;
  onClose: () => void;
  onShow: (v: VariableInfo) => void;
  onModel: (v: VariableInfo) => void;
}) {
  const [tab, setTab] = useState<Tab>("data");
  const [data, setData] = useState<Awaited<ReturnType<typeof previewOf>> | null>(null);
  const [stats, setStats] = useState<Awaited<ReturnType<typeof previewOf>> | null>(null);
  const tabular = !!info && (!!info.columns || info.type === "Series");

  useEffect(() => {
    setTab("data");
    setData(null);
    setStats(null);
    if (!info || info.model) return;
    void previewOf(info.name).then(setData).catch((e) => setData({ kind: "error", message: String(e) }));
  }, [info]);

  useEffect(() => {
    if (!info || tab !== "stats" || stats) return;
    void previewOf(`${info.name}.describe(include="all").T`)
      .then(setStats)
      .catch((e) => setStats({ kind: "error", message: String(e) }));
  }, [info, tab, stats]);

  if (!info) return null;
  const shape = info.shape ? `${info.shape[0]}행 × ${info.shape[1]}열` : "";

  return (
    <Dialog open onOpenChange={(o) => !o && onClose()}>
      <DialogContent className="max-h-[88vh] overflow-hidden sm:max-w-4xl" data-testid="variable-detail">
        <DialogHeader>
          <DialogTitle className="font-mono">{info.name}</DialogTitle>
          <DialogDescription>
            {info.type}
            {shape ? ` · ${shape}` : ""}
            {info.model ? ` · ${info.model.kind} 모델 — 속성 ${info.model.members.length}개` : ""}
          </DialogDescription>
        </DialogHeader>

        {info.model ? (
          <div className="max-h-[55vh] overflow-auto rounded border">
            <ResizableTable
              className="text-xs"
              head={["묶음", "항목", "현재 값", "코드"].map((h) => (
                <span key={h} className="font-normal text-muted-foreground">{h}</span>
              ))}
            >
              <tbody>
                {info.model.members.map((m) => (
                  <tr key={m.expr} className="border-t align-top">
                    <td className="px-2 py-1 text-muted-foreground">{m.group}</td>
                    <td className="px-2 py-1">{m.label}</td>
                    <td className="truncate px-2 py-1 font-mono" title={m.preview}>
                      {m.preview}
                    </td>
                    <td className="px-2 py-1 font-mono text-muted-foreground">{m.code}</td>
                  </tr>
                ))}
              </tbody>
            </ResizableTable>
          </div>
        ) : (
          <div className="grid min-w-0 gap-2">
            {tabular && (
              <div role="tablist" className="flex gap-1 text-xs">
                {(
                  [
                    ["data", "데이터 (상위 100행)"],
                    ["stats", "요약 통계"],
                  ] as const
                ).map(([id, label]) => (
                  <button
                    key={id}
                    role="tab"
                    aria-selected={tab === id}
                    onClick={() => setTab(id)}
                    className={cn(
                      "rounded border-b-2 px-2 py-1",
                      tab === id ? "border-primary font-medium" : "border-transparent text-muted-foreground",
                    )}
                  >
                    {label}
                  </button>
                ))}
              </div>
            )}
            <PreviewView key={`${info.name}:${tab}`} p={tab === "data" ? data : stats} />
          </div>
        )}

        <DialogFooter className="gap-2">
          {info.model && (
            <Button variant="outline" onClick={() => onModel(info)}>
              <ChartLineUp /> 모델 결과 → 시트
            </Button>
          )}
          <Button onClick={() => onShow(info)}>
            <Table /> 스프레드시트에 보이기
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
