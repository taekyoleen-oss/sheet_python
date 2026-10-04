"use client";

// 부록 P: 변수 세부 내용 팝업 — 속성 창에서 변수를 누르면 열린다.
//  표(DataFrame·Series·배열): 열·형 + 상위 100행, 요약 통계(describe) 탭
//  모델: 시트로 보낼 수 있는 속성(현재 값·코드) 목록
//  그 밖의 값: repr 원문
// 런타임 객체 모드 실행(client.run)으로 미리보기를 받는다 — 셀에는 아무것도 쓰지 않는다.

import { useEffect, useState } from "react";
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

function PreviewView({ p }: { p: PreviewPayload | { kind: "error"; message: string } | null }) {
  if (!p) return <p className="py-6 text-center text-sm text-muted-foreground">불러오는 중…</p>;
  if (p.kind === "error") return <p className="text-sm text-destructive">{p.message}</p>;
  if (p.kind === "image") return <p className="text-sm text-muted-foreground">이미지 값입니다 — 시트에 객체로 놓아 보세요.</p>;
  if (p.kind === "repr")
    return <pre className="max-h-[55vh] overflow-auto rounded bg-muted/40 p-3 font-mono text-xs whitespace-pre">{p.repr}</pre>;
  return (
    <div className="max-h-[55vh] overflow-auto rounded border">
      <table className="w-full border-collapse font-mono text-xs tabular-nums">
        <thead className="sticky top-0 bg-muted">
          <tr>
            <th className="border-b px-2 py-1 text-right font-normal text-muted-foreground">#</th>
            {p.columns.map((c, i) => (
              <th key={i} className="border-b px-2 py-1 text-left">
                <div>{c}</div>
                <div className="font-normal text-muted-foreground">{p.dtypes[i]}</div>
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {p.rows.map((row, r) => (
            <tr key={r} className="border-b last:border-0">
              <td className="px-2 py-0.5 text-right text-muted-foreground">{r}</td>
              {row.map((v, c) => (
                <td key={c} className={cn("px-2 py-0.5", typeof v === "number" ? "text-right" : "text-left")}>
                  {v === null ? <span className="text-muted-foreground">NaN</span> : String(v)}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
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
            <table className="w-full text-xs">
              <thead className="sticky top-0 bg-muted text-left text-muted-foreground">
                <tr>
                  <th className="px-2 py-1">묶음</th>
                  <th className="px-2 py-1">항목</th>
                  <th className="px-2 py-1">현재 값</th>
                  <th className="px-2 py-1">코드</th>
                </tr>
              </thead>
              <tbody>
                {info.model.members.map((m) => (
                  <tr key={m.expr} className="border-t align-top">
                    <td className="px-2 py-1 text-muted-foreground">{m.group}</td>
                    <td className="px-2 py-1">{m.label}</td>
                    <td className="max-w-56 truncate px-2 py-1 font-mono" title={m.preview}>
                      {m.preview}
                    </td>
                    <td className="px-2 py-1 font-mono text-muted-foreground">{m.code}</td>
                  </tr>
                ))}
              </tbody>
            </table>
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
            <PreviewView p={tab === "data" ? data : stats} />
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
