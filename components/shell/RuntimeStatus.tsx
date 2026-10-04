"use client";

// 런타임 상태 표시(헤더/상태 바 인라인): 로드 진행률 → 준비됨/실행 중/오류 + 재설정 버튼

import { useEffect, useState } from "react";
import { toast } from "sonner";

import { Button } from "@/components/ui/button";
import { Progress } from "@/components/ui/progress";
import { useWorkbookStore } from "@/lib/grid/model";
import type { RuntimeClient, RuntimeStatusName } from "@/lib/runtime/client";

const DOT: Record<Exclude<RuntimeStatusName, "idle" | "loading">, { cls: string; label: string }> = {
  ready: { cls: "text-green-600", label: "준비됨" },
  running: { cls: "text-warning", label: "실행 중" },
  error: { cls: "text-destructive", label: "오류" },
  rebooting: { cls: "text-warning", label: "재설정 중" },
};

export function RuntimeStatus({ client }: { client: RuntimeClient }) {
  const [status, setStatus] = useState<RuntimeStatusName>(client.getStatus());
  const [prog, setProg] = useState({ pct: 0, label: "" });

  useEffect(() => {
    const offs = [
      client.on("status", setStatus),
      client.on("progress", setProg),
      client.on("reboot", () => {
        useWorkbookStore.getState().resetExecCounts(); // 변수와 함께 실행 순번도 처음부터
        toast("런타임이 재설정되어 변수가 초기화되었습니다");
      }),
    ];
    setStatus(client.getStatus());
    return () => offs.forEach((off) => off());
  }, [client]);

  if (status === "idle" || status === "loading") {
    return (
      <div className="flex items-center gap-2 text-xs text-muted-foreground">
        <Progress value={prog.pct} className="w-24" aria-label="런타임 로드 진행률" />
        <span>
          런타임 로드 중 {prog.pct}%{prog.label ? ` · ${prog.label}` : ""}
        </span>
      </div>
    );
  }

  const dot = DOT[status];
  return (
    <div className="flex items-center gap-2 text-xs">
      <span className={dot.cls}>●</span>
      <span>{dot.label}</span>
      {(status === "ready" || status === "error") && (
        <Button
          variant="ghost"
          size="xs"
          className="text-muted-foreground"
          onClick={() => void client.reset()}
        >
          런타임 재설정
        </Button>
      )}
    </div>
  );
}
