"use client";

// 부록 O.5: 모델 결과 → 시트 — R처럼 ① 모델 ② 항목 ③ 데이터셋·조건(예측·구간·확률·평가)을 고르면
// 항목마다 출력 바인딩을 만들어 셀에 값으로 놓는다. 블록을 다시 실행할 때마다 값이 갱신된다.
// 항목 목록은 모델 종류마다 다르다 — 각 항목의 현재 값과 라이브러리 코드를 함께 보여 주고,
// 코드는 복사하거나 새 블록으로 보내 사용자가 직접 코드로 쓸 수 있게 한다.

import { Fragment, useEffect, useMemo, useState } from "react";
import { ArrowSquareOut, Copy } from "@phosphor-icons/react";
import { toast } from "sonner";
import { Button } from "@/components/ui/button";
import { Checkbox } from "@/components/ui/checkbox";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from "@/components/ui/dialog";
import { Label } from "@/components/ui/label";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { copyText } from "@/components/reference/code-popup";
import { formatA1 } from "@/lib/grid/a1";
import { notifyWorkbookEdit } from "@/lib/grid/calc-host";
import { insertSnippetAsBlock } from "@/lib/grid/insert-snippet";
import { useWorkbookStore } from "@/lib/grid/model";
import { buildModelOutputSpecs, withImports, type Interval } from "@/lib/grid/model-output";
import { getRuntimeClient } from "@/lib/runtime/client";
import type { ModelMember, VariableInfo } from "@/lib/runtime/protocol";
import type { PyBlock } from "@/types/workbook";

const NONE = "__none__";

const copy = async (code: string) => {
  if (await copyText(code)) toast.success("코드를 복사했습니다");
};

/** 코드 한 줄을 새 블록(현재 블록 아래)으로 — 마지막 표현식이라 실행하면 결과가 셀에 놓인다 */
function codeToBlock(block: PyBlock, label: string, code: string) {
  const res = insertSnippetAsBlock(block.id, "below", label, `# ${label}\n${withImports(code)}`);
  if (!res) return;
  notifyWorkbookEdit([], [res.id]);
  toast.success("새 블록을 추가했습니다 — 코드를 고쳐 쓰거나 그대로 실행하면 결과가 셀에 놓입니다");
}

function CodeActions({ block, label, code }: { block: PyBlock; label: string; code: string }) {
  return (
    <span className="flex shrink-0 items-center">
      <button
        type="button"
        className="rounded p-1 hover:bg-accent"
        title="코드 복사"
        aria-label={`${label} 코드 복사`}
        onClick={() => void copy(code)}
      >
        <Copy className="size-3.5" />
      </button>
      <button
        type="button"
        className="rounded p-1 hover:bg-accent"
        title="이 코드로 새 블록 만들기"
        aria-label={`${label} 코드를 새 블록으로`}
        onClick={() => codeToBlock(block, label, code)}
      >
        <ArrowSquareOut className="size-3.5" />
      </button>
    </span>
  );
}

export default function ModelResultDialog({
  block,
  open,
  onOpenChange,
}: {
  block: PyBlock;
  open: boolean;
  onOpenChange: (open: boolean) => void;
}) {
  const [vars, setVars] = useState<VariableInfo[] | null>(null);
  const [model, setModel] = useState<string>("");
  const [members, setMembers] = useState<string[]>([]);
  const [data, setData] = useState<string>(NONE);
  const [predict, setPredict] = useState(true);
  const [proba, setProba] = useState(false);
  const [interval, setIntervalKind] = useState<Interval>("none");
  const [keep, setKeep] = useState(false);
  const [target, setTarget] = useState<string>(NONE);
  const [where, setWhere] = useState<"auto" | "selection">("auto");
  const selection = useWorkbookStore((s) => s.selection);
  const activeSheetId = useWorkbookStore((s) => s.activeSheetId);

  useEffect(() => {
    if (!open) return;
    setVars(null);
    const client = getRuntimeClient();
    if (client.getStatus() !== "ready") {
      setVars([]);
      return;
    }
    client
      .inspect()
      .then((vs) => {
        setVars(vs);
        const first = vs.find((v) => v.model);
        setModel((m) => (vs.some((v) => v.name === m && v.model) ? m : (first?.name ?? "")));
      })
      .catch(() => setVars([]));
  }, [open]);

  const models = useMemo(() => (vars ?? []).filter((v) => v.model), [vars]);
  const tables = useMemo(() => (vars ?? []).filter((v) => v.columns), [vars]);
  const info = models.find((v) => v.name === model)?.model;
  const dataVar = tables.find((v) => v.name === data);

  // 모델이 바뀌면 기본 선택: 요약표 묶음
  useEffect(() => {
    if (!info) return;
    setMembers(info.members.filter((m) => m.group === "요약표").map((m) => m.expr).slice(0, 2));
    if (!info.intervals) setIntervalKind("none");
    setProba(false);
  }, [info]);

  const groups = useMemo(() => {
    const out = new Map<string, ModelMember[]>();
    for (const m of info?.members ?? []) out.set(m.group, [...(out.get(m.group) ?? []), m]);
    return [...out];
  }, [info]);

  const choice = {
    model,
    members,
    data: data === NONE ? undefined : data,
    predict: data !== NONE && predict ? { interval, level: 0.95, keep } : undefined,
    proba: data !== NONE && proba,
    target: data !== NONE && target !== NONE ? target : undefined,
  };
  const specs = info ? buildModelOutputSpecs(choice, info, dataVar) : [];

  const place = () => {
    const st = useWorkbookStore.getState();
    const start =
      where === "selection" && selection
        ? { sheetId: activeSheetId, r: selection.r0, c: selection.c0 }
        : undefined;
    const ids = st.addOutputs(block.id, specs, start);
    if (ids.length === 0) return;
    notifyWorkbookEdit([], [block.id]); // 자동 모드: 바로 재실행해 값이 놓인다
    toast.success(`출력 ${ids.length}개를 추가했습니다 — 블록을 실행하면 값이 셀에 놓입니다`);
    onOpenChange(false);
  };

  const selAddr = selection
    ? formatA1({ r0: selection.r0, c0: selection.c0, r1: selection.r0, c1: selection.c0 })
    : null;
  const toggle = (expr: string, on: boolean) =>
    setMembers((cur) => (on ? [...cur, expr] : cur.filter((x) => x !== expr)));

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="max-h-[88vh] overflow-y-auto sm:max-w-3xl">
        <DialogHeader>
          <DialogTitle>모델 결과 → 시트</DialogTitle>
          <DialogDescription>
            모델 종류에 따라 쓸 수 있는 속성이 다릅니다. 체크한 항목은 셀에 값으로 놓이고(블록을 다시
            실행하면 갱신), 각 항목의 코드는 복사하거나 새 블록으로 보내 직접 쓸 수 있습니다.
          </DialogDescription>
        </DialogHeader>

        {vars === null ? (
          <p className="text-sm text-muted-foreground">변수 목록을 읽는 중…</p>
        ) : models.length === 0 ? (
          <p className="text-sm text-muted-foreground">
            적합된 모델 변수가 없습니다. 먼저 블록을 실행해 모델을 만드세요 — 예:{" "}
            <code className="font-mono">model = smf.ols(&quot;y ~ x1 + x2&quot;, df).fit()</code> 또는{" "}
            <code className="font-mono">model = LinearRegression().fit(X, y)</code>
          </p>
        ) : (
          <div className="space-y-4 text-sm">
            <section className="space-y-1.5">
              <Label>① 모델</Label>
              <Select value={model} onValueChange={setModel}>
                <SelectTrigger className="h-8 w-full" aria-label="모델 변수">
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  {models.map((v) => (
                    <SelectItem key={v.name} value={v.name}>
                      <span className="font-mono">{v.name}</span>
                      <span className="ml-2 text-xs text-muted-foreground">
                        {v.type} · {v.model!.kind}
                      </span>
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </section>

            {info && (
              <section className="space-y-1.5">
                <Label>
                  ② 속성 ({info.members.length}개 — {models.find((v) => v.name === model)?.type})
                </Label>
                <div className="max-h-80 overflow-y-auto rounded border" data-testid="model-members">
                  <table className="w-full text-xs">
                    <thead className="sticky top-0 bg-muted text-left text-muted-foreground">
                      <tr>
                        <th className="w-8 px-2 py-1">셀</th>
                        <th className="px-2 py-1">항목</th>
                        <th className="px-2 py-1">현재 값</th>
                        <th className="px-2 py-1">코드</th>
                      </tr>
                    </thead>
                    <tbody>
                      {groups.map(([group, items]) => (
                        <Fragment key={group}>
                          <tr>
                            <td colSpan={4} className="bg-muted/40 px-2 py-0.5 font-semibold text-muted-foreground">
                              {group}
                            </td>
                          </tr>
                          {items.map((m) => (
                            <tr key={m.expr} className="border-t align-top">
                              <td className="px-2 py-1">
                                <Checkbox
                                  aria-label={`${m.label} 셀에 놓기`}
                                  checked={members.includes(m.expr)}
                                  onCheckedChange={(v) => toggle(m.expr, v === true)}
                                />
                              </td>
                              <td className="px-2 py-1">{m.label}</td>
                              <td className="max-w-44 truncate px-2 py-1 font-mono text-muted-foreground" title={m.preview}>
                                {m.preview}
                              </td>
                              <td className="px-2 py-1">
                                <span className="flex items-center gap-1">
                                  <code className="min-w-0 flex-1 truncate font-mono" title={m.code}>
                                    {m.code}
                                  </code>
                                  <CodeActions block={block} label={`${model}: ${m.label}`} code={m.code} />
                                </span>
                              </td>
                            </tr>
                          ))}
                        </Fragment>
                      ))}
                    </tbody>
                  </table>
                </div>
              </section>
            )}

            {info?.predict && (
              <section className="space-y-1.5">
                <Label>③ 데이터셋 (예측·평가용, 선택)</Label>
                <Select value={data} onValueChange={setData}>
                  <SelectTrigger className="h-8 w-full" aria-label="데이터셋 변수">
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value={NONE}>사용 안 함</SelectItem>
                    {tables.map((v) => (
                      <SelectItem key={v.name} value={v.name}>
                        <span className="font-mono">{v.name}</span>
                        <span className="ml-2 text-xs text-muted-foreground">
                          {v.shape ? `${v.shape[0]}×${v.shape[1]}` : v.type}
                        </span>
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
                {dataVar && (
                  <div className="space-y-2 rounded border p-2">
                    <label className="flex items-center gap-2">
                      <Checkbox checked={predict} onCheckedChange={(v) => setPredict(v === true)} />
                      예측값 (predict)
                    </label>
                    {predict && (
                      <div className="flex flex-wrap items-center gap-3 pl-6">
                        <Select value={interval} onValueChange={(v) => setIntervalKind(v as Interval)}>
                          <SelectTrigger className="h-7 w-40 text-xs" aria-label="구간">
                            <SelectValue />
                          </SelectTrigger>
                          <SelectContent>
                            <SelectItem value="none">구간 없음</SelectItem>
                            <SelectItem value="confidence" disabled={!info.intervals}>
                              신뢰구간 95%
                            </SelectItem>
                            <SelectItem value="prediction" disabled={!info.intervals}>
                              예측구간 95%
                            </SelectItem>
                          </SelectContent>
                        </Select>
                        <label className="flex items-center gap-2 text-xs">
                          <Checkbox checked={keep} onCheckedChange={(v) => setKeep(v === true)} />
                          원본 열과 함께
                        </label>
                      </div>
                    )}
                    {info.proba && (
                      <label className="flex items-center gap-2">
                        <Checkbox checked={proba} onCheckedChange={(v) => setProba(v === true)} />
                        클래스 확률 (predict_proba)
                      </label>
                    )}
                    <div className="flex flex-wrap items-center gap-2">
                      <span>평가 지표 (R²·RMSE·MAE) 목표 열</span>
                      <Select value={target} onValueChange={setTarget}>
                        <SelectTrigger className="h-7 w-40 text-xs" aria-label="목표 열">
                          <SelectValue />
                        </SelectTrigger>
                        <SelectContent>
                          <SelectItem value={NONE}>사용 안 함</SelectItem>
                          {dataVar.columns!.map((c) => (
                            <SelectItem key={c} value={c} className="font-mono">
                              {c}
                            </SelectItem>
                          ))}
                        </SelectContent>
                      </Select>
                    </div>
                  </div>
                )}
              </section>
            )}

            <section className="space-y-1.5">
              <Label>④ 놓을 위치</Label>
              <Select value={where} onValueChange={(v) => setWhere(v as "auto" | "selection")}>
                <SelectTrigger className="h-8 w-full" aria-label="놓을 위치">
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="auto">블록 출력 오른쪽 빈 칸부터</SelectItem>
                  <SelectItem value="selection" disabled={!selAddr}>
                    선택한 셀부터 {selAddr ? `(${selAddr})` : ""}
                  </SelectItem>
                </SelectContent>
              </Select>
              <p className="text-xs text-muted-foreground">
                항목마다 한 칸씩 띄워 가로로 놓습니다. 겹치면 #SPILL! — 출력 주소를 눌러 옮길 수 있습니다.
              </p>
            </section>

            {specs.length > 0 && (
              <section className="space-y-1">
                <div className="flex items-center justify-between">
                  <Label>셀에 놓을 항목 ({specs.length})</Label>
                  <Button
                    variant="ghost"
                    size="sm"
                    className="h-6 text-xs"
                    onClick={() => void copy(specs.map((s) => `# ${s.label}\n${s.code}`).join("\n\n"))}
                  >
                    <Copy className="size-3" /> 코드 전체 복사
                  </Button>
                </div>
                <ul className="space-y-0.5 rounded bg-muted/40 p-2 text-xs">
                  {specs.map((s) => (
                    <li key={s.selection.variable} className="flex items-center gap-2">
                      <span className="shrink-0">{s.label}</span>
                      <code className="min-w-0 flex-1 truncate font-mono text-muted-foreground" title={s.code}>
                        {s.code}
                      </code>
                      <CodeActions block={block} label={s.label} code={s.code} />
                    </li>
                  ))}
                </ul>
              </section>
            )}
          </div>
        )}

        <DialogFooter>
          <Button variant="ghost" onClick={() => onOpenChange(false)}>
            취소
          </Button>
          <Button onClick={place} disabled={specs.length === 0}>
            시트에 놓기 ({specs.length})
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
