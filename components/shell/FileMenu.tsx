"use client";

// 파일 메뉴 — 새 워크북·열기·저장·내보내기·샘플·최근 워크북 (§1.5, §4.4 FileMenu)

import { Fragment, useRef, useState } from "react";
import { CaretDown } from "@phosphor-icons/react";
import { toast } from "sonner";
import chainLadderSample from "@/data/sample-workbooks/chain-ladder.pygrid.json";
import claimSeveritySample from "@/data/sample-workbooks/claim-severity.pygrid.json";
import lifeTableSample from "@/data/sample-workbooks/life-table.pygrid.json";
import lossRatioSample from "@/data/sample-workbooks/loss-ratio.pygrid.json";
import ApiKeyDialog, { openApiKeyDialog } from "@/components/shell/ApiKeyDialog";
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
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuSub,
  DropdownMenuSubContent,
  DropdownMenuSubTrigger,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { Input } from "@/components/ui/input";
import { useWorkbookStore } from "@/lib/grid/model";
import { importData, type ImportOptions } from "@/lib/io/data-import";
import {
  downloadBlob,
  downloadWorkbookJson,
  parseWorkbookJson,
} from "@/lib/io/workbook-json";
// SheetJS(~140kB gz)는 첫 페인트를 막지 않도록 사용 시점에 동적 로드한다
import { listWorkbooks, getWorkbook, loadSettings, saveSettings } from "@/lib/storage/db";
import { createWorkbook } from "@/lib/grid/model";
import type { Workbook } from "@/types/workbook";

/** 샘플/파일 워크북 로드 — 파일의 id 유지(최근 목록 중복 시 덮어씀, 단순 규칙) */
export function loadWorkbookData(wb: Workbook): void {
  useWorkbookStore.getState().loadWorkbook(structuredClone(wb));
}

/** 첫 방문 기본 워크북 (WorkbookShell) */
export const SAMPLE_LIFE_TABLE = lifeTableSample as unknown as Workbook;

interface SampleWorkbook {
  label: string;
  /** 데이터 내장 워크북은 수백 KB라 첫 페인트를 막지 않도록 클릭 시 동적 로드한다 */
  load: () => Promise<Workbook>;
}

const dyn = (m: Promise<{ default: unknown }>): Promise<Workbook> =>
  m.then((r) => r.default as Workbook);

/**
 * 예제 워크북 상위 카테고리 (부록 N).
 *  · 통계분석 — 표본에서 모형을 고르는 절차(회귀·GLM·빈도심도·생존분석)
 *  · 위험률 산출 — 확정된 위험률에서 기수·급부 현가·보험료로 내려가는 계리 절차
 * 두 갈래는 절차가 달라 한 목록에 섞지 않는다. 코드 삽입 팝업의 카테고리와 같은 구분이다.
 */
const SAMPLE_STAT: SampleWorkbook[] = [
  {
    // 부록 N — 통계분석 → 전처리 → 특성공학 → Ridge/Lasso/ElasticNet·다항·로그 → 평가
    label: "임금 회귀 예측 (5단계)",
    load: () => dyn(import("@/data/sample-workbooks/wage-regression.pygrid.json")),
  },
  {
    label: "보험료 요인 분석 (GLM)",
    load: () => dyn(import("@/data/sample-workbooks/premium-glm.pygrid.json")),
  },
  {
    label: "빈도·심도 모형",
    load: () => dyn(import("@/data/sample-workbooks/freq-severity.pygrid.json")),
  },
  {
    label: "생존분석·유지율",
    load: () => dyn(import("@/data/sample-workbooks/survival-retention.pygrid.json")),
  },
];

/** 위험률 산출 — 하위 카테고리(위험률·생명표 / 보험료 / 다급부 / 준비금) 순서로 묶는다 */
const SAMPLE_RISK: { group: string; items: SampleWorkbook[] }[] = [
  {
    group: "위험률·생명표",
    items: [
      { label: "위험률·생명표", load: async () => SAMPLE_LIFE_TABLE },
      {
        // 부록 M.4 — 발생자수÷추계인구 → 조율 → 비만동반·안전할증 → 보간·평활 → 원본 대조
        label: "위험률 산출 (원시통계)",
        load: () => dyn(import("@/data/sample-workbooks/risk-rate.pygrid.json")),
      },
    ],
  },
  {
    group: "보험료 산출",
    items: [
      {
        // 부록 M — 계산기수 → 보험료 → 준비금·환급금 → 표준/적용 비교
        label: "정기보험 (계산기수)",
        load: () => dyn(import("@/data/sample-workbooks/premium-term.pygrid.json")),
      },
      {
        // 부록 M.5 — 체증형(Rx 기수) · 미달체(사망률 ×3) → 3종 비교 → 원본 P·미달P 대조
        label: "정기보험 변형 (체증형·미달체)",
        load: () => dyn(import("@/data/sample-workbooks/term-variants.pygrid.json")),
      },
      {
        // 부록 M.5 — 직종축 위험률 → 경과 기수표 → 순수보장형·만기환급형(50%) → PV테이블 대조
        label: "상해공제 (직종축)",
        load: () => dyn(import("@/data/sample-workbooks/accident-class.pygrid.json")),
      },
    ],
  },
  {
    group: "다중탈퇴·다급부",
    items: [
      {
        // 부록 M.4 — 다중탈퇴 생존자표 → Cx·Mx(90일 면책) → 급부배율 → 공제료
        label: "암보험 (다중탈퇴)",
        load: () => dyn(import("@/data/sample-workbooks/cancer-multi.pygrid.json")),
      },
      {
        // 부록 M.6 — 주요암 qc → 암발생후사망률 2차원표 → 진단 후 생활비 연금현가 → 급부 A~F
        label: "암보험 (진단 후 생활비)",
        load: () => dyn(import("@/data/sample-workbooks/cancer-annuity.pygrid.json")),
      },
      {
        // 부록 M.5 — 이중탈퇴 lx·lx′ → 급부 5종 기수 차분 → 급부배율 SUMX → 순·영업공제료
        label: "종신공제 (다급부)",
        load: () => dyn(import("@/data/sample-workbooks/whole-life-multi.pygrid.json")),
      },
      {
        // 부록 M.6 — CI 11종 경합 제거 결합 → 3중탈퇴 생존자표 → CI 선지급 급부 현가 → 보험료
        label: "CI종신 (3중탈퇴)",
        load: () => dyn(import("@/data/sample-workbooks/ci-whole-life.pygrid.json")),
      },
    ],
  },
  {
    group: "준비금·환급금",
    items: [
      {
        // 부록 M.4 — CDR 발생률 → 이중탈퇴 생존자 → 현금흐름 PV → 환급률 → P테이블 대조
        label: "무해지환급형 (해지율 PV)",
        load: () => dyn(import("@/data/sample-workbooks/nonsurrender.pygrid.json")),
      },
      {
        label: "지급준비금 (체인래더)",
        load: async () => chainLadderSample as unknown as Workbook,
      },
    ],
  },
];

/** 기능 소개용 소형 예제 (부록 H.2) */
const SAMPLE_BASIC: SampleWorkbook[] = [
  { label: "손해율 집계", load: async () => lossRatioSample as unknown as Workbook },
  { label: "청구 심도 적합", load: async () => claimSeveritySample as unknown as Workbook },
];

/** 샘플 워크북 열기 — 동적 로드 실패(오프라인 등)는 토스트로 알린다 */
async function openSampleWorkbook(s: SampleWorkbook): Promise<void> {
  try {
    loadWorkbookData(await s.load());
  } catch (e) {
    toast.error(`샘플을 불러오지 못했습니다: ${(e as Error).message}`);
  }
}

/** 확장자별 열기 — FileMenu 선택·드래그 앤 드롭 공용 */
export async function openWorkbookFile(file: File): Promise<void> {
  const name = file.name.toLowerCase();
  try {
    if (name.endsWith(".pygrid.json") || name.endsWith(".json")) {
      loadWorkbookData(parseWorkbookJson(await file.text()));
    } else if (name.endsWith(".csv")) {
      const { csvToSheet } = await import("@/lib/io/csv");
      const sheet = csvToSheet(await file.text(), file.name.replace(/\.csv$/i, ""));
      const wb = createWorkbook();
      wb.title = file.name.replace(/\.csv$/i, "");
      wb.sheets = [sheet];
      useWorkbookStore.getState().loadWorkbook(wb);
    } else if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
      const { sheetsFromFileData } = await import("@/lib/io/xlsx");
      const sheets = sheetsFromFileData(await file.arrayBuffer());
      if (sheets.length === 0) throw new Error("시트가 없습니다");
      const wb = createWorkbook();
      wb.title = file.name.replace(/\.(xlsx|xls)$/i, "");
      wb.sheets = sheets;
      useWorkbookStore.getState().loadWorkbook(wb);
    } else {
      toast.error("지원하지 않는 파일 형식입니다 (.pygrid.json / .csv / .xlsx)");
      return;
    }
    toast(`"${file.name}"을(를) 열었습니다`);
  } catch (e) {
    toast.error(`파일 열기 실패: ${(e as Error).message}`);
  }
}

/** 샘플 데이터셋 (public/samples, 부록 E.1) — 소스 SAMPLE_DATASETS 이식 */
const SAMPLE_DATASETS = [
  { file: "policy.xlsx", label: "policy.xlsx — 계약·고객 (600행)" },
  { file: "claims.xlsx", label: "claims.xlsx — 청구·손해액 (600행)" },
  { file: "experience.xlsx", label: "experience.xlsx — 경험데이터 (800행)" },
  { file: "triangle.xlsx", label: "triangle.xlsx — 런오프 누적 삼각형" },
  { file: "mortality_table.xlsx", label: "mortality_table.xlsx — 생명표 qx (0~100세)" },
  { file: "wage.xlsx", label: "wage.xlsx — 임금·근로자 특성 (534행)" },
];

const DEFAULT_IMPORT_OPTS: ImportOptions = { toSheet: true, toFs: true, makeBlock: "xl" };


/** 최근 워크북 열기 — 파일 메뉴와 헤더의 최근 메뉴가 공유 */
async function openRecent(id: string): Promise<void> {
  const wb = await getWorkbook(id);
  if (wb) useWorkbookStore.getState().loadWorkbook(wb);
  else toast.error("워크북을 찾을 수 없습니다 (저장소에서 정리됨)");
}

/** 샘플 워크북 메뉴 — 파일 메뉴에서 분리해 헤더에 직접 노출 (사용 빈도 높음) */
export function SampleWorkbookMenu() {
  return (
    <DropdownMenu>
      <DropdownMenuTrigger asChild>
        <Button variant="ghost" size="sm" className="h-7 px-2 text-xs">
          샘플 워크북
        </Button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="start" className="max-h-[70vh] w-64 overflow-y-auto">
        <DropdownMenuLabel className="text-xs text-primary">
          통계분석 — 데이터 + 단계별 코드
        </DropdownMenuLabel>
        {SAMPLE_STAT.map((s) => (
          <DropdownMenuItem key={s.label} onClick={() => void openSampleWorkbook(s)}>
            {s.label}
          </DropdownMenuItem>
        ))}
        <DropdownMenuSeparator />
        <DropdownMenuLabel className="text-xs text-primary">
          위험률 산출 — 데이터 + 단계별 코드
        </DropdownMenuLabel>
        {SAMPLE_RISK.map(({ group, items }) => (
          <Fragment key={group}>
            <DropdownMenuLabel className="pl-3 text-[11px] font-normal text-muted-foreground">
              {group}
            </DropdownMenuLabel>
            {items.map((s) => (
              <DropdownMenuItem
                key={s.label}
                className="pl-4"
                onClick={() => void openSampleWorkbook(s)}
              >
                {s.label}
              </DropdownMenuItem>
            ))}
          </Fragment>
        ))}
        <DropdownMenuSeparator />
        <DropdownMenuLabel className="text-xs text-muted-foreground">기본 예제</DropdownMenuLabel>
        {SAMPLE_BASIC.map((s) => (
          <DropdownMenuItem key={s.label} onClick={() => void openSampleWorkbook(s)}>
            {s.label}
          </DropdownMenuItem>
        ))}
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

/** 최근 워크북 메뉴 — 열 때마다 idb에서 목록을 다시 읽는다 */
export function RecentWorkbookMenu() {
  const [recent, setRecent] = useState<Workbook[]>([]);
  const currentId = useWorkbookStore((s) => s.workbook.id);
  return (
    <DropdownMenu
      onOpenChange={(open) => {
        if (open) void listWorkbooks().then(setRecent);
      }}
    >
      <DropdownMenuTrigger asChild>
        <Button variant="ghost" size="sm" className="h-7 px-2 text-xs">
          최근 워크북
        </Button>
      </DropdownMenuTrigger>
      <DropdownMenuContent align="start" className="max-h-72 w-64 overflow-y-auto">
        {recent.length === 0 ? (
          <DropdownMenuItem disabled>저장된 워크북 없음</DropdownMenuItem>
        ) : (
          recent.map((wb) => (
            <DropdownMenuItem key={wb.id} onClick={() => void openRecent(wb.id)}>
              <div className="min-w-0">
                <div className="truncate text-xs">
                  {wb.id === currentId ? "● " : ""}
                  {wb.title}
                </div>
                <div className="font-mono text-[10px] text-muted-foreground">
                  {wb.updatedAt.slice(0, 16).replace("T", " ")}
                </div>
              </div>
            </DropdownMenuItem>
          ))
        )}
      </DropdownMenuContent>
    </DropdownMenu>
  );
}

export default function FileMenu() {
  const inputRef = useRef<HTMLInputElement>(null);
  const dataInputRef = useRef<HTMLInputElement>(null);
  const [recent, setRecent] = useState<Workbook[]>([]);
  // 데이터 불러오기 (R5) — pending이 있으면 옵션 다이얼로그가 열린다
  const [pending, setPending] = useState<{ name: string; bytes: Uint8Array } | null>(null);
  const [impOpts, setImpOpts] = useState<ImportOptions>(DEFAULT_IMPORT_OPTS);
  const [importing, setImporting] = useState(false);
  const [urlOpen, setUrlOpen] = useState(false);
  const [url, setUrl] = useState("");

  /** 소스 확보 후 옵션 다이얼로그 열기 — 지난 선택(app settings)을 기본값으로 */
  const askImport = async (name: string, bytes: Uint8Array) => {
    const s = await loadSettings();
    setImpOpts(s?.dataImport ?? DEFAULT_IMPORT_OPTS);
    setPending({ name, bytes });
  };

  const loadSample = async (file: string) => {
    try {
      const res = await fetch(`/samples/${file}`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      await askImport(file, new Uint8Array(await res.arrayBuffer()));
    } catch (e) {
      toast.error(`샘플을 불러오지 못했습니다: ${(e as Error).message}`);
    }
  };

  const loadUrl = async () => {
    const u = url.trim();
    if (!u) return;
    try {
      const res = await fetch(u);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const bytes = new Uint8Array(await res.arrayBuffer());
      let name = "data.csv";
      try {
        name = decodeURIComponent(new URL(u, location.href).pathname.split("/").pop() || "") || name;
      } catch {
        // URL 파싱 실패 — 기본 이름 사용
      }
      setUrlOpen(false);
      await askImport(name, bytes);
    } catch (e) {
      toast.error(
        `URL에서 불러오지 못했습니다: ${(e as Error).message} — 대상 서버가 CORS(교차 출처 요청)를 허용하지 않으면 브라우저에서 직접 받을 수 없습니다. 파일을 내려받은 뒤 "내 파일 추가"로 여세요.`,
      );
    }
  };

  const confirmImport = async () => {
    if (!pending) return;
    setImporting(true);
    try {
      void saveSettings({ dataImport: impOpts });
      const res = await importData(pending.name, pending.bytes, impOpts);
      const parts = [
        res.sheetNames.length > 0 ? `시트 ${res.sheetNames.join(", ")}` : null,
        impOpts.toFs ? "워커 FS" : null,
        res.blockId ? "로드 블록" : null,
      ].filter(Boolean);
      toast(`"${pending.name}" 불러옴 — ${parts.join(" · ") || "반영 대상 없음"}`);
      res.fs?.catch((e) => toast.error(`워커 FS 기록 실패: ${(e as Error).message}`));
      setPending(null);
    } catch (e) {
      toast.error(`데이터 불러오기 실패: ${(e as Error).message}`);
    } finally {
      setImporting(false);
    }
  };

  const save = () => {
    const verdict = downloadWorkbookJson(useWorkbookStore.getState().workbook);
    if (verdict === "warn") {
      toast.warning("워크북이 50MB를 넘습니다. 데이터 시트를 나누는 것을 권장합니다.");
    }
  };

  const exportXlsx = async () => {
    const { sheetsToXlsxBlob } = await import("@/lib/io/xlsx");
    const wb = useWorkbookStore.getState().workbook;
    downloadBlob(sheetsToXlsxBlob(wb.sheets), `${wb.title || "워크북"}.xlsx`);
  };

  const exportCsv = async () => {
    const { sheetToCsv } = await import("@/lib/io/csv");
    const st = useWorkbookStore.getState();
    const sheet = st.workbook.sheets.find((s) => s.id === st.activeSheetId);
    if (!sheet) return;
    downloadBlob(
      new Blob(["﻿" + sheetToCsv(sheet)], { type: "text/csv;charset=utf-8" }), // BOM: Excel 한글 인식
      `${sheet.name}.csv`,
    );
  };

  const refreshRecent = () => {
    void listWorkbooks().then((list) => setRecent(list.slice(0, 10)));
  };

  const currentId = useWorkbookStore((s) => s.workbook.id);

  return (
    <>
      <input
        ref={inputRef}
        type="file"
        accept=".json,.csv,.xlsx,.xls"
        className="hidden"
        aria-label="워크북 파일 열기"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) void openWorkbookFile(file);
          e.target.value = ""; // 같은 파일 재선택 허용
        }}
      />
      <input
        ref={dataInputRef}
        type="file"
        accept=".csv,.txt,.xlsx,.xls"
        className="hidden"
        aria-label="데이터 파일 추가"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) void f.arrayBuffer().then((b) => askImport(f.name, new Uint8Array(b)));
          e.target.value = "";
        }}
      />
      <DropdownMenu onOpenChange={(open) => open && refreshRecent()}>
        <DropdownMenuTrigger asChild>
          <Button variant="ghost" size="sm" className="gap-1">
            파일 <CaretDown className="size-3" />
          </Button>
        </DropdownMenuTrigger>
        <DropdownMenuContent align="start" className="w-56">
          <DropdownMenuItem onClick={() => useWorkbookStore.getState().newWorkbook()}>
            새 워크북
          </DropdownMenuItem>
          <DropdownMenuItem onClick={() => inputRef.current?.click()}>
            열기… (.pygrid.json / .csv / .xlsx)
          </DropdownMenuItem>
          <DropdownMenuItem onClick={save}>저장 (.pygrid.json)</DropdownMenuItem>
          <DropdownMenuSeparator />
          {/* 데이터 불러오기 (R5) — '열기'와 달리 현재 워크북에 추가한다 */}
          <DropdownMenuSub>
            <DropdownMenuSubTrigger>데이터 불러오기</DropdownMenuSubTrigger>
            <DropdownMenuSubContent className="w-64">
              <DropdownMenuItem onClick={() => dataInputRef.current?.click()}>
                내 파일 추가… (.csv/.xlsx)
              </DropdownMenuItem>
              <DropdownMenuItem onClick={() => setUrlOpen(true)}>
                URL로 불러오기…
              </DropdownMenuItem>
              <DropdownMenuSeparator />
              <DropdownMenuSub>
                <DropdownMenuSubTrigger>샘플 데이터셋</DropdownMenuSubTrigger>
                <DropdownMenuSubContent className="w-80">
                  {SAMPLE_DATASETS.map((d) => (
                    <DropdownMenuItem key={d.file} onClick={() => void loadSample(d.file)}>
                      {d.label}
                    </DropdownMenuItem>
                  ))}
                </DropdownMenuSubContent>
              </DropdownMenuSub>
            </DropdownMenuSubContent>
          </DropdownMenuSub>
          <DropdownMenuItem onClick={() => openApiKeyDialog()}>AI 설정…</DropdownMenuItem>
          <DropdownMenuSeparator />
          <DropdownMenuItem onClick={() => void exportXlsx()}>
            XLSX로 내보내기 (전 시트)
          </DropdownMenuItem>
          <DropdownMenuItem onClick={() => void exportCsv()}>
            CSV로 내보내기 (활성 시트)
          </DropdownMenuItem>
        </DropdownMenuContent>
      </DropdownMenu>

      <ApiKeyDialog />

      {/* 데이터 불러오기 옵션 — 선택은 app settings에 기억된다 */}
      <Dialog open={pending !== null} onOpenChange={(o) => !o && setPending(null)}>
        <DialogContent className="max-w-[440px]">
          <DialogHeader>
            <DialogTitle>데이터 불러오기</DialogTitle>
            <DialogDescription className="font-mono">{pending?.name}</DialogDescription>
          </DialogHeader>
          <div className="space-y-3 text-sm">
            <label className="flex items-center gap-2">
              <Checkbox
                checked={impOpts.toSheet}
                onCheckedChange={(v) => setImpOpts({ ...impOpts, toSheet: v === true })}
              />
              시트로 표시 (CSV/XLSX는 새 시트로 추가)
            </label>
            <label className="flex items-center gap-2">
              <Checkbox
                checked={impOpts.toFs}
                onCheckedChange={(v) => setImpOpts({ ...impOpts, toFs: v === true })}
              />
              워커 파일시스템에 기록 (pd.read_* 사용 가능)
            </label>
            <fieldset className="space-y-1.5">
              <legend className="mb-1 text-xs text-muted-foreground">로드 블록 생성</legend>
              {(
                [
                  ["xl", "xl() 참조 (권장)"],
                  ["pandas", "pandas 코드 (pd.read_*)"],
                  ["none", "만들지 않음"],
                ] as const
              ).map(([v, label]) => (
                <label key={v} className="flex items-center gap-2">
                  <input
                    type="radio"
                    name="makeBlock"
                    value={v}
                    checked={impOpts.makeBlock === v}
                    onChange={() => setImpOpts({ ...impOpts, makeBlock: v })}
                  />
                  {label}
                </label>
              ))}
            </fieldset>
          </div>
          <DialogFooter>
            <Button variant="ghost" onClick={() => setPending(null)} disabled={importing}>
              취소
            </Button>
            <Button onClick={() => void confirmImport()} disabled={importing}>
              {importing ? "불러오는 중…" : "불러오기"}
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* URL로 불러오기 */}
      <Dialog open={urlOpen} onOpenChange={setUrlOpen}>
        <DialogContent className="max-w-[480px]">
          <DialogHeader>
            <DialogTitle>URL로 불러오기</DialogTitle>
            <DialogDescription>
              CSV/XLSX 파일 주소를 입력하세요. 대상 서버가 CORS(교차 출처 요청)를 허용해야
              브라우저에서 받을 수 있습니다.
            </DialogDescription>
          </DialogHeader>
          <Input
            value={url}
            onChange={(e) => setUrl(e.target.value)}
            placeholder="https://example.com/data.csv"
            aria-label="데이터 URL"
            onKeyDown={(e) => {
              if (e.key === "Enter") void loadUrl();
            }}
          />
          <DialogFooter>
            <Button variant="ghost" onClick={() => setUrlOpen(false)}>
              취소
            </Button>
            <Button onClick={() => void loadUrl()}>가져오기</Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </>
  );
}
