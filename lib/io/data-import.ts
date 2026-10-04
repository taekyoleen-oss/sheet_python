// 데이터 불러오기 통합 (부록 E R5) — 모든 데이터 로드는 두 곳에 착지한다:
// 시트(눈에 보이는 표) + 워커 FS(참조 코드의 pd.read_*가 파일명 그대로 동작).
// 인코딩 감지·pandas 로드 셀은 소스(Actuarial_Platform sheetCsv.ts·PyRunner buildLoadCell) 이식.

import { parseCellKey, type Sheet } from "@/types/workbook";
import { formatA1 } from "@/lib/grid/a1";
import { createReferenceBlocks } from "@/lib/grid/import-blocks";
import { useWorkbookStore } from "@/lib/grid/model";

/** 텍스트 파일 인코딩 감지 — 한글 Windows/Excel CSV(CP949) 대응. null = 기본(utf-8) */
export function detectTextEncoding(bytes: Uint8Array): "utf-8-sig" | "cp949" | null {
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf)
    return "utf-8-sig";
  try {
    new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    return null;
  } catch {
    return "cp949";
  }
}

/** 바이트 → 문자열. CP949는 euc-kr로 디코드(U+FFFD 깨짐 방지). utf-8 디코더는 BOM 자동 제거 */
export function decodeSmart(bytes: Uint8Array): string {
  const enc = detectTextEncoding(bytes) === "cp949" ? "euc-kr" : "utf-8";
  return new TextDecoder(enc).decode(bytes);
}

/** 기존 이름과 겹치지 않는 시트 이름 — "base", "base (2)", "base (3)" … */
export function uniqueSheetName(base: string, taken: Set<string>): string {
  if (!taken.has(base)) return base;
  let i = 2;
  while (taken.has(`${base} (${i})`)) i++;
  return `${base} (${i})`;
}

/** 파일 확장자 → pandas 로드 호출 문자열 (소스 loadCallFor 이식) */
function loadCallFor(name: string, encoding: string | null): string {
  const l = name.toLowerCase();
  const enc = encoding ? `, encoding="${encoding}"` : "";
  if (l.endsWith(".xlsx") || l.endsWith(".xls")) return `pd.read_excel("${name}")`;
  if (l.endsWith(".json")) return `pd.read_json("${name}"${enc})`;
  if (l.endsWith(".txt")) return `pd.read_csv("${name}", sep=None, engine="python"${enc})`;
  return `pd.read_csv("${name}"${enc})`;
}

/** 워커 FS의 파일을 pandas로 읽는 '로드 · 속성 확인' 블록 코드 (소스 buildLoadCell 이식) */
export function buildPandasLoadCode(name: string, encoding: string | null): string {
  const isExcel = /\.(xlsx|xls)$/i.test(name);
  return [
    "import pandas as pd",
    // 워커는 import 문 기반으로만 패키지를 지연 로드한다(loadPackagesFromImports) —
    // pd.read_excel의 지연 openpyxl import는 못 보므로 명시 import로 로드를 유도한다.
    // 현재 Pyodide 배포에는 Excel 엔진이 없어 ImportError가 난다(런타임 소관, parity E7) —
    // xlsx는 sheet() 모드를 권장. 엔진이 추가되면 이 코드가 그대로 동작한다.
    ...(isExcel ? ["import openpyxl  # pd.read_excel 엔진 (없으면 sheet() 모드 사용)"] : []),
    ...(encoding === "cp949"
      ? ["# 한글 Windows(CP949) 인코딩으로 감지되어 encoding을 지정했습니다"]
      : []),
    `df = ${loadCallFor(name, encoding)}`,
    "",
    'print("행·열:", df.shape)',
    'print("열 이름:", df.columns.tolist())',
    "df.head()",
  ].join("\n");
}

/** 시트 사용 범위를 sheet()로 읽는 로드 블록 코드 */
export function buildXlLoadCode(ref: string): string {
  return [
    `df = sheet("${ref}", headers=True)`,
    "",
    'print("행·열:", df.shape)',
    'print("열 이름:", df.columns.tolist())',
    "df.head()",
  ].join("\n");
}

/** 시트 사용 범위(A1 기준). 빈 시트는 A1 단일 셀로 폴백 (AI 컨텍스트 수집도 공용) */
export function usedRange(sheet: Sheet): { r0: number; c0: number; r1: number; c1: number } {
  let maxR = 0;
  let maxC = 0;
  for (const key of Object.keys(sheet.cells)) {
    const { r, c } = parseCellKey(key);
    if (r > maxR) maxR = r;
    if (c > maxC) maxC = c;
  }
  return { r0: 0, c0: 0, r1: maxR, c1: maxC };
}

export interface ImportOptions {
  /** CSV/XLSX를 새 시트로 추가(현재 워크북 유지) */
  toSheet: boolean;
  /** 워커 FS에 기록 — pd.read_*("파일명")이 그대로 동작. 같은 이름은 덮어쓴다 */
  toFs: boolean;
  /** 로드 블록 자동 생성(실행은 하지 않음) */
  makeBlock: "xl" | "pandas" | "none";
}

export interface ImportResult {
  /** 추가된 시트 이름(순서대로). toSheet=false 또는 미지원 확장자면 [] */
  sheetNames: string[];
  /** 생성된 로드 블록 id. 만들지 않았으면 null */
  blockId: string | null;
  /** 감지된 텍스트 인코딩 (xlsx는 null) */
  encoding: string | null;
  /** 워커 FS 기록 promise — 호출부가 실패 toast에 쓴다. toFs=false면 null.
   *  기록은 클라이언트 캐시에 즉시 반영되어 부트 전이어도 ready 시 자동 기록된다 */
  fs: Promise<void> | null;
}

/**
 * 파일 하나를 불러온다 — 시트 추가 + 워커 FS 기록 + 로드 블록 생성(옵션별).
 * 시트 추가는 한 스토어 트랜잭션(= 한 undo 단계), 블록 생성은 별도 한 단계다.
 */
export async function importData(
  name: string,
  bytes: Uint8Array,
  opts: ImportOptions,
): Promise<ImportResult> {
  const base = name.replace(/^.*[/\\]/, "") || name;
  const stem = base.replace(/\.[^.]+$/, "");
  const lower = base.toLowerCase();
  const isExcel = lower.endsWith(".xlsx") || lower.endsWith(".xls");
  const isText = lower.endsWith(".csv") || lower.endsWith(".txt");
  const encoding = isExcel ? null : detectTextEncoding(bytes);

  // 워커 FS — 소스와 동일하게 같은 이름은 그대로 덮어쓴다
  let fs: Promise<void> | null = null;
  if (opts.toFs) {
    const { getRuntimeClient } = await import("@/lib/runtime/client");
    fs = getRuntimeClient().writeFile(base, bytes);
    fs.catch(() => {}); // 호출부가 fs로 다시 잡는다 — unhandled rejection 방지
    // 캐시는 동기 반영 — 상태 바 칩이 바로 갱신된다
    if (typeof window !== "undefined") window.dispatchEvent(new Event("pygrid:data-files"));
  }

  // 시트 추가 — CSV/XLSX만(기존 lib/io 파서 재사용). 그 외 확장자는 FS 전용
  const added: Sheet[] = [];
  if (opts.toSheet && (isExcel || isText)) {
    if (isExcel) {
      const { sheetsFromFileData } = await import("@/lib/io/xlsx");
      const parsed = sheetsFromFileData(bytes.slice().buffer as ArrayBuffer);
      for (const sh of parsed) {
        sh.name = parsed.length === 1 ? stem : `${stem}-${sh.name}`;
        added.push(sh);
      }
    } else {
      const { csvToSheet } = await import("@/lib/io/csv");
      added.push(csvToSheet(decodeSmart(bytes), stem));
    }
    const taken = new Set(useWorkbookStore.getState().workbook.sheets.map((s) => s.name));
    for (const sh of added) {
      sh.name = uniqueSheetName(sh.name, taken);
      taken.add(sh.name);
    }
    useWorkbookStore.setState((state) => {
      state.workbook.sheets.push(...added);
      state.activeSheetId = added[0].id;
      state.selection = null;
    });
  }

  // 로드 블록 — xl 참조는 생성된 시트가 있어야 한다. 없으면 pandas(FS 기록 시)로 강등
  let mode = opts.makeBlock;
  if (mode === "xl" && added.length === 0) mode = opts.toFs ? "pandas" : "none";
  let blockId: string | null = null;
  if (mode !== "none") {
    const code =
      mode === "xl"
        ? buildXlLoadCode(formatA1(usedRange(added[0]), added[0].name))
        : buildPandasLoadCode(base, encoding);
    blockId = createReferenceBlocks(null, [{ title: base, code }])[0] ?? null;
  }

  return { sheetNames: added.map((s) => s.name), blockId, encoding, fs };
}
