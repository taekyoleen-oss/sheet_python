// 워커 메시지 계약 — 이 파일이 유일한 계약이다 (CLAUDE.md §3).
// 변경 시 /output/runtime-protocol.md를 함께 갱신한다.
// 원칙: PyProxy·비직렬화 객체 금지. 이미지 ArrayBuffer는 transferable로 넘긴다.

import type {
  CellType,
  IncludeIndex,
  OutputMode,
  OutputSelection,
} from "@/types/workbook";

/** sheet() 참조 범위의 2D 스냅샷. 단일 셀 참조도 1×1 2D로 전달하고, 스칼라 변환은 xl.py가 참조 형태를 보고 결정한다 */
export interface RangeSnapshot {
  values: (string | number | boolean | null)[][];
  types: CellType[][];
  /** 참조가 단일 셀(`A1`)이면 true → sheet()이 스칼라 반환 */
  scalar: boolean;
}

/** 워커가 값 모드 결과로 돌려주는 셀. converters.ts가 그대로 Cell로 감싼다 */
export interface OutCell {
  v: string | number | boolean | null;
  t: CellType;
  f?: string;
}

export type PreviewPayload =
  | {
      kind: "table";
      columns: string[];
      dtypes: string[];
      /** 상위 100행 */
      rows: (string | number | boolean | null)[][];
      shape: [number, number];
    }
  | { kind: "repr"; repr: string }
  | { kind: "image" };

/** 다중 출력 요청 — 코드는 1회만 실행하고 출력마다 변환한다 */
export interface OutputRequest {
  /** PyBlock.outputs[].id */
  id: string;
  mode: OutputMode;
  includeIndex: IncludeIndex;
  selection?: OutputSelection;
}

export interface OutputItemSuccess {
  id: string;
  ok: true;
  kind: "scalar" | "table" | "image" | "object";
  /** 값 모드 결과 */
  cells?: OutCell[][];
  typeName?: string;
  shape?: [number, number];
  preview?: PreviewPayload;
  imagePng?: ArrayBuffer;
}

/** 출력 단위 실패(예: 지정 변수 없음) — 코드 자체는 성공한 경우 */
export interface OutputItemFailure {
  id: string;
  ok: false;
  errorType: string;
  message: string;
  traceback?: string;
}

export type OutputItem = OutputItemSuccess | OutputItemFailure;

export interface RunSuccess {
  ok: true;
  /** 다중 출력 결과. run에 outputs를 보낸 경우 채워진다 */
  outputs?: OutputItem[];
  kind: "scalar" | "table" | "image" | "object";
  /** 값(spill) 모드 결과. 객체 모드에서는 없음 */
  cells?: OutCell[][];
  /** 객체 카드 요약용 타입명 (예: "DataFrame") */
  typeName?: string;
  shape?: [number, number];
  preview?: PreviewPayload;
  /** PNG bytes (transferable) */
  imagePng?: ArrayBuffer;
  stdout: string;
  stderr: string;
  durationMs: number;
}

export interface RunFailure {
  ok: false;
  /** Python 예외 클래스명 (NameError 등). 한국어 요약 매핑 키 */
  errorType: string;
  message: string;
  traceback: string;
  stdout: string;
  stderr: string;
  durationMs: number;
}

export type RunPayload = RunSuccess | RunFailure;

export interface VariableInfo {
  name: string;
  type: string;
  shape?: [number, number];
  /** repr 첫 줄 등 짧은 요약 */
  summary?: string;
  /** 2차원 표(DataFrame)의 열 이름 — 최대 500개 (부록 O.5 목표 열 선택) */
  columns?: string[];
  /** 적합된 모델(statsmodels 결과·scikit-learn 추정기)이면 시트로 보낼 항목 카탈로그 */
  model?: ModelInfo;
}

/** 시트로 보낼 수 있는 모델 항목 하나 — 모델 종류에 따라 목록이 달라진다(부록 O.5) */
export interface ModelMember {
  /** 묶음 (요약표·적합 통계·계수·추론·적합값·잔차·교차검증·군집 …) */
  group: string;
  label: string;
  /** OutputSelection.variable에 그대로 들어가는 출력 식 (블록 전역에서 평가) */
  expr: string;
  /** 사용자에게 보여 주는 라이브러리 원래 코드 (코드로 옮겨 쓸 때) */
  code: string;
  /** 현재 값 미리보기 (스칼라는 값, 표·벡터는 크기와 앞부분) */
  preview: string;
  /** 예상 spill 크기 [행, 열] (index 포함 추정치 — 배치 간격 계산용) */
  shape: [number, number];
}

export interface ModelInfo {
  kind: "statsmodels" | "sklearn";
  /** 공식(formula) API로 적합 — 예측 데이터에 원본 열 이름만 있으면 된다 */
  formula: boolean;
  /** 신뢰·예측구간 지원 (statsmodels get_prediction) */
  intervals: boolean;
  /** predict 지원 (PCA 등 비지도 변환기는 없다) */
  predict: boolean;
  /** 클래스 확률 predict_proba 지원 (분류 모델) */
  proba: boolean;
  /** scikit-learn 모델이 열 이름(feature_names_in_)으로 학습됨 */
  featureNames: boolean;
  members: ModelMember[];
}

// ── 메인 → 워커 ──────────────────────────────────────────

export type MainToWorker =
  | {
      t: "boot";
      indexURL: string;
      /** 부트 시 선로드 패키지 (numpy, pandas) */
      packages: string[];
      initScript: string;
      /** matplotlib 한글 폰트 URL (Pyodide FS에 기록) */
      fontUrl: string;
    }
  | { t: "setInterruptBuffer"; buffer: SharedArrayBuffer }
  | { t: "analyze"; id: number; code: string }
  | {
      t: "run";
      id: number;
      blockId: string;
      code: string;
      snapshots: Record<string, RangeSnapshot>;
      outputMode: OutputMode;
      includeIndex: IncludeIndex;
      /** 출력 선택: 마지막 표현식 대신 특정 변수 / DataFrame 열·행 제한 (레거시 단일 출력) */
      output?: OutputSelection;
      /** 다중 출력 요청. 있으면 outputMode/includeIndex/output 대신 이쪽을 쓴다 */
      outputs?: OutputRequest[];
    }
  | { t: "repl"; id: number; code: string }
  | { t: "inspect"; id: number }
  | { t: "resetRuntime"; id: number; initScript: string }
  /** Pyodide FS에 파일 기록. path는 경로 없는 파일 이름만(클라이언트가 검증). bytes는 transferable */
  | { t: "writeFile"; id: number; path: string; bytes: ArrayBuffer }
  /** Pyodide FS에서 파일 읽기 → fileRead(bytes transferable) / 없으면 fileError */
  | { t: "readFile"; id: number; path: string }
  /** 속성 창 탐색기(부록 P): 폴더 목록. path 생략 시 Python 작업 폴더(os.getcwd()) → dirListing */
  | { t: "listDir"; id: number; path?: string };

// ── 워커 → 메인 ──────────────────────────────────────────

export type WorkerToMain =
  | { t: "progress"; pct: number; label: string }
  | { t: "ready"; pyVersion: string; pyodideVersion: string }
  | { t: "bootError"; message: string }
  | { t: "analyzed"; id: number; refs: string[] }
  | { t: "analyzeError"; id: number; message: string }
  | { t: "stdout"; id: number; chunk: string }
  | { t: "stderr"; id: number; chunk: string }
  | ({ t: "result"; id: number; blockId: string } & RunPayload)
  | {
      t: "replResult";
      id: number;
      repr: string | null;
      stdout: string;
      stderr: string;
      traceback?: string;
    }
  | { t: "variables"; id: number; vars: VariableInfo[] }
  | { t: "resetDone"; id: number }
  | { t: "fileWritten"; id: number }
  | { t: "fileRead"; id: number; bytes: ArrayBuffer }
  | { t: "dirListing"; id: number; cwd: string; path: string; entries: DirEntry[] }
  /** writeFile/readFile/listDir 실패(파일 없음·FS 오류 등). 한국어 메시지 */
  | { t: "fileError"; id: number; message: string };

/** 폴더 항목 (listDir) */
export interface DirEntry {
  name: string;
  dir: boolean;
  /** 바이트 (폴더는 0) */
  size: number;
}

export const DEFAULT_PYODIDE_INDEX_URL =
  "https://cdn.jsdelivr.net/pyodide/v314.0.6/full/";

export const BOOT_PACKAGES = ["numpy", "pandas"];

/** 코드가 엑셀 API를 쓰는지 감지 — Pyodide 314 배포판에 엑셀 엔진이 선로드되지 않아
 *  워커가 이 게이트에 걸리면 openpyxl을 지연 설치한다(loadPackage → micropip 폴백, 세션당 1회) */
export const EXCEL_CODE_RE = /read_excel|to_excel|ExcelWriter|\.xlsx|\.xls\b/;

/** 인터럽트: 버퍼[0]=2 (SIGINT) */
export const INTERRUPT_SIGINT = 2;
