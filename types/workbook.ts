// 워크북 데이터 모델 — 설계서 §3.1. 저장(.pygrid.json)·IndexedDB·스토어가 공유하는 단일 스키마.

export type CellType = "n" | "s" | "b" | "d" | "e";

export interface Cell {
  v: string | number | boolean | null;
  t: CellType;
  /** 표시 서식 힌트 ('0.0%', '#,##0', 'yyyy-mm-dd') */
  f?: string;
  /**
   * spill 출처 표시 `"<blockId>:<outputId>"`. 있으면 직접 편집 잠김.
   * 구 워크북의 `"<blockId>"` 단독 표기는 로드 시 정규화된다 (부록 D.1).
   */
  src?: string;
  /**
   * 미니 수식 원문 (부록 I). `=` 포함 저장, `v`는 계산 결과 캐시
   * (`t:'n'` 숫자 또는 `t:'e'` + 오류 코드). 선택 필드라 기존 워크북 호환.
   */
  fx?: string;
  /**
   * 셀 서식 (부록 J.2): b=굵게, fs=글자 크기(px). 선택 필드 — 기존 워크북 호환.
   * .pygrid.json에만 보존 (TSV/HTML 복사·XLSX 내보내기는 값만).
   */
  st?: { b?: boolean; fs?: number };
}

export interface Sheet {
  id: string;
  name: string;
  rowCount: number;
  colCount: number;
  /** key "r:c" (0-based). 빈 셀은 저장하지 않음 */
  cells: Record<string, Cell>;
  colWidths?: Record<number, number>;
  frozenCols?: number;
}

export type OutputMode = "values" | "object";
export type IncludeIndex = "auto" | "always" | "never";

export interface CellRange {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

export type RunStatus = "ok" | "error" | "busy" | "spill";
export type ResultKind = "scalar" | "table" | "image" | "object";

export interface RunResult {
  status: RunStatus;
  kind?: ResultKind;
  shape?: [number, number];
  /** 상위 100행 미리보기 또는 repr */
  preview?: unknown;
  /** blobs 스토어 참조 (이미지 결과) */
  imageBlobId?: string;
  stdout: string;
  stderr: string;
  traceback?: string;
  /** 초보자용 한국어 오류 요약 */
  summaryKo?: string;
  spillRange?: CellRange;
  durationMs: number;
  ranAt: string;
}

/** 블록 종류 — 마크다운 블록은 실행되지 않고 문서·목차 용도로만 쓰인다 */
export type BlockKind = "code" | "markdown";

/** 실행 결과 중 무엇을 셀에 표시할지 (설계서 확장 §7.1) */
export interface OutputSelection {
  /** 출력할 전역 변수명. 없으면 마지막 표현식 값 */
  variable?: string;
  /** DataFrame 결과에서 표시할 열. 없으면 전체 열 */
  columns?: string[];
  /** 표시할 상위 행 수. 없으면 전체 행 */
  rowLimit?: number;
}

/** 한 블록의 출력 하나 — 결과의 일부를 원하는 셀 영역에 배치한다 (설계서 부록 D) */
export interface OutputBinding {
  id: string;
  /** 결과가 놓일 시트. 없으면 블록의 sheetId */
  sheetId?: string;
  anchor: { r: number; c: number };
  mode: OutputMode;
  includeIndex: IncludeIndex;
  /** 어떤 값을 표시할지 (변수·열·행) */
  selection?: OutputSelection;
  /** 카드 목록에 표시할 이름 (없으면 변수명 또는 "마지막 표현식") */
  label?: string;
  /** 이 출력의 마지막 실행 결과 */
  last?: RunResult;
}

export interface PyBlock {
  id: string;
  sheetId: string;
  anchor: { r: number; c: number };
  code: string;
  outputMode: OutputMode;
  /** DataFrame spill 시 index 포함 규칙 */
  includeIndex: IncludeIndex;
  last?: RunResult;
  /** 기본 'code'. 'markdown'이면 code/outputMode는 무시된다 */
  kind?: BlockKind;
  /** 목차에 표시되는 제목 (마크다운 블록은 본문 첫 헤딩에서 자동 추출 가능) */
  title?: string;
  /** kind==='markdown'일 때의 본문 */
  markdown?: string;
  /**
   * 코드 블록의 마크다운 설명 (부록 J.4). 헤더 아래 편집/미리보기 영역에 표시되고
   * 헤딩은 목차 서브 항목으로 병합된다. undefined = 영역 없음, "" = 빈 설명(편집 중).
   * 마크다운 블록에는 쓰지 않는다(본문이 곧 문서).
   */
  note?: string;
  /** 카드 접기 상태 (마크다운·코드·결과 전부 숨김) */
  collapsed?: boolean;
  /**
   * 결과를 시트 셀에 쓸지. false면 카드 아래 Python 결과로만 본다(새 블록 기본).
   * undefined는 이 필드 이전 워크북 — 시트로 보낸다(기존 동작 유지).
   */
  toSheet?: boolean;
  /** 출력 선택 (변수·열·행) — 레거시 단일 출력. 로드 시 outputs[0]로 정규화된다 */
  output?: OutputSelection;
  /**
   * 다중 출력. 정본이며 최소 1개를 가진다(로드 시 레거시 필드에서 정규화).
   * anchor/outputMode/includeIndex/output/last는 outputs[0]와 동기화된 레거시 뷰다.
   */
  outputs?: OutputBinding[];
}

export type CalcMode = "auto" | "manual";

export interface WorkbookSettings {
  timeoutSec: number;
  inferTypesOnPaste: boolean;
}

/** 정의된 이름 (부록 O.4) — 수식에서 `=SUM(보험료)`처럼 쓴다. ref는 시트 접두어 포함 절대 참조 */
export interface DefinedName {
  name: string;
  /** 예: "Sheet1!$A$2:$A$100" */
  ref: string;
}

export interface Workbook {
  id: string;
  /** 스키마 버전 (마이그레이션용) */
  version: 1;
  title: string;
  sheets: Sheet[];
  /**
   * 부록 P: Python 작업 폴더 (속성 창에서 지정). 런타임이 준비될 때마다
   * `os.makedirs(…); os.chdir(…)` 코드로 적용된다. 선택 필드 — 기존 워크북 호환.
   */
  workDir?: string;
  /** 정의된 이름 (선택 필드 — 기존 워크북 호환) */
  names?: DefinedName[];
  pyBlocks: PyBlock[];
  initScript: string;
  calcMode: CalcMode;
  settings: WorkbookSettings;
  createdAt: string;
  updatedAt: string;
}

/** cells 레코드 키 */
export const cellKey = (r: number, c: number): string => `${r}:${c}`;

export const parseCellKey = (key: string): { r: number; c: number } => {
  const i = key.indexOf(":");
  return { r: Number(key.slice(0, i)), c: Number(key.slice(i + 1)) };
};
