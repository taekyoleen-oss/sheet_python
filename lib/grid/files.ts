// 부록 P: 속성 창 탐색기의 순수 도우미 — 경로 계산, 작업 폴더 지정 코드, 파일 불러오기 코드.
// 모든 동작은 "코드"로 남는다: 작업 폴더는 os.chdir, 불러오기는 pd.read_* 블록.

/** 사용자 폴더 — Windows의 C:/Users/tklee처럼 보이게 한다. 탐색기는 이보다 위로 가지 않는다 */
export const HOME_DIR = "/Users/tklee";
/** Windows 기본 폴더 (부트마다 만든다) */
export const WIN_FOLDERS = ["Desktop", "Documents", "Downloads", "Music", "Pictures", "Videos"];
/** 작업 폴더 기본값 — 다운로드 */
export const DEFAULT_WORK_DIR = `${HOME_DIR}/Downloads`;

/** "/a/b/../c/" → "/a/c" (절대 경로 정규화) */
export function normPath(path: string): string {
  const out: string[] = [];
  for (const part of path.split("/")) {
    if (part === "" || part === ".") continue;
    if (part === "..") out.pop();
    else out.push(part);
  }
  return "/" + out.join("/");
}

export const joinPath = (dir: string, name: string): string => normPath(`${dir}/${name}`);

export const parentPath = (path: string): string => normPath(`${path}/..`);

/** 경로 → 빵부스러기 [{name, path}] (루트 "/" 포함). root 안쪽이면 root부터 */
export function breadcrumbs(path: string, root = "/"): { name: string; path: string }[] {
  const parts = normPath(path).split("/").filter(Boolean);
  const all = [
    { name: "/", path: "/" },
    ...parts.map((name, i) => ({ name, path: "/" + parts.slice(0, i + 1).join("/") })),
  ];
  const i = all.findIndex((b) => b.path === normPath(root));
  return i > 0 ? all.slice(i) : all;
}

/** 파이썬 문자열 리터럴 (JSON 문자열은 그대로 유효한 Python 문자열이다) */
const py = (s: string): string => JSON.stringify(s);

/** 작업 폴더 지정 코드 — 런타임 재부트(메모리 FS 초기화) 뒤에도 같은 코드로 다시 만든다 */
export const workDirCode = (dir: string): string =>
  `import os\nos.makedirs(${py(dir)}, exist_ok=True)\nos.chdir(${py(dir)})`;

/** 부트 코드 — Windows 기본 폴더를 만들고 작업 폴더로 이동 */
export const bootDirsCode = (dir: string): string =>
  `import os\nfor _d in ${JSON.stringify(WIN_FOLDERS)}:\n    os.makedirs(${py(HOME_DIR)} + "/" + _d, exist_ok=True)\ndel _d\n` +
  workDirCode(dir);

/** 작업 폴더 기준 상대 경로 (안쪽이면 상대, 아니면 절대) */
export function relPath(path: string, cwd: string): string {
  const p = normPath(path);
  const c = normPath(cwd);
  if (c === "/") return p.slice(1);
  return p.startsWith(c + "/") ? p.slice(c.length + 1) : p;
}

/** 파일 이름 → 파이썬 변수 이름 (확장자 제거, 영숫자·밑줄·한글, 숫자 시작이면 df_ 접두) */
export function varNameFor(fileName: string): string {
  const stem = fileName.replace(/\.[^.]+$/, "");
  let name = stem.replace(/[^A-Za-z0-9_가-힣]+/g, "_").replace(/^_+|_+$/g, "").toLowerCase();
  if (name === "" || /^\d/.test(name)) name = `df_${name}`;
  return name;
}

export type LoadKind = "csv" | "excel" | "json" | "parquet" | "pickle" | "text";

const KIND_BY_EXT: Record<string, LoadKind> = {
  csv: "csv",
  tsv: "csv",
  txt: "text",
  md: "text",
  py: "text",
  xlsx: "excel",
  xls: "excel",
  json: "json",
  parquet: "parquet",
  pkl: "pickle",
  pickle: "pickle",
};

export const loadKindOf = (fileName: string): LoadKind | null =>
  KIND_BY_EXT[fileName.split(".").pop()?.toLowerCase() ?? ""] ?? null;

/**
 * 파일 불러오기 코드 (사용자가 실행한다). 마지막 줄은 미리보기 식이라 실행하면 결과가 셀에 펼쳐지고,
 * 변수는 속성 창 변수 목록에 나타난다. CSV는 한국어 엑셀 저장(cp949)을 함께 시도한다.
 */
export function loadCode(filePath: string, cwd: string): { varName: string; code: string } | null {
  const fileName = filePath.split("/").pop() ?? filePath;
  const kind = loadKindOf(fileName);
  if (!kind) return null;
  const v = varNameFor(fileName);
  const p = py(relPath(filePath, cwd));
  const sep = /\.tsv$/i.test(fileName) ? ', sep="\\t"' : "";
  const body: Record<LoadKind, string> = {
    csv: [
      "import pandas as pd",
      "try:",
      `    ${v} = pd.read_csv(${p}${sep})`,
      "except UnicodeDecodeError:  # 한국어 엑셀에서 저장한 CSV",
      `    ${v} = pd.read_csv(${p}${sep}, encoding="cp949")`,
      `${v}.head(20)`,
    ].join("\n"),
    excel: ["import pandas as pd", `${v} = pd.read_excel(${p})`, `${v}.head(20)`].join("\n"),
    json: ["import pandas as pd", `${v} = pd.read_json(${p})`, `${v}.head(20)`].join("\n"),
    parquet: ["import pandas as pd", `${v} = pd.read_parquet(${p})`, `${v}.head(20)`].join("\n"),
    pickle: ["import pandas as pd", `${v} = pd.read_pickle(${p})`, v].join("\n"),
    text: [`with open(${p}, encoding="utf-8") as f:`, `    ${v} = f.read()`, `${v}[:500]`].join("\n"),
  };
  return { varName: v, code: body[kind] };
}

/** 바이트 → "12.3 KB" */
export function fmtSize(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 ** 2) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / 1024 ** 2).toFixed(1)} MB`;
}
