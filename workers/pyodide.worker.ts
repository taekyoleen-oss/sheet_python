// workers/pyodide.worker.ts — Pyodide 모듈 워커 (M3: 부트·REPL·스트리밍·인터럽트)
// 계약: lib/runtime/protocol.ts (문서: /output/runtime-protocol.md)
// 규칙: PyProxy를 postMessage로 넘기지 않는다 — JSON-safe 페이로드만.

import type { PyodideInterface } from "pyodide";

import {
  EXCEL_CODE_RE,
  type MainToWorker,
  type OutCell,
  type OutputItem,
  type PreviewPayload,
  type VariableInfo,
  type WorkerToMain,
} from "../lib/runtime/protocol";
import bootstrapPy from "../lib/runtime/py/bootstrap.py";
import convertPy from "../lib/runtime/py/convert.py";
import modelOutPy from "../lib/runtime/py/model_out.py";
import xlPy from "../lib/runtime/py/xl.py";
import { bootDirsCode, DEFAULT_WORK_DIR } from "../lib/grid/files";

// tsconfig lib이 dom이므로 webworker 전역을 좁은 타입으로 캐스팅해 쓴다
// Next는 청크 끝에 `_N_E = __webpack_exports__` (output.library 대입)를 붙인다.
// 클래식 워커(sloppy)에서는 암묵적 전역이 되지만 모듈 워커(strict)에서는 ReferenceError가
// 나므로 전역 프로퍼티를 미리 만들어 둔다. Next 업그레이드로 다시 깨지면
// client.ts의 ["type"] 해크와 같은 폴백: public/ 정적 워커로 전환한다.
(globalThis as Record<string, unknown>)._N_E = undefined;

const ctx = globalThis as unknown as {
  postMessage(msg: WorkerToMain, transfer?: Transferable[]): void;
  location: { origin: string };
  addEventListener(type: "message", fn: (ev: MessageEvent<MainToWorker>) => void): void;
};

const post = (msg: WorkerToMain, transfer?: Transferable[]): void =>
  ctx.postMessage(msg, transfer);

let pyodide: PyodideInterface | null = null;
let initScript = "";
/** 실행 중인 run/repl id. 유휴(부트·리셋 스크립트 출력)는 0 */
let currentRunId = 0;
/** 진행 중인 run의 stdout/stderr 사본 — result에 실어 카드 아래 결과로 보인다(스트리밍은 그대로) */
const runIo = { out: "", err: "" };
/** boot 완료 전에 setInterruptBuffer가 오면 보관했다가 적용 */
let pendingInterrupt: Uint8Array | null = null;

interface PyRunResult {
  ok: boolean;
  repr?: string | null;
  type?: string | null;
  etype?: string;
  msg?: string;
  tb?: string;
}

/** _pygrid_run 헬퍼로 코드 실행(마지막 표현식 repr 캡처) */
async function pyRun(py: PyodideInterface, code: string): Promise<PyRunResult> {
  py.globals.set("_pygrid_code", code);
  const raw: string = await py.runPythonAsync("_pygrid_run(_pygrid_code)");
  return JSON.parse(raw) as PyRunResult;
}

/** 엑셀 엔진(openpyxl) 지연 설치 — Pyodide 314 배포판은 엑셀 엔진 미선로드라 pd.read_excel이
 *  ImportError를 낸다. 코드가 엑셀 API를 쓰면 세션당 1회만 설치를 시도한다(프라미스 캐시).
 *  loadPackage(배포판) 실패 시 micropip(PyPI) 폴백. 실패해도 실행은 계속 — stderr로만 알린다.
 *  micropip은 pyimport로 다뤄 사용자 전역에 이름을 남기지 않는다(PyProxy는 워커 안에만). */
let excelSupport: Promise<void> | null = null;

function ensureExcelSupport(py: PyodideInterface, code: string): Promise<void> {
  if (!EXCEL_CODE_RE.test(code)) return Promise.resolve();
  excelSupport ??= (async () => {
    try {
      await py.loadPackage("openpyxl");
    } catch {
      try {
        await py.loadPackage("micropip");
        const micropip = py.pyimport("micropip") as { install(pkg: string): Promise<void> };
        await micropip.install("openpyxl");
      } catch (err) {
        post({
          t: "stderr",
          id: currentRunId,
          chunk: `엑셀 지원(openpyxl) 설치 실패 — read_excel/to_excel을 쓸 수 없습니다: ${
            err instanceof Error ? err.message : String(err)
          }`,
        });
      }
    }
  })();
  return excelSupport;
}

/** import 기반 패키지 지연 로드 + matplotlib이 로드됐다면 Agg·폰트 설정(멱등) */
async function loadImports(py: PyodideInterface, code: string): Promise<void> {
  await ensureExcelSupport(py, code); // 내부에서 모든 실패를 삼킨다 — 실행을 막지 않는다
  try {
    await py.loadPackagesFromImports(code);
  } catch {
    // 구문 오류 등은 _pygrid_run이 더 나은 트레이스백으로 보고한다
  }
  try {
    py.runPython("_pygrid_mpl_setup()");
  } catch {
    // 백엔드/폰트 설정 실패는 실행을 막지 않는다
  }
}

/** 초기화 스크립트 실행(부트·리셋 공용). 실패해도 부트는 계속한다 — stderr(id 0)로 알린다.
 *  주의: 여기서는 loadPackagesFromImports를 하지 않는다 — 기본 스크립트의 guarded
 *  `import matplotlib`을 부트 시 내려받지 않기 위해서다(지연 로드, CLAUDE.md §5).
 *  커스텀 초기화 스크립트가 추가 패키지를 import하면 ImportError가 stderr로 보고된다. */
async function runInitScript(py: PyodideInterface): Promise<void> {
  const r = await pyRun(py, initScript);
  if (!r.ok) {
    post({ t: "stderr", id: 0, chunk: `초기화 스크립트 오류:\n${r.tb ?? r.msg ?? ""}` });
  }
  try {
    py.runPython("_pygrid_mpl_setup()");
  } catch {
    // 무시
  }
  try {
    py.runPython("_pygrid_mark_baseline()"); // 부록 P: 변수 목록은 이 뒤에 생긴 것만
  } catch {
    // 무시
  }
}

async function boot(msg: Extract<MainToWorker, { t: "boot" }>): Promise<void> {
  if (pyodide) return; // 중복 부트 무시
  try {
    initScript = msg.initScript;

    post({ t: "progress", pct: 5, label: "스크립트 로드" });
    // webpackIgnore: npm pyodide는 번들에 넣지 않고 CDN(indexURL)에서 로드한다
    const mod = (await import(/* webpackIgnore: true */ msg.indexURL + "pyodide.mjs")) as {
      loadPyodide(opts: { indexURL: string }): Promise<PyodideInterface>;
    };

    post({ t: "progress", pct: 20, label: "런타임 초기화" });
    const py = await mod.loadPyodide({ indexURL: msg.indexURL });

    py.setStdout({
      batched: (chunk) => {
        if (currentRunId !== 0) runIo.out += chunk + "\n";
        post({ t: "stdout", id: currentRunId, chunk });
      },
    });
    py.setStderr({
      batched: (chunk) => {
        if (currentRunId !== 0) runIo.err += chunk + "\n";
        post({ t: "stderr", id: currentRunId, chunk });
      },
    });

    post({ t: "progress", pct: 45, label: "패키지 로드(numpy·pandas)" });
    if (msg.packages.length > 0) await py.loadPackage(msg.packages);

    post({ t: "progress", pct: 75, label: "폰트 등록" });
    try {
      const fontUrl = new URL(msg.fontUrl, ctx.location.origin).href;
      const resp = await fetch(fontUrl);
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const bytes = new Uint8Array(await resp.arrayBuffer());
      py.FS.mkdirTree("/fonts");
      py.FS.writeFile("/fonts/Pretendard-Regular.otf", bytes);
    } catch {
      post({ t: "stderr", id: 0, chunk: "한글 폰트 등록 실패 — 차트 한글이 깨질 수 있습니다" });
    }

    // _pygrid_* 헬퍼 + xl() 브리지 + §3.3 변환 정의 (리셋에서도 살아남는다)
    py.runPython(bootstrapPy);
    py.runPython(xlPy);
    py.runPython(convertPy);
    py.runPython(modelOutPy);
    // Windows 기본 폴더 + 기본 작업 폴더(다운로드) — ready 전에 옮겨야 캐시 파일 복원이 여기에 떨어진다
    py.runPython(bootDirsCode(DEFAULT_WORK_DIR));

    post({ t: "progress", pct: 90, label: "초기화 스크립트" });
    pyodide = py;
    if (pendingInterrupt) {
      py.setInterruptBuffer(pendingInterrupt);
      pendingInterrupt = null;
    }
    await runInitScript(py);

    post({ t: "progress", pct: 100, label: "완료" });
    post({
      t: "ready",
      pyVersion: py.runPython("__import__('sys').version.split()[0]") as string,
      pyodideVersion: py.version,
    });
  } catch (err) {
    post({ t: "bootError", message: err instanceof Error ? err.message : String(err) });
  }
}

async function handleRepl(msg: Extract<MainToWorker, { t: "repl" }>): Promise<void> {
  const py = pyodide;
  if (!py) {
    post({
      t: "replResult",
      id: msg.id,
      repr: null,
      stdout: "",
      stderr: "",
      traceback: "런타임이 아직 준비되지 않았습니다",
    });
    return;
  }
  currentRunId = msg.id;
  try {
    await loadImports(py, msg.code);
    const r = await pyRun(py, msg.code);
    post({
      t: "replResult",
      id: msg.id,
      repr: r.ok ? (r.repr ?? null) : null,
      stdout: "", // stdout/stderr는 이미 스트리밍됨
      stderr: "",
      ...(r.ok ? {} : { traceback: r.tb ?? r.msg ?? "실행 오류" }),
    });
  } catch (err) {
    post({
      t: "replResult",
      id: msg.id,
      repr: null,
      stdout: "",
      stderr: "",
      traceback: err instanceof Error ? err.message : String(err),
    });
  } finally {
    currentRunId = 0;
  }
}

/** convert.py _pygrid_run_convert 결과 */
interface PyConvertResult {
  ok: boolean;
  kind?: "scalar" | "table" | "image" | "object";
  typeName?: string;
  shape?: [number, number];
  cells?: OutCell[][];
  preview?: PreviewPayload;
  pngB64?: string;
  etype?: string;
  msg?: string;
  tb?: string;
}

function b64ToArrayBuffer(b64: string): ArrayBuffer {
  const bin = atob(b64);
  const u8 = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
  return u8.buffer;
}

/** _pygrid_run_convert_multi 결과: 성공이면 요청 순서대로 items */
interface PyMultiResult {
  ok: boolean;
  items?: (PyConvertResult & { id: string })[];
  etype?: string;
  msg?: string;
  tb?: string;
}

/** convert.py 출력 하나 → 프로토콜 OutputItem (PNG는 ArrayBuffer로 디코드) */
function toOutputItem(r: PyConvertResult & { id: string }): OutputItem {
  if (!r.ok) {
    return {
      id: r.id,
      ok: false,
      errorType: r.etype ?? "Exception",
      message: r.msg ?? "",
      traceback: r.tb ?? "",
    };
  }
  return {
    id: r.id,
    ok: true,
    kind: r.kind ?? "object",
    cells: r.cells,
    typeName: r.typeName,
    shape: r.shape,
    preview: r.preview,
    imagePng: r.pngB64 ? b64ToArrayBuffer(r.pngB64) : undefined,
  };
}

async function handleRun(msg: Extract<MainToWorker, { t: "run" }>): Promise<void> {
  const t0 = performance.now();
  const fail = (errorType: string, message: string, traceback: string): void =>
    post({
      t: "result",
      id: msg.id,
      blockId: msg.blockId,
      ok: false,
      errorType,
      message,
      traceback,
      stdout: runIo.out,
      stderr: runIo.err,
      durationMs: Math.round(performance.now() - t0),
    });
  const py = pyodide;
  if (!py) {
    fail("RuntimeNotReady", "런타임이 아직 준비되지 않았습니다", "");
    return;
  }
  currentRunId = msg.id;
  runIo.out = "";
  runIo.err = "";
  try {
    await loadImports(py, msg.code);
    // xl() 스냅샷 주입 → 실행+변환 → finally에서 캐시 비움 (계약: runtime-protocol.md)
    py.globals.set("_pygrid_snapshots", JSON.stringify(msg.snapshots));
    py.runPython("_pygrid_xl_load(_pygrid_snapshots)");
    py.globals.set("_pygrid_code", msg.code);
    if (msg.outputs) {
      // 다중 출력: 본문 1회 실행 + 출력별 변환 (출력 단위 실패는 그 출력만 실패)
      py.globals.set("_pygrid_outputs", JSON.stringify(msg.outputs));
      const rawMulti: string = await py.runPythonAsync(
        "_pygrid_run_convert_multi(_pygrid_code, _pygrid_outputs)",
      );
      const m = JSON.parse(rawMulti) as PyMultiResult;
      if (!m.ok) {
        fail(m.etype ?? "Exception", m.msg ?? "", m.tb ?? "");
        return;
      }
      const outputs = (m.items ?? []).map(toOutputItem);
      const transfer = outputs
        .map((o) => (o.ok ? o.imagePng : undefined))
        .filter((b): b is ArrayBuffer => b !== undefined);
      const first = outputs.find((o) => o.ok);
      post(
        {
          t: "result",
          id: msg.id,
          blockId: msg.blockId,
          ok: true,
          kind: first && first.ok ? first.kind : "object",
          outputs,
          stdout: runIo.out,
          stderr: runIo.err,
          durationMs: Math.round(performance.now() - t0),
        },
        transfer.length > 0 ? transfer : undefined,
      );
      return;
    }
    py.globals.set("_pygrid_output_mode", msg.outputMode);
    py.globals.set("_pygrid_include_index", msg.includeIndex);
    // 출력 선택은 JSON 문자열로 넘긴다(PyProxy 금지). 없으면 "null" → Python None
    py.globals.set("_pygrid_output_sel", JSON.stringify(msg.output ?? null));
    const raw: string = await py.runPythonAsync(
      "_pygrid_run_convert(_pygrid_code, _pygrid_output_mode, _pygrid_include_index, _pygrid_output_sel)",
    );
    const r = JSON.parse(raw) as PyConvertResult;
    const durationMs = Math.round(performance.now() - t0);
    if (!r.ok) {
      fail(r.etype ?? "Exception", r.msg ?? "", r.tb ?? "");
      return;
    }
    const imagePng = r.pngB64 ? b64ToArrayBuffer(r.pngB64) : undefined;
    post(
      {
        t: "result",
        id: msg.id,
        blockId: msg.blockId,
        ok: true,
        kind: r.kind ?? "object",
        typeName: r.typeName,
        shape: r.shape,
        cells: r.cells,
        preview: r.preview,
        imagePng,
        stdout: runIo.out,
        stderr: runIo.err,
        durationMs,
      },
      imagePng ? [imagePng] : undefined,
    );
  } catch (err) {
    fail("WorkerError", err instanceof Error ? err.message : String(err), "");
  } finally {
    try {
      py.runPython("_pygrid_xl_cache.clear()");
    } catch {
      // 무시
    }
    currentRunId = 0;
  }
}

function handleAnalyze(msg: Extract<MainToWorker, { t: "analyze" }>): void {
  const py = pyodide;
  if (!py) {
    post({ t: "analyzeError", id: msg.id, message: "런타임이 아직 준비되지 않았습니다" });
    return;
  }
  try {
    py.globals.set("_pygrid_code", msg.code);
    const raw = py.runPython("_pygrid_extract_refs(_pygrid_code)") as string;
    const r = JSON.parse(raw) as { ok: boolean; refs?: string[]; message?: string };
    if (r.ok) post({ t: "analyzed", id: msg.id, refs: r.refs ?? [] });
    else post({ t: "analyzeError", id: msg.id, message: r.message ?? "분석 오류" });
  } catch (err) {
    post({
      t: "analyzeError",
      id: msg.id,
      message: err instanceof Error ? err.message : String(err),
    });
  }
}

function handleInspect(msg: Extract<MainToWorker, { t: "inspect" }>): void {
  const py = pyodide;
  if (!py) {
    post({ t: "variables", id: msg.id, vars: [] });
    return;
  }
  try {
    const raw = py.runPython("_pygrid_inspect()") as string;
    post({ t: "variables", id: msg.id, vars: JSON.parse(raw) as VariableInfo[] });
  } catch {
    post({ t: "variables", id: msg.id, vars: [] });
  }
}

// ── 파일 I/O (R4) — path 검증은 클라이언트(writeFile/readFile)가 담당한다 ──

function handleWriteFile(msg: Extract<MainToWorker, { t: "writeFile" }>): void {
  const py = pyodide;
  if (!py) {
    post({ t: "fileError", id: msg.id, message: "런타임이 아직 준비되지 않았습니다" });
    return;
  }
  try {
    py.FS.writeFile(msg.path, new Uint8Array(msg.bytes));
    post({ t: "fileWritten", id: msg.id });
  } catch (err) {
    post({
      t: "fileError",
      id: msg.id,
      message: `파일 쓰기 실패(${msg.path}): ${err instanceof Error ? err.message : String(err)}`,
    });
  }
}

/** 부록 P: 폴더 목록 — 절대 경로로 정규화해 돌려준다 (Python 작업 폴더 = FS cwd) */
function handleListDir(msg: Extract<MainToWorker, { t: "listDir" }>): void {
  const py = pyodide;
  if (!py) {
    post({ t: "fileError", id: msg.id, message: "런타임이 아직 준비되지 않았습니다" });
    return;
  }
  const FS = py.FS;
  try {
    const cwd = FS.cwd();
    const path = FS.lookupPath(msg.path || cwd, {}).path || "/";
    const entries = FS.readdir(path)
      .filter((n) => n !== "." && n !== "..")
      .map((name) => {
        const st = FS.stat(`${path === "/" ? "" : path}/${name}`);
        const dir = FS.isDir(st.mode);
        return { name, dir, size: dir ? 0 : st.size };
      })
      .sort((a, b) => Number(b.dir) - Number(a.dir) || a.name.localeCompare(b.name));
    post({ t: "dirListing", id: msg.id, cwd, path, entries });
  } catch (err) {
    post({
      t: "fileError",
      id: msg.id,
      message: `폴더를 열 수 없습니다(${msg.path ?? "작업 폴더"}): ${err instanceof Error ? err.message : String(err)}`,
    });
  }
}

function handleReadFile(msg: Extract<MainToWorker, { t: "readFile" }>): void {
  const py = pyodide;
  if (!py) {
    post({ t: "fileError", id: msg.id, message: "런타임이 아직 준비되지 않았습니다" });
    return;
  }
  try {
    // Emscripten readFile은 새로 할당한 Uint8Array를 돌려준다 → buffer 통째로 transfer
    const bytes = py.FS.readFile(msg.path) as Uint8Array;
    const buf = bytes.buffer as ArrayBuffer;
    post({ t: "fileRead", id: msg.id, bytes: buf }, [buf]);
  } catch {
    post({ t: "fileError", id: msg.id, message: `파일을 찾을 수 없습니다: ${msg.path}` });
  }
}

async function handleReset(msg: Extract<MainToWorker, { t: "resetRuntime" }>): Promise<void> {
  const py = pyodide;
  if (py) {
    initScript = msg.initScript;
    try {
      py.runPython("_pygrid_reset()");
      await runInitScript(py);
    } catch {
      // best-effort — 완전 리셋은 클라이언트의 terminateAndReboot가 담당
    }
  }
  post({ t: "resetDone", id: msg.id });
}

async function dispatch(msg: MainToWorker): Promise<void> {
  switch (msg.t) {
    case "boot":
      return boot(msg);
    case "repl":
      return handleRepl(msg);
    case "run":
      return handleRun(msg);
    case "analyze":
      handleAnalyze(msg);
      return;
    case "inspect":
      handleInspect(msg);
      return;
    case "resetRuntime":
      return handleReset(msg);
    case "writeFile":
      handleWriteFile(msg);
      return;
    case "readFile":
      handleReadFile(msg);
      return;
    case "listDir":
      handleListDir(msg);
      return;
    default:
      return;
  }
}

// 메시지는 프라미스 체인으로 직렬 처리한다(run/repl 인터리빙 시 currentRunId가 섞이는 것 방지).
// setInterruptBuffer만 예외 — 실행 중에도 즉시 적용해야 하므로 체인에 넣지 않는다.
let chain: Promise<void> = Promise.resolve();

ctx.addEventListener("message", (ev) => {
  const msg = ev.data;
  if (msg.t === "setInterruptBuffer") {
    const view = new Uint8Array(msg.buffer);
    if (pyodide) pyodide.setInterruptBuffer(view);
    else pendingInterrupt = view;
    return;
  }
  chain = chain.then(() => dispatch(msg)).catch(() => undefined);
});
