// §3.3 변환 규칙 행 단위 검증 — Node용 Pyodide로 xl.py/convert.py를 실제 실행한다.
// 첫 실행은 numpy·pandas·matplotlib CDN 다운로드로 느리다(node_modules/.pyodide-cache에 캐시).
// 결과 요약 → output/conversion-report.json

import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

import { loadPyodide, type PyodideInterface } from "pyodide";
import { afterAll, beforeAll, expect, test } from "vitest";

import type { RangeSnapshot } from "@/lib/runtime/protocol";
import type { CellType, OutputSelection } from "@/types/workbook";

let py: PyodideInterface;
const results: { row: string; pass: boolean }[] = [];

/** §3.3 행 하나 = 테스트 하나. 결과를 conversion-report.json에 적재한다 */
const row = (name: string, fn: () => void | Promise<void>) =>
  test(name, async () => {
    try {
      await fn();
      results.push({ row: name, pass: true });
    } catch (e) {
      results.push({ row: name, pass: false });
      throw e;
    }
  });

const pyEval = (expr: string): unknown => py.runPython(expr);

const inject = (snapshots: Record<string, RangeSnapshot>) => {
  py.globals.set("_pygrid_snapshots", JSON.stringify(snapshots));
  py.runPython("_pygrid_xl_load(_pygrid_snapshots)");
};

interface ConvertResult {
  ok: boolean;
  kind?: string;
  typeName?: string;
  shape?: [number, number];
  cells?: { v: unknown; t: string; f?: string }[][];
  preview?: {
    kind: string;
    columns?: string[];
    dtypes?: string[];
    rows?: unknown[][];
    shape?: [number, number];
    repr?: string;
  };
  pngB64?: string;
  etype?: string;
  msg?: string;
}

const runConvert = (
  code: string,
  mode: "values" | "object" = "values",
  idx: "auto" | "always" | "never" = "auto",
  output: OutputSelection | null = null,
): ConvertResult => {
  py.globals.set("_pygrid_code", code);
  py.globals.set("_pygrid_output_mode", mode);
  py.globals.set("_pygrid_include_index", idx);
  // 워커와 같은 호출 형태 — 출력 선택은 JSON 문자열로 넘긴다
  py.globals.set("_pygrid_output_sel", JSON.stringify(output));
  return JSON.parse(
    py.runPython(
      "_pygrid_run_convert(_pygrid_code, _pygrid_output_mode, _pygrid_include_index, _pygrid_output_sel)",
    ) as string,
  ) as ConvertResult;
};

interface MultiResult {
  ok: boolean;
  items?: (ConvertResult & { id: string })[];
  etype?: string;
  msg?: string;
}

interface OutReq {
  id: string;
  mode?: "values" | "object";
  includeIndex?: "auto" | "always" | "never";
  selection?: OutputSelection;
}

/** 다중 출력: 워커와 같은 호출 형태 — 요청 배열을 JSON 문자열로 넘긴다 */
const runMulti = (code: string, outputs: OutReq[]): MultiResult => {
  py.globals.set("_pygrid_code", code);
  py.globals.set(
    "_pygrid_outputs",
    JSON.stringify(outputs.map((o) => ({ mode: "values", includeIndex: "auto", ...o }))),
  );
  return JSON.parse(
    py.runPython("_pygrid_run_convert_multi(_pygrid_code, _pygrid_outputs)") as string,
  ) as MultiResult;
};

const extractRefs = (code: string): { ok: boolean; refs?: string[]; message?: string } => {
  py.globals.set("_pygrid_code", code);
  return JSON.parse(py.runPython("_pygrid_extract_refs(_pygrid_code)") as string);
};

beforeAll(async () => {
  py = await loadPyodide({
    packageCacheDir: path.resolve("node_modules/.pyodide-cache"),
  });
  await py.loadPackage(["numpy", "pandas", "matplotlib"]);
  for (const f of ["bootstrap.py", "xl.py", "convert.py"]) {
    py.runPython(readFileSync(path.resolve("lib/runtime/py", f), "utf8"));
  }
  // Node에는 /fonts가 없으므로 폰트 등록은 건너뛰고 Agg 백엔드만 적용된다
  py.runPython("import matplotlib; _pygrid_mpl_setup()");
}, 600_000);

afterAll(() => {
  writeFileSync(
    path.resolve("output/conversion-report.json"),
    JSON.stringify(results, null, 2),
  );
});

// ── §3.3 입력 행 ─────────────────────────────────────────

row("입력 t:'n' — 열 전체 정수 → int64", () => {
  inject({ "A1:A3": { values: [[1], [2], [3]], types: [["n"], ["n"], ["n"]], scalar: false } });
  expect(pyEval('str(xl("A1:A3")[0].dtype)')).toBe("int64");
  expect(pyEval('list(xl("A1:A3")[0]) == [1, 2, 3]')).toBe(true);
});

row("입력 t:'n' — 소수 혼재 → float64", () => {
  inject({ "A1:A2": { values: [[1], [2.5]], types: [["n"], ["n"]], scalar: false } });
  expect(pyEval('str(xl("A1:A2")[0].dtype)')).toBe("float64");
});

row("입력 t:'s' → str(object)", () => {
  inject({ "B1:B2": { values: [["가"], ["나"]], types: [["s"], ["s"]], scalar: false } });
  expect(pyEval('str(xl("B1:B2")[0].dtype)')).toBe("object");
  expect(pyEval('list(xl("B1:B2")[0]) == ["가", "나"]')).toBe(true);
});

row("입력 t:'b' → bool", () => {
  inject({ "C1:C2": { values: [[true], [false]], types: [["b"], ["b"]], scalar: false } });
  expect(pyEval('str(xl("C1:C2")[0].dtype)')).toBe("bool");
});

row("입력 t:'d' → datetime64[ns] (Timestamp)", () => {
  inject({
    "D1:D2": { values: [["2026-09-02"], ["2026-09-03"]], types: [["d"], ["d"]], scalar: false },
  });
  expect(pyEval('str(xl("D1:D2")[0].dtype)')).toBe("datetime64[ns]");
  expect(pyEval('xl("D1:D2")[0].iloc[0].day')).toBe(2);
});

row("입력 빈 셀 → NaN/NaT/None (경계 케이스 #1)", () => {
  inject({
    "E1:E3": { values: [[1], [null], [3]], types: [["n"], ["s"], ["n"]], scalar: false },
    "F1:F2": { values: [["2026-09-02"], [null]], types: [["d"], ["s"]], scalar: false },
    "G1:G2": { values: [["가"], [null]], types: [["s"], ["s"]], scalar: false },
  });
  // 정수열 + 빈 셀 → int64 불가, float64 + NaN
  expect(pyEval('str(xl("E1:E3")[0].dtype)')).toBe("float64");
  expect(pyEval('import pandas as pd; bool(pd.isna(xl("E1:E3")[0].iloc[1]))')).toBe(true);
  expect(pyEval('bool(pd.isna(xl("F1:F2")[0].iloc[1]))')).toBe(true); // NaT
  expect(pyEval('xl("G1:G2")[0].iloc[1] is None')).toBe(true);
});

row("입력 단일 셀 xl('A1') → 스칼라", () => {
  inject({
    A1: { values: [[5]], types: [["n"]], scalar: true },
    D1: { values: [["2026-09-02"]], types: [["d"]], scalar: true },
    N1: { values: [[null]], types: [["s"]], scalar: true },
  });
  expect(pyEval('xl("A1")')).toBe(5);
  expect(pyEval('type(xl("D1")).__name__')).toBe("Timestamp");
  expect(pyEval('xl("N1") is None')).toBe(true);
});

// ── §3.3 출력 행 (값 모드) ───────────────────────────────

row("출력 스칼라 → 1셀 (None→빈 셀, bool→'b')", () => {
  expect(runConvert("42").cells).toEqual([[{ v: 42, t: "n" }]]);
  expect(runConvert("42").kind).toBe("scalar");
  expect(runConvert("None").cells).toEqual([[{ v: null, t: "s" }]]);
  expect(runConvert("True").cells).toEqual([[{ v: true, t: "b" }]]);
  expect(runConvert('"텍스트"').cells).toEqual([[{ v: "텍스트", t: "s" }]]);
});

row("출력 list/1D ndarray/Series → 세로 1열 (Series name은 헤더)", () => {
  expect(runConvert("[1, 2, 3]").cells).toEqual([
    [{ v: 1, t: "n" }],
    [{ v: 2, t: "n" }],
    [{ v: 3, t: "n" }],
  ]);
  expect(runConvert("import numpy as np\nnp.array([1.5, 2.5])").cells).toEqual([
    [{ v: 1.5, t: "n" }],
    [{ v: 2.5, t: "n" }],
  ]);
  const named = runConvert('import pandas as pd\npd.Series([1, 2], name="점수")');
  expect(named.cells).toEqual([[{ v: "점수", t: "s" }], [{ v: 1, t: "n" }], [{ v: 2, t: "n" }]]);
  const unnamed = runConvert("import pandas as pd\npd.Series([1, 2])");
  expect(unnamed.cells).toHaveLength(2);
});

row("출력 중첩 list/2D ndarray → 2D spill", () => {
  expect(runConvert('[[1, "a"], [2, "b"]]').cells).toEqual([
    [
      { v: 1, t: "n" },
      { v: "a", t: "s" },
    ],
    [
      { v: 2, t: "n" },
      { v: "b", t: "s" },
    ],
  ]);
  const nd = runConvert("import numpy as np\nnp.arange(4).reshape(2, 2)");
  expect(nd.shape).toEqual([2, 2]);
  expect(nd.kind).toBe("table");
});

row("출력 DataFrame → 헤더+값, includeIndex auto/always/never", () => {
  const code = 'import pandas as pd\npd.DataFrame({"a": [1, 2], "b": [3.5, None]})';
  const auto = runConvert(code);
  // 기본 RangeIndex → auto에서 제외 (경계 케이스 #3), NaN → 빈 셀
  expect(auto.cells?.[0]).toEqual([
    { v: "a", t: "s" },
    { v: "b", t: "s" },
  ]);
  expect(auto.cells?.[2]).toEqual([
    { v: 2, t: "n" },
    { v: null, t: "s" },
  ]);
  const always = runConvert(code, "values", "always");
  expect(always.cells?.[0]).toEqual([
    { v: "", t: "s" },
    { v: "a", t: "s" },
    { v: "b", t: "s" },
  ]);
  expect(always.cells?.[1]?.[0]).toEqual({ v: 0, t: "n" });

  const namedIdx =
    'import pandas as pd\npd.DataFrame({"a": [1, 2]}, index=pd.Index(["x", "y"], name="k"))';
  expect(runConvert(namedIdx).cells?.[0]).toEqual([
    { v: "k", t: "s" },
    { v: "a", t: "s" },
  ]); // 비기본 index → auto에서 포함
  expect(runConvert(namedIdx, "values", "never").cells?.[0]).toEqual([{ v: "a", t: "s" }]);
});

row("출력 dict → 2열(키·값)", () => {
  expect(runConvert('{"a": 1, "나": 2.5}').cells).toEqual([
    [
      { v: "a", t: "s" },
      { v: 1, t: "n" },
    ],
    [
      { v: "나", t: "s" },
      { v: 2.5, t: "n" },
    ],
  ]);
});

row("출력 datetime/Timestamp → t:'d' ISO + 서식", () => {
  expect(runConvert('import pandas as pd\npd.Timestamp("2026-09-02")').cells).toEqual([
    [{ v: "2026-09-02", t: "d", f: "yyyy-mm-dd" }],
  ]);
  expect(runConvert('import datetime\ndatetime.datetime(2026, 9, 2, 10, 30)').cells).toEqual([
    [{ v: "2026-09-02 10:30:00", t: "d", f: "yyyy-mm-dd hh:mm:ss" }],
  ]);
});

row("출력 Figure → 값 모드 오류 / 객체 모드 PNG dpi150", () => {
  const err = runConvert("import matplotlib.pyplot as plt\n_f = plt.figure()\n_f");
  expect(err.ok).toBe(false);
  expect(err.etype).toBe("PyGridImageError");
  expect(err.msg).toBe("이미지는 값으로 펼칠 수 없습니다");

  const img = runConvert(
    'import matplotlib.pyplot as plt\n_fig, _ax = plt.subplots()\n_ax.set_title("한글 제목")\n_ax.plot([1, 2], [3, 4])\n_fig',
    "object",
  );
  expect(img.ok).toBe(true);
  expect(img.kind).toBe("image");
  expect(img.preview).toEqual({ kind: "image" });
  expect((img.pngB64 ?? "").length).toBeGreaterThan(1000);
  // PNG 시그니처(base64 "iVBORw0KGgo")
  expect(img.pngB64?.startsWith("iVBORw0KGgo")).toBe(true);
});

row("출력 그 외 객체 → 값 모드 str() 1셀 / 객체 모드 repr 카드", () => {
  const vals = runConvert("{1, 2}");
  expect(vals.cells?.[0]?.[0]?.t).toBe("s");
  expect(vals.cells?.[0]?.[0]?.v).toBe("{1, 2}");
  const obj = runConvert("{1, 2}", "object");
  expect(obj.kind).toBe("object");
  expect(obj.typeName).toBe("set");
  expect(obj.preview?.kind).toBe("repr");
  expect(obj.preview?.repr).toBe("{1, 2}");
});

// ── 골든 G2 + analyze + 안전망 ───────────────────────────

function g2Snapshot(): RangeSnapshot {
  const values: (string | number | boolean | null)[][] = [
    ["id", "금액", "비율", "날짜", "이름", "메모"],
  ];
  const types: CellType[][] = [["s", "s", "s", "s", "s", "s"]];
  for (let i = 1; i <= 20; i++) {
    const noRatio = i % 7 === 0;
    const noDate = i % 9 === 0;
    const noMemo = i % 3 === 0;
    values.push([
      i,
      i * 1000 + 0.5,
      noRatio ? null : i / 100,
      noDate ? null : `2026-01-${String(i).padStart(2, "0")}`,
      `이름${i}`,
      noMemo ? null : `메모${i}`,
    ]);
    types.push(["n", "n", noRatio ? "s" : "n", noDate ? "s" : "d", "s", "s"]);
  }
  return { values, types, scalar: false };
}

row("G2: xl('A1:F21', headers=True) 골든 dtype + NaN/NaT", () => {
  inject({ "A1:F21": g2Snapshot() });
  py.runPython('_g2 = xl("A1:F21", headers=True)');
  expect(pyEval("_g2.shape == (20, 6)")).toBe(true);
  expect(pyEval('str(_g2["id"].dtype)')).toBe("int64");
  expect(pyEval('str(_g2["금액"].dtype)')).toBe("float64");
  expect(pyEval('str(_g2["비율"].dtype)')).toBe("float64");
  expect(pyEval('str(_g2["날짜"].dtype)')).toBe("datetime64[ns]");
  expect(pyEval('str(_g2["이름"].dtype)')).toBe("object");
  expect(pyEval('int(_g2["비율"].isna().sum())')).toBe(2); // 7, 14
  expect(pyEval('int(_g2["날짜"].isna().sum())')).toBe(2); // 9, 18
  expect(pyEval('_g2["메모"].iloc[2] is None')).toBe(true); // 3행
});

row("extract_refs: 리터럴 추출·중복 제거·순서 유지", () => {
  const r = extractRefs('a = xl("A1")\nb = xl("B1:B5", headers=True)\nc = xl("A1")');
  expect(r).toEqual({ ok: true, refs: ["A1", "B1:B5"] });
  expect(extractRefs("print(1)")).toEqual({ ok: true, refs: [] });
});

row("sheet(): 정식 이름 — xl()과 같은 함수, 참조 추출·비리터럴 오류도 같은 규칙", () => {
  inject({ A1: { values: [[7]], types: [["n"]], scalar: true } });
  expect(pyEval('sheet("A1") == 7 and sheet is xl')).toBe(true);
  expect(extractRefs('a = sheet("A1")\nb = xl("B1:B2")')).toEqual({ ok: true, refs: ["A1", "B1:B2"] });
  const r = extractRefs('ref = "A1"\nsheet(ref)');
  expect(r.ok).toBe(false);
  expect(r.message).toBe("sheet() 인수는 문자열 리터럴이어야 합니다");
});

row("extract_refs: 비리터럴 인수 → 한국어 오류 (§2.4)", () => {
  const r = extractRefs('ref = "A1"\nxl(ref)');
  expect(r.ok).toBe(false);
  expect(r.message).toBe("xl() 인수는 문자열 리터럴이어야 합니다");
  expect(extractRefs('xl(f"A{n}")').ok).toBe(false);
  expect(extractRefs("xl()").ok).toBe(false);
  const h = extractRefs('xl("A1", headers=h)');
  expect(h.ok).toBe(false);
  expect(h.message).toBe("xl() headers 인수는 True/False 리터럴이어야 합니다");
});

row("빈 결과(0열 DataFrame·빈 list) → 값 모드 1×1 빈 셀 보장", () => {
  const emptyDf = runConvert("import pandas as pd\npd.DataFrame()");
  expect(emptyDf.cells).toEqual([[{ v: null, t: "s" }]]);
  expect(emptyDf.shape).toEqual([1, 1]);
  expect(runConvert("[]").cells).toEqual([[{ v: null, t: "s" }]]);
  expect(runConvert("import numpy as np\nnp.array([])").cells).toEqual([[{ v: null, t: "s" }]]);
});

row("초 미만 datetime → 초 단위 절단 (f 서식과 일치, #10)", () => {
  expect(
    runConvert('import datetime\ndatetime.datetime(2026, 9, 2, 10, 30, 15, 123456)').cells,
  ).toEqual([[{ v: "2026-09-02 10:30:15", t: "d", f: "yyyy-mm-dd hh:mm:ss" }]]);
});

row("들쭉날쭉한 중첩 list → 빈 셀 패딩 + 직사각 shape (#11)", () => {
  const r = runConvert("[[1, 2], [3]]");
  expect(r.shape).toEqual([2, 2]);
  expect(r.cells).toEqual([
    [
      { v: 1, t: "n" },
      { v: 2, t: "n" },
    ],
    [
      { v: 3, t: "n" },
      { v: null, t: "s" },
    ],
  ]);
});

row("xl 별칭 호출 → 분석 누락 + 런타임 안전망 (#14, parity #4)", () => {
  const code = 'f = xl\nf("H1")';
  expect(extractRefs(code)).toEqual({ ok: true, refs: [] }); // 별칭은 ast 분석에서 빠진다
  py.runPython("_pygrid_xl_cache.clear()");
  const r = runConvert(code);
  expect(r.ok).toBe(false);
  expect(r.etype).toBe("RuntimeError"); // 안전망
});

row("xl 미주입 참조 → RuntimeError 안전망", () => {
  py.runPython("_pygrid_xl_cache.clear()");
  const r = runConvert('xl("Z9")');
  expect(r.ok).toBe(false);
  expect(r.etype).toBe("RuntimeError");
  expect(r.msg).toContain("준비되지 않았습니다");
});

row("객체 모드 DataFrame → table preview 상위 100행 + dtypes + NaN null", () => {
  const r = runConvert(
    'import pandas as pd\nimport numpy as np\npd.DataFrame({"x": np.arange(120.0), "y": [np.nan] * 120})',
    "object",
  );
  expect(r.kind).toBe("table");
  expect(r.typeName).toBe("DataFrame");
  expect(r.shape).toEqual([120, 2]);
  expect(r.preview?.kind).toBe("table");
  expect(r.preview?.columns).toEqual(["x", "y"]);
  expect(r.preview?.dtypes).toEqual(["float64", "float64"]);
  expect(r.preview?.rows).toHaveLength(100);
  expect(r.preview?.rows?.[0]).toEqual([0, null]);
});

row("객체 모드 preview — 행 이름(기본 번호가 아닌 인덱스)은 맨 앞 열, MultiIndex는 수준마다", () => {
  const corr = runConvert(
    'import pandas as pd\npd.DataFrame({"a": [1.0, 2.0, 4.0], "b": [3.0, 1.0, 0.0]}).corr()',
    "object",
  );
  expect(corr.preview?.columns).toEqual(["", "a", "b"]);
  expect(["object", "str"]).toContain(corr.preview?.dtypes?.[0]); // pandas 3는 문자열 dtype이 str
  expect(corr.preview?.rows?.map((r) => r[0])).toEqual(["a", "b"]);
  expect(corr.shape).toEqual([2, 2]); // 데이터 모양은 그대로
  const multi = runConvert(
    'import pandas as pd\npd.DataFrame({"v": [1, 2]}, index=pd.MultiIndex.from_tuples([("x", 1), ("y", 2)], names=["k", "n"]))',
    "object",
  );
  expect(multi.preview?.columns).toEqual(["k", "n", "v"]);
  expect(multi.preview?.rows?.[1]).toEqual(["y", 2, 2]);
  const plain = runConvert('import pandas as pd\npd.DataFrame({"v": [1, 2]})', "object");
  expect(plain.preview?.columns).toEqual(["v"]); // 기본 번호는 붙이지 않는다
});

// ── 출력 선택(v1.1): variable · columns · rowLimit ────────

row("출력 선택 variable — 지정 전역 변수 (없으면 NameError, #15)", () => {
  const code = [
    "import pandas as pd",
    'sel_df = pd.DataFrame({"a": [1, 2]})',
    "sel_last = 99",
    "sel_last",
  ].join("\n");
  expect(runConvert(code).cells).toEqual([[{ v: 99, t: "n" }]]); // 선택 없으면 마지막 표현식
  const picked = runConvert(code, "values", "auto", { variable: "sel_df" });
  expect(picked.cells).toEqual([
    [{ v: "a", t: "s" }],
    [{ v: 1, t: "n" }],
    [{ v: 2, t: "n" }],
  ]);
  const miss = runConvert(code, "values", "auto", { variable: "없는변수" });
  expect(miss.ok).toBe(false);
  expect(miss.etype).toBe("NameError"); // errors-ko가 한국어 요약을 붙인다
  expect(miss.msg).toContain("없는변수");
});

row("출력 선택 columns — 요청 순서 부분집합, 없는 열 무시, 전부 없으면 전체 (#16)", () => {
  const code = 'import pandas as pd\npd.DataFrame({"a": [1, 2], "b": [3, 4], "c": [5, 6]})';
  const sub = runConvert(code, "values", "auto", { columns: ["c", "a"] });
  expect(sub.cells?.[0]).toEqual([
    { v: "c", t: "s" },
    { v: "a", t: "s" },
  ]); // 요청 순서 그대로
  expect(sub.cells?.[1]).toEqual([
    { v: 5, t: "n" },
    { v: 1, t: "n" },
  ]);
  const partial = runConvert(code, "values", "auto", { columns: ["b", "없는열"] });
  expect(partial.cells?.[0]).toEqual([{ v: "b", t: "s" }]); // 없는 열은 조용히 무시
  const allUnknown = runConvert(code, "values", "auto", { columns: ["x", "y"] });
  expect(allUnknown.cells?.[0]).toEqual([
    { v: "a", t: "s" },
    { v: "b", t: "s" },
    { v: "c", t: "s" },
  ]); // 전부 없으면 전체 열 — 빈 spill 방지
  // 비 DataFrame은 columns를 무시한다
  expect(runConvert("[1, 2]", "values", "auto", { columns: ["a"] }).shape).toEqual([2, 1]);
});

row("출력 선택 rowLimit — 상위 N 데이터 행(헤더 제외), 0 이하는 무제한 (#17)", () => {
  const df = 'import pandas as pd\npd.DataFrame({"a": [1, 2, 3, 4], "b": [5, 6, 7, 8]})';
  const r = runConvert(df, "values", "auto", { rowLimit: 2 });
  expect(r.shape).toEqual([3, 2]); // 헤더 1행 + 데이터 2행
  expect(r.cells?.[0]).toEqual([
    { v: "a", t: "s" },
    { v: "b", t: "s" },
  ]);
  expect(r.cells?.[2]?.[0]).toEqual({ v: 2, t: "n" });
  // Series name 헤더도 데이터 행에 포함되지 않는다
  expect(
    runConvert('import pandas as pd\npd.Series([1, 2, 3], name="점수")', "values", "auto", {
      rowLimit: 2,
    }).cells,
  ).toEqual([[{ v: "점수", t: "s" }], [{ v: 1, t: "n" }], [{ v: 2, t: "n" }]]);
  // 평범한 list · 2D ndarray
  expect(runConvert("[1, 2, 3, 4, 5]", "values", "auto", { rowLimit: 2 }).cells).toEqual([
    [{ v: 1, t: "n" }],
    [{ v: 2, t: "n" }],
  ]);
  expect(
    runConvert("import numpy as np\nnp.arange(9).reshape(3, 3)", "values", "auto", {
      rowLimit: 1,
    }).shape,
  ).toEqual([1, 3]);
  expect(runConvert(df, "values", "auto", { rowLimit: 0 }).shape).toEqual([5, 2]); // 0 → 무제한
  expect(runConvert("42", "values", "auto", { rowLimit: 1 }).cells).toEqual([
    [{ v: 42, t: "n" }],
  ]); // 스칼라는 무시
});

row("출력 선택 — 객체 모드 preview·shape에도 같은 선택 반영 (#18)", () => {
  const code = [
    "import pandas as pd",
    'sel_obj = pd.DataFrame({"a": range(10), "b": range(10), "c": range(10)})',
    "None",
  ].join("\n");
  const r = runConvert(code, "object", "auto", {
    variable: "sel_obj",
    columns: ["b", "a"],
    rowLimit: 3,
  });
  expect(r.ok).toBe(true);
  expect(r.kind).toBe("table");
  expect(r.typeName).toBe("DataFrame");
  expect(r.shape).toEqual([3, 2]);
  expect(r.preview?.columns).toEqual(["b", "a"]);
  expect(r.preview?.shape).toEqual([3, 2]);
  expect(r.preview?.rows).toHaveLength(3);
});

// ── 다중 출력(v1.2): 코드 1회 실행 · 출력별 변환 · 출력 단위 실패 ──

row("다중 출력 — 본문은 1회만 실행, 출력마다 독립 선택 (#20)", () => {
  const code = [
    "import pandas as pd",
    'mo_calls = globals().get("mo_calls", 0) + 1',
    'mo_df = pd.DataFrame({"a": [1, 2], "b": [3, 4]})',
    "mo_total = 10",
    "mo_total",
  ].join("\n");
  py.runPython("globals().pop('mo_calls', None)");
  const r = runMulti(code, [
    { id: "o1", selection: { variable: "mo_df" } },
    { id: "o2", selection: { variable: "mo_total" } },
    { id: "o3" }, // 선택 없음 → 마지막 표현식
  ]);
  expect(r.ok).toBe(true);
  expect(r.items?.map((i) => i.id)).toEqual(["o1", "o2", "o3"]); // 요청 순서 유지
  expect(pyEval("mo_calls")).toBe(1); // 출력이 3개여도 본문 실행은 1회
  expect(r.items?.[0].cells).toEqual([
    [
      { v: "a", t: "s" },
      { v: "b", t: "s" },
    ],
    [
      { v: 1, t: "n" },
      { v: 3, t: "n" },
    ],
    [
      { v: 2, t: "n" },
      { v: 4, t: "n" },
    ],
  ]);
  expect(r.items?.[1].cells).toEqual([[{ v: 10, t: "n" }]]);
  expect(r.items?.[2].cells).toEqual([[{ v: 10, t: "n" }]]);
});

row("다중 출력 — 한 출력의 실패(없는 변수)는 그 출력만 (#21)", () => {
  const r = runMulti("mo_ok = 7\nmo_ok", [
    { id: "bad", selection: { variable: "없는변수" } },
    { id: "good", selection: { variable: "mo_ok" } },
  ]);
  expect(r.ok).toBe(true); // 코드 본문은 성공 → run 전체는 성공
  expect(r.items?.[0].ok).toBe(false);
  expect(r.items?.[0].etype).toBe("NameError");
  expect(r.items?.[0].msg).toContain("없는변수");
  expect(r.items?.[1].ok).toBe(true);
  expect(r.items?.[1].cells).toEqual([[{ v: 7, t: "n" }]]);
});

row("다중 출력 — 값 모드 이미지 실패는 그 출력만, 객체 모드 출력은 PNG (#22)", () => {
  const code = [
    "import matplotlib.pyplot as plt",
    "mo_fig, mo_ax = plt.subplots()",
    "mo_ax.plot([1, 2], [3, 4])",
    "mo_n = 5",
    "mo_n",
  ].join("\n");
  const r = runMulti(code, [
    { id: "img_values", mode: "values", selection: { variable: "mo_fig" } },
    { id: "img_card", mode: "object", selection: { variable: "mo_fig" } },
    { id: "num" },
  ]);
  expect(r.ok).toBe(true);
  expect(r.items?.[0].ok).toBe(false);
  expect(r.items?.[0].etype).toBe("PyGridImageError");
  expect(r.items?.[0].msg).toBe("이미지는 값으로 펼칠 수 없습니다");
  expect(r.items?.[1].ok).toBe(true);
  expect(r.items?.[1].kind).toBe("image");
  expect(r.items?.[1].pngB64?.startsWith("iVBORw0KGgo")).toBe(true);
  expect(r.items?.[2].cells).toEqual([[{ v: 5, t: "n" }]]);
});

row("다중 출력 — columns·rowLimit는 출력마다 따로 적용 (#20)", () => {
  const code = [
    "import pandas as pd",
    'mo_wide = pd.DataFrame({"a": range(5), "b": range(5), "c": range(5)})',
    "None",
  ].join("\n");
  const r = runMulti(code, [
    { id: "x", selection: { variable: "mo_wide", columns: ["c", "a"], rowLimit: 2 } },
    { id: "y", selection: { variable: "mo_wide", rowLimit: 1 } },
    { id: "z", mode: "object", selection: { variable: "mo_wide", columns: ["b"] } },
  ]);
  expect(r.items?.[0].shape).toEqual([3, 2]); // 헤더 + 2행, 요청 열 순서
  expect(r.items?.[0].cells?.[0]).toEqual([
    { v: "c", t: "s" },
    { v: "a", t: "s" },
  ]);
  expect(r.items?.[1].shape).toEqual([2, 3]); // 같은 변수, 다른 선택
  expect(r.items?.[2].preview?.columns).toEqual(["b"]);
});

row("다중 출력 — 코드 본문 오류는 실행 전체 실패 (items 없음)", () => {
  const r = runMulti('raise ValueError("폭발")', [{ id: "a" }, { id: "b" }]);
  expect(r.ok).toBe(false);
  expect(r.etype).toBe("ValueError");
  expect(r.msg).toBe("폭발");
  expect(r.items).toBeUndefined();
});
