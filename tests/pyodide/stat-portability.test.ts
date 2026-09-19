// 부록 N — "다른 표로 바꿀 때 고치는 곳은 머리말 네 줄뿐"이라는 약속을 실제로 지키는지 본다.
//
// 통계분석 조각 27개를 서로 구조가 다른 4개 표에 대해 전부 실행한다.
// 바꾸는 것은 STAT_HEAD(df·TARGET·CAT·NUM)뿐이고 조각 본문은 손대지 않는다.
//
//  wage        534행 11열  범주형이 전부 '숫자 코드', 결측 없음
//  policy      600행 16열  문자열 범주형 + ID 2개 + bool + 결측 75개 · **선형 신호 없음**
//  claims      600행 11열  문자열 범주형 + ID, 강한 선형 신호(prem_after ~ prem_before)
//  experience  800행  6열  설명변수가 3개뿐(수치형 1개) — 최소 구성
//
// 이 표들이 함께 걸러 주는 것: ID 열 혼입, 문자열 범주형, 결측, 단위 스케일 차이,
// NUM이 1개뿐일 때(파생변수·다항 조각), 범주 수준이 2개뿐일 때(ANOVA → t검정).
// 신호 세기도 일부러 섞었다 — 강함(claims) / 중간(wage) / 없음(policy) / 최소 구성(experience).
// 신호가 없는 표에서 R2_test가 0 근처·음수로 나오고 Lasso가 변수를 전부 버리는 것도
// '정상 동작'이다. 5단계가 '모델이 없다'고 말해 주는 것도 결과다.

import { readFileSync } from "node:fs";
import path from "node:path";

import { loadPyodide, type PyodideInterface } from "pyodide";
import { beforeAll, describe, expect, test } from "vitest";
import * as XLSX from "xlsx";

import { substitutePlaceholders } from "@/lib/grid/snippet-placeholders";
import { EXAMPLE_SNIPPET_FAMILIES, STAT_HEAD } from "@/lib/reference/exampleSnippets";
import { snippetInsertCode, type WrangleSnippet } from "@/lib/reference/wrangleSnippets";

interface Dataset {
  /** public/samples/<file>.xlsx */
  file: string;
  sheet?: string;
  target: string;
  cat: string[];
  num: string[];
  /** 이 표만의 특징 — 실패하면 무엇이 원인인지 바로 보이게 */
  note: string;
}

const DATASETS: Dataset[] = [
  {
    file: "wage",
    sheet: "wage",
    target: "WAGE",
    cat: ["SOUTH", "SEX", "UNION", "RACE", "OCCUPATION", "SECTOR", "MARR"],
    num: ["EDUCATION", "EXPERIENCE", "AGE"],
    note: "숫자 코드 범주형 · 결측 없음",
  },
  {
    file: "policy",
    target: "premium",
    cat: ["product", "channel", "region", "sex"],
    num: ["age", "bmi", "dependents", "tenure_months", "income"],
    note: "문자열 범주형 · ID 2개 · income 결측 75 · 선형 신호 없음",
  },
  {
    file: "claims",
    target: "prem_after",
    cat: ["product", "channel", "sex", "region", "age_band"],
    num: ["prem_before", "age", "claim_cnt"],
    note: "강한 신호(prem_before r=0.99) · 타깃이 10만 단위",
  },
  {
    file: "experience",
    target: "duration_years",
    cat: ["product", "sex"],
    num: ["entry_age"],
    note: "설명변수 최소 구성 — NUM 1개 · 범주 2수준",
  },
];

/** 통계분석 조각 27개 (위험률 산출은 표에 의존하지 않으므로 제외) */
const STAT_SNIPPETS: { group: string; s: WrangleSnippet }[] =
  EXAMPLE_SNIPPET_FAMILIES[0].groups.flatMap((g) =>
    g.snippets.map((s) => ({ group: g.label, s })),
  );

/** 조각의 머리말만 그 표의 것으로 갈아 끼운다 — 본문은 한 글자도 건드리지 않는다 */
function retarget(s: WrangleSnippet, d: Dataset): string {
  const head = [
    `df = xl("${d.file}!A1:A1", headers=True)`,
    `TARGET = ${JSON.stringify(d.target)}`,
    `CAT = ${JSON.stringify(d.cat)}`,
    `NUM = ${JSON.stringify(d.num)}`,
  ].join("\n");
  const text = substitutePlaceholders(snippetInsertCode(s), {
    "{{range}}": 'xl("SRC", headers=True)',
  });
  // STAT_HEAD 안의 {{range}}도 함께 치환된 상태이므로 같은 치환을 적용해 맞춘다
  const original = substitutePlaceholders(STAT_HEAD, {
    "{{range}}": 'xl("SRC", headers=True)',
  });
  expect(text, `${s.id}: 머리말을 찾지 못했다`).toContain(original);
  return text.replace(original, head);
}

const HARNESS = `
import json, io, ast, contextlib, traceback, warnings
warnings.filterwarnings("ignore")
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

_TABLES = {}

def xl(ref, headers=True):
    """조각이 넘긴 참조에서 표 이름만 떼어 해당 표를 돌려준다."""
    return _TABLES[ref.split("!")[0]].copy()

_G = {}

def _run(code):
    global _G
    g = {"pd": pd, "np": np, "xl": xl, "__name__": "__pygrid__"}
    _G = g                         # 실행 뒤 조각 안의 변수를 테스트에서 꺼내 볼 수 있게
    buf = io.StringIO()
    try:
        tree = ast.parse(code)
        tail = None
        if tree.body and isinstance(tree.body[-1], ast.Expr):
            tail = ast.Expression(tree.body[-1].value)
            tree.body = tree.body[:-1]
        with contextlib.redirect_stdout(buf):
            exec(compile(tree, "<snippet>", "exec"), g)
            value = eval(compile(tail, "<snippet>", "eval"), g) if tail is not None else None
    except Exception:
        return json.dumps({"ok": False, "err": traceback.format_exc()[-1500:],
                           "out": buf.getvalue()[-600:]})
    finally:
        plt.close("all")
    return json.dumps({"ok": True, "kind": type(value).__name__,
                       "empty": value is None, "out": buf.getvalue()[-600:]})
`;

let py: PyodideInterface;

beforeAll(async () => {
  py = await loadPyodide({ packageCacheDir: "node_modules/.pyodide-cache" });
  await py.loadPackage([
    "numpy",
    "pandas",
    "scipy",
    "statsmodels",
    "scikit-learn",
    "matplotlib",
  ]);
  py.runPython(HARNESS);

  for (const d of DATASETS) {
    const wb = XLSX.read(readFileSync(path.resolve(`public/samples/${d.file}.xlsx`)), {
      type: "buffer",
    });
    const csv = XLSX.utils.sheet_to_csv(wb.Sheets[d.sheet ?? wb.SheetNames[0]]);
    py.FS.writeFile(`/${d.file}.csv`, new TextEncoder().encode(csv));
    py.globals.set("_NAME", d.file);
    py.runPython(`_TABLES[_NAME] = pd.read_csv("/" + _NAME + ".csv")`);
  }
  // 표가 실제로 그 열들을 갖고 있는지 먼저 확인 — 아니면 머리말이 틀린 것
  for (const d of DATASETS) {
    py.globals.set("_NAME", d.file);
    py.globals.set("_COLS", JSON.stringify([d.target, ...d.cat, ...d.num]));
    const missing = py.runPython(
      `json.dumps([c for c in json.loads(_COLS) if c not in _TABLES[_NAME].columns])`,
    );
    expect(JSON.parse(missing as string), `${d.file}: 없는 열`).toEqual([]);
  }
}, 600_000);

// 오류 없이 도는 것과 '쓸 만한 모델이 나오는 것'은 다르다 —
// 마지막 단계(모델 비교표)가 표마다 5종을 모두 내고 R2가 유한한지 따로 확인한다.
test(
  "5단계 끝 — 표마다 회귀 5종 비교표가 나온다",
  () => {
    const compare = STAT_SNIPPETS.find(({ s }) => s.id === "eval-compare")!.s;
    const summary: string[] = [];
    for (const d of DATASETS) {
      py.globals.set("_CODE", retarget(compare, d));
      py.runPython(`_OUT = _run(_CODE)`);
      const r = JSON.parse(py.runPython("_OUT")) as { ok: boolean; err?: string; out: string };
      expect(r.err ?? "", `${d.file}: eval-compare 실패`).toBe("");

      // 비교표 본문을 다시 만들어 모델 이름과 R2를 꺼낸다
      const table = JSON.parse(
        py.runPython(`json.dumps(_G["out"][["model", "R2_test", "RMSE_test"]].to_dict("list"))`),
      ) as { model: string[]; R2_test: number[]; RMSE_test: number[] };
      expect(new Set(table.model), d.file).toEqual(
        new Set(["Ridge", "Lasso", "ElasticNet", "Polynomial(2)", "Log(Ridge)"]),
      );
      for (const v of table.R2_test) expect(Number.isFinite(v), `${d.file}: R2 비유한`).toBe(true);
      // 내림차순 정렬(비교표의 약속)
      expect([...table.R2_test].sort((a, b) => b - a)).toEqual(table.R2_test);
      summary.push(
        `${d.file.padEnd(11)} 최고 ${table.model[0]} R2=${table.R2_test[0].toFixed(4)} ` +
          `RMSE=${table.RMSE_test[0].toPrecision(4)}`,
      );
    }
    console.log("\n" + summary.join("\n"));
  },
  600_000,
);

for (const d of DATASETS) {
  describe(`${d.file} — ${d.note}`, () => {
    for (const { group, s } of STAT_SNIPPETS) {
      test(
        `${group} · ${s.label}`,
        () => {
          py.globals.set("_CODE", retarget(s, d));
          const r = JSON.parse(py.runPython("_run(_CODE)")) as {
            ok: boolean;
            kind?: string;
            empty?: boolean;
            out: string;
            err?: string;
          };
          expect(r.err ?? "", `${d.file} / ${s.id}\n${r.out}`).toBe("");
          expect(r.empty, `${d.file} / ${s.id}: 결과 값이 없다`).toBe(false);
        },
        180_000,
      );
    }
  });
}
