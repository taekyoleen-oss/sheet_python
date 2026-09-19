// 예제 코드 라이브러리(부록 N) — 상위 카테고리 2종의 모든 조각을 Node용 Pyodide에서
// 실제로 실행한다. 앱의 코드 삽입 팝업이 넣는 텍스트(snippetInsertCode) 그대로 돌리고,
// 마지막 식의 값이 실제로 생기는지(= 셀에 나올 결과가 있는지)까지 확인한다.
//
// 통계분석 조각의 xl()은 public/samples/wage.xlsx(임금 534행)를 돌려주도록 스텁한다 —
// 브라우저에서 워크북의 wage 시트를 읽는 것과 같은 데이터다.
// 첫 실행은 scikit-learn·statsmodels·matplotlib을 CDN에서 받아 느리다(캐시됨).

import { readFileSync } from "node:fs";
import path from "node:path";

import { loadPyodide, type PyodideInterface } from "pyodide";
import { beforeAll, describe, expect, test } from "vitest";
import * as XLSX from "xlsx";

import { substitutePlaceholders } from "@/lib/grid/snippet-placeholders";
import { EXAMPLE_SNIPPET_FAMILIES } from "@/lib/reference/exampleSnippets";
import { snippetInsertCode, type WrangleSnippet } from "@/lib/reference/wrangleSnippets";

/** 앱이 삽입하는 텍스트 그대로 — 설명 헤더 + 그리드 선택 범위로 치환한 {{range}} */
const insertText = (s: WrangleSnippet): string =>
  substitutePlaceholders(snippetInsertCode(s), {
    "{{range}}": 'xl("wage!A1:K535", headers=True)',
  });

let py: PyodideInterface;

/** 워크북 초기화 스크립트와 같은 전역 + xl() 스텁 + 조각 실행기 */
const HARNESS = `
import json, io, traceback, warnings
warnings.filterwarnings("ignore")
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

_WAGE = pd.read_csv("/wage.csv")

def xl(ref, headers=True):
    """예제 조각용 스텁 — 어떤 참조든 임금 표를 돌려준다."""
    return _WAGE.copy()

def _run_snippet(code):
    g = {"pd": pd, "np": np, "xl": xl, "__name__": "__pygrid__"}
    buf = io.StringIO()
    import contextlib, ast
    try:
        tree = ast.parse(code)
        last = tree.body[-1] if tree.body else None
        tail = None
        if isinstance(last, ast.Expr):
            tail = ast.Expression(last.value)
            tree.body = tree.body[:-1]
        with contextlib.redirect_stdout(buf):
            exec(compile(tree, "<snippet>", "exec"), g)
            value = eval(compile(tail, "<snippet>", "eval"), g) if tail is not None else None
    except Exception:
        return json.dumps({"ok": False, "err": traceback.format_exc()[-1600:],
                           "out": buf.getvalue()[-800:]})
    finally:
        plt.close("all")
    return json.dumps({
        "ok": True,
        "kind": type(value).__name__,
        "empty": value is None,
        "out": buf.getvalue()[-800:],
    })
`;

interface SnippetRun {
  ok: boolean;
  kind?: string;
  empty?: boolean;
  out: string;
  err?: string;
}

/** _CODE 전역에 넣어 둔 코드를 실행 (긴 문자열을 Python 소스에 끼워 넣지 않는다) */
const runStaged = (): SnippetRun =>
  JSON.parse(py.runPython("_run_snippet(_CODE)")) as SnippetRun;

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

  // public/samples/wage.xlsx → CSV → Pyodide FS (브라우저에서 시트를 읽는 것과 같은 값)
  const wb = XLSX.read(readFileSync(path.resolve("public/samples/wage.xlsx")), {
    type: "buffer",
  });
  const csv = XLSX.utils.sheet_to_csv(wb.Sheets["wage"]);
  py.FS.writeFile("/wage.csv", new TextEncoder().encode(csv));
  py.runPython(HARNESS);
  expect(py.runPython("len(_WAGE)")).toBe(534);
}, 600_000);

for (const fam of EXAMPLE_SNIPPET_FAMILIES) {
  describe(`${fam.label}`, () => {
    for (const group of fam.groups) {
      for (const s of group.snippets) {
        test(`${group.label} · ${s.label}`, () => {
          py.globals.set("_CODE", insertText(s));
          const r = runStaged();
          expect(r.err ?? "", `${s.id}\n${r.out}`).toBe("");
          expect(r.ok, s.id).toBe(true);
          // 마지막 줄이 값(식)이어야 셀에 결과가 나온다 — print로 끝나면 빈 결과
          expect(r.empty, `${s.id}: 마지막 줄이 값(식)이 아니다`).toBe(false);
        }, 180_000);
      }
    }
  });
}
