// 부록 N — 임금 회귀 예측 워크북을 Node용 Pyodide에서 1단계부터 16단계까지 실제로 돌린다.
//
// 브라우저 실행과 같은 조건을 만든다: 워크북 JSON의 시트 셀을 그대로 표로 복원하고,
// xl("wage!A1:K535", headers=True)가 그 표를 돌려주게 한 뒤 블록을 계산 순서대로 실행한다.
// 블록마다 전역을 새로 만들어(앱보다 엄격) 각 단계가 혼자서도 도는지 함께 검증한다.
//
// 첫 실행은 scikit-learn·statsmodels·matplotlib CDN 다운로드로 느리다(캐시됨).

import { readFileSync } from "node:fs";
import path from "node:path";

import { loadPyodide, type PyodideInterface } from "pyodide";
import { beforeAll, expect, test } from "vitest";

import { blocksInOrder } from "@/lib/grid/model";
import { parseWorkbookJson } from "@/lib/io/workbook-json";
import type { PyBlock } from "@/types/workbook";

const wb = parseWorkbookJson(
  readFileSync(path.resolve("data/sample-workbooks/wage-regression.pygrid.json"), "utf8"),
);

/** 시트 셀 → 조밀한 2차원 값 배열 (브라우저의 xl() 스냅샷과 같은 모양) */
function sheetGrid(name: string): (string | number | boolean | null)[][] {
  const sheet = wb.sheets.find((s) => s.name === name);
  if (!sheet) throw new Error(`시트 없음: ${name}`);
  let maxR = -1;
  let maxC = -1;
  for (const key of Object.keys(sheet.cells)) {
    const [r, c] = key.split(":").map(Number);
    maxR = Math.max(maxR, r);
    maxC = Math.max(maxC, c);
  }
  const grid: (string | number | boolean | null)[][] = [];
  for (let r = 0; r <= maxR; r++) {
    const row: (string | number | boolean | null)[] = [];
    for (let c = 0; c <= maxC; c++) row.push(sheet.cells[`${r}:${c}`]?.v ?? null);
    grid.push(row);
  }
  return grid;
}

/** 워커의 xl()과 같은 규칙: 참조 범위를 잘라 DataFrame(headers=True면 첫 행이 열 이름) */
const HARNESS = `
import json, io, re, ast, contextlib, traceback, warnings
warnings.filterwarnings("ignore")
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

_SHEETS = json.loads(_SHEETS_JSON)

def _col(s):
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n - 1

def xl(ref, headers=False):
    m = re.fullmatch(r"(?:([^!]+)!)?([A-Z]+)(\\d+)(?::([A-Z]+)(\\d+))?", ref)
    if not m:
        raise ValueError("참조 형식 오류: " + ref)
    name, c0, r0, c1, r1 = m.group(1), _col(m.group(2)), int(m.group(3)) - 1, m.group(4), m.group(5)
    c1 = _col(c1) if c1 else c0
    r1 = int(r1) - 1 if r1 else r0
    grid = _SHEETS[name]
    block = [[(row[c] if c < len(row) else None) for c in range(c0, c1 + 1)]
             for row in grid[r0:r1 + 1]]
    if headers:
        return pd.DataFrame(block[1:], columns=[str(h) for h in block[0]]).infer_objects()
    if r0 == r1 and c0 == c1:
        return block[0][0]
    return pd.DataFrame(block).infer_objects()

def _run_block(code):
    g = {"pd": pd, "np": np, "xl": xl, "__name__": "__pygrid__"}
    buf = io.StringIO()
    try:
        tree = ast.parse(code)
        tail = None
        if tree.body and isinstance(tree.body[-1], ast.Expr):
            tail = ast.Expression(tree.body[-1].value)
            tree.body = tree.body[:-1]
        with contextlib.redirect_stdout(buf):
            exec(compile(tree, "<block>", "exec"), g)
            value = eval(compile(tail, "<block>", "eval"), g) if tail is not None else None
    except Exception:
        return json.dumps({"ok": False, "err": traceback.format_exc()[-1800:],
                           "out": buf.getvalue()[-1200:]})
    finally:
        plt.close("all")
    kind = type(value).__name__
    shape = None
    if isinstance(value, pd.DataFrame):
        shape = [int(value.shape[0]), int(value.shape[1])]
    elif isinstance(value, pd.Series):
        shape = [int(value.shape[0]), 1]
    return json.dumps({"ok": True, "kind": kind, "shape": shape,
                       "out": buf.getvalue()[-2500:]})
`;

interface BlockRun {
  ok: boolean;
  kind?: string;
  shape?: [number, number] | null;
  out: string;
  err?: string;
}

let py: PyodideInterface;
const results = new Map<string, BlockRun>();

const codeBlocks = (): PyBlock[] =>
  blocksInOrder(wb).filter((b) => b.kind !== "markdown");

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
  py.globals.set(
    "_SHEETS_JSON",
    JSON.stringify({ wage: sheetGrid("wage"), meta: sheetGrid("meta") }),
  );
  py.runPython(HARNESS);
  expect(py.runPython('len(xl("wage!A1:K535", headers=True))')).toBe(534);
}, 600_000);

test("워크북 구조 — 16단계 [설명·코드] 쌍", () => {
  const blocks = blocksInOrder(wb);
  expect(blocks[0].kind).toBe("markdown");
  expect(codeBlocks()).toHaveLength(16);
  // 계산 순서 = 앵커 행 오름차순 (셀 의존이 없는 블록은 위에서 아래로)
  const rows = codeBlocks().map((b) => b.anchor.r);
  expect([...rows].sort((a, b) => a - b)).toEqual(rows);
});

// 단계별 실행 — 앞 단계가 실패해도 뒤 단계를 마저 돌려 전체 상태를 한 번에 본다
for (const [i, block] of codeBlocks().entries()) {
  test(
    `${i + 1}단계 — ${block.title}`,
    () => {
      py.globals.set("_CODE", block.code);
      const r = JSON.parse(py.runPython("_run_block(_CODE)")) as BlockRun;
      results.set(block.id, r);
      expect(r.err ?? "", `${block.id}\n${r.out}`).toBe("");
      expect(r.kind, `${block.id}: 마지막 줄이 값(식)이 아니다`).not.toBe("NoneType");
      // spill이 시트 밖으로 나가면 안 된다 (앵커 열 13 + 표 너비)
      const sheet = wb.sheets.find((s) => s.id === block.sheetId)!;
      if (r.shape) {
        expect(block.anchor.c + r.shape[1], `${block.id}: spill 열 초과`)
          .toBeLessThanOrEqual(sheet.colCount);
        expect(block.anchor.r + r.shape[0], `${block.id}: spill 행 초과`)
          .toBeLessThanOrEqual(sheet.rowCount);
      }
    },
    240_000,
  );
}

test("산출 결과 검증 — 데이터·공선성·모델 성능", () => {
  const out = (id: string): string => {
    const r = results.get(id);
    if (!r) throw new Error(`블록 미실행: ${id}`);
    return r.out;
  };

  // 1단계 — 원본 노트북과 같은 534행 11열, 결측·중복 없음
  expect(out("blk-wr-c1")).toContain("shape: (534, 11)");
  expect(out("blk-wr-c1")).toContain("중복 행: 0  결측 합계: 0");

  // 3단계 — AGE = EDUCATION + EXPERIENCE + {2, 6} 이라는 정의상 공선성
  expect(out("blk-wr-c3")).toContain("[2, 6]");
  expect(out("blk-wr-c3")).toMatch(/'AGE'.*'EXPERIENCE'|'EXPERIENCE'.*'AGE'/);

  // 7단계 — 원-핫으로 10열 → 16열, 7:3 분할
  expect(out("blk-wr-c7")).toContain("설계행렬 16열 (원본 10열)  train 373행 / test 161행");

  // 12단계 — Lasso·ElasticNet만 계수를 0으로 만든다 (Ridge·OLS는 전부 유지)
  const m12 = out("blk-wr-c12");
  expect(m12).toMatch(/OLS\s+R2 test\s+[\d.]+\s+0이 아닌 계수 16\/16/);
  expect(m12).toMatch(/Ridge\s+R2 test\s+[\d.]+\s+0이 아닌 계수 16\/16/);
  expect(m12).toMatch(/Lasso\s+R2 test\s+[\d.]+\s+0이 아닌 계수 1[0-5]\/16/);

  // 13단계 — 스미어링은 'RMSE 개선'이 아니라 '편의 제거'다.
  //   train에서는 편의가 -0.86 → -0.07로 거의 사라지고 RMSE도 좋아진다(이론대로).
  //   그런데 이 7:3 분할은 test 평균이 train보다 낮아 test RMSE가 오히려 나빠진다 —
  //   워크북이 이 사실을 그대로 싣고 있는지 확인한다(분할 한 번으로 판단하지 말 것).
  const m13 = out("blk-wr-c13");
  const tr = /train 편의 ([-+][\d.]+) -> ([-+][\d.]+)\s+RMSE ([\d.]+) -> ([\d.]+)/.exec(m13);
  const te = /test 편의 ([-+][\d.]+) -> ([-+][\d.]+)\s+RMSE ([\d.]+) -> ([\d.]+)/.exec(m13);
  expect(tr, m13).not.toBeNull();
  expect(te, m13).not.toBeNull();
  expect(Math.abs(Number(tr![2]))).toBeLessThan(Math.abs(Number(tr![1]))); // 편의 축소
  expect(Number(tr![4])).toBeLessThan(Number(tr![3])); // train RMSE 개선
  expect(Number(te![4])).toBeGreaterThan(Number(te![3])); // test RMSE는 악화 — 문서화된 결과
  expect(m13).toContain("14단계 교차검증으로 확인한다");

  // 16단계 — 95% 예측구간 적중률이 합리적 범위
  const hit = /95% 예측구간 적중률 (\d+\.\d)%/.exec(out("blk-wr-c16"));
  expect(hit, out("blk-wr-c16")).not.toBeNull();
  expect(Number(hit![1])).toBeGreaterThan(85);
  expect(Number(hit![1])).toBeLessThanOrEqual(100);
});
