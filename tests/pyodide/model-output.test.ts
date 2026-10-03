// 부록 O.5 — 모델 결과 → 시트: 실제 statsmodels·scikit-learn 모델을 적합하고
// 카탈로그(inspect) → 출력 식(buildModelOutputSpecs) → 다중 출력 변환까지 앱과 같은 경로로 돌린다.

import { readFileSync } from "node:fs";
import path from "node:path";

import { loadPyodide, type PyodideInterface } from "pyodide";
import { beforeAll, describe, expect, test } from "vitest";

import { buildModelOutputSpecs } from "@/lib/grid/model-output";
import type { VariableInfo } from "@/lib/runtime/protocol";

let py: PyodideInterface;

type Cell = { v: unknown; t: string };
type Item = { id: string; ok: boolean; cells?: Cell[][]; msg?: string; etype?: string };

const FIT = `
import numpy as np, pandas as pd
import statsmodels.formula.api as smf
from sklearn.linear_model import LinearRegression, Ridge
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import StandardScaler
rng = np.random.default_rng(0)
df = pd.DataFrame({"x1": rng.normal(size=60), "x2": rng.normal(size=60)})
df["y"] = 1.5 + 2.0 * df.x1 - 0.5 * df.x2 + rng.normal(scale=0.3, size=60)
test = df.head(5)
ols = smf.ols("y ~ x1 + x2", df).fit()
lr = LinearRegression().fit(df[["x1", "x2"]], df.y)
pipe = make_pipeline(StandardScaler(), Ridge(alpha=1.0)).fit(df[["x1", "x2"]], df.y)
unfitted = LinearRegression()
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import RandomForestClassifier
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
from sklearn.model_selection import GridSearchCV
df["hi"] = (df.y > df.y.median()).astype(int)
logit = smf.logit("hi ~ x1 + x2", df).fit(disp=0)
clf = LogisticRegression().fit(df[["x1", "x2"]], df.hi)
rf = RandomForestClassifier(n_estimators=10, random_state=0).fit(df[["x1", "x2"]], df.hi)
km = KMeans(n_clusters=2, n_init=3, random_state=0).fit(df[["x1", "x2"]])
pca = PCA(n_components=2).fit(df[["x1", "x2"]])
gs = GridSearchCV(Ridge(), {"alpha": [0.1, 1.0]}, cv=3).fit(df[["x1", "x2"]], df.y)
`;

const inspect = (): VariableInfo[] => JSON.parse(py.runPython("_pygrid_inspect()") as string);

/** 블록 본문은 이미 실행됨 — 출력 식만 평가 (본문 "None") */
function run(selections: string[]): Item[] {
  py.globals.set(
    "_pygrid_outputs",
    JSON.stringify(
      selections.map((variable, i) => ({
        id: `o${i}`,
        mode: "values",
        includeIndex: "auto",
        selection: { variable },
      })),
    ),
  );
  const res = JSON.parse(py.runPython('_pygrid_run_convert_multi("None", _pygrid_outputs)') as string);
  expect(res.ok).toBe(true);
  return res.items;
}

/** 성공 출력의 셀 (실패면 메시지와 함께 실패) */
function okCells(item: Item): Cell[][] {
  expect(item.ok, `${item.etype}: ${item.msg}`).toBe(true);
  return item.cells!;
}

/** 첫 열 라벨 → 둘째 열 값 */
const asMap = (cells: Cell[][]) => Object.fromEntries(cells.slice(1).map((r) => [r[0].v, r[1].v]));

beforeAll(async () => {
  py = await loadPyodide({ packageCacheDir: "node_modules/.pyodide-cache" });
  await py.loadPackage(["numpy", "pandas", "scipy", "statsmodels", "scikit-learn"]);
  for (const f of ["bootstrap.py", "xl.py", "convert.py", "model_out.py"]) {
    py.runPython(readFileSync(path.resolve("lib/runtime/py", f), "utf8"));
  }
  py.runPython(FIT);
}, 600_000);

describe("카탈로그 (inspect)", () => {
  test("적합된 모델만 model 정보, 표는 columns", () => {
    const vars = inspect();
    const by = (n: string) => vars.find((v) => v.name === n)!;
    expect(by("ols").model?.kind).toBe("statsmodels");
    expect(by("ols").model?.formula).toBe(true);
    expect(by("ols").model?.intervals).toBe(true);
    expect(by("lr").model?.kind).toBe("sklearn");
    expect(by("pipe").model?.kind).toBe("sklearn");
    expect(by("unfitted").model).toBeUndefined(); // 적합 전 추정기는 제외
    expect(by("df").columns).toEqual(["x1", "x2", "y", "hi"]);
    const labels = by("ols").model!.members.map((m) => m.expr);
    expect(labels).toContain("ols.rsquared");
    expect(labels).toContain("_pygrid_coef_table(ols)");
  });
});

describe("statsmodels OLS → 시트", () => {
  test("계수표·적합 통계·R² 단일 셀이 값으로 펼쳐진다", () => {
    const [coef, stats, r2] = run(["_pygrid_coef_table(ols)", "_pygrid_fit_stats(ols)", "ols.rsquared"]);
    expect(coef.ok).toBe(true);
    expect(okCells(coef)[0].map((c) => c.v)).toEqual(["변수", "계수", "표준오차", "t(z)", "p값", "하한 95%", "상한 95%"]);
    expect(okCells(coef).slice(1).map((r) => r[0].v)).toEqual(["Intercept", "x1", "x2"]);
    expect(okCells(coef)[2][1].v as number).toBeCloseTo(2.0, 0);

    const rsq = py.runPython("float(ols.rsquared)") as number;
    expect(asMap(okCells(stats))["결정계수 R²"] as number).toBeCloseTo(rsq, 12);
    expect(r2.cells).toEqual([[{ v: rsq, t: "n" }]]);
  });

  test("R식 predict(newdata, interval='prediction') + 평가 지표", () => {
    const vars = inspect();
    const ols = vars.find((v) => v.name === "ols")!;
    const specs = buildModelOutputSpecs(
      {
        model: "ols",
        members: [],
        data: "df",
        predict: { interval: "prediction", level: 0.95, keep: false },
        target: "y",
      },
      ols.model!,
      vars.find((v) => v.name === "df"),
    );
    const [pred, metrics] = run(specs.map((s) => s.selection.variable!));
    expect(pred.ok).toBe(true);
    expect(okCells(pred).length).toBe(61); // 헤더 + 60행
    expect(okCells(pred)[0].map((c) => c.v).slice(-3)).toEqual(["예측", "예측구간 하한 95%", "예측구간 상한 95%"]);
    const [p, lo, hi] = okCells(pred)[1].slice(-3).map((c) => c.v as number);
    expect(lo).toBeLessThan(p);
    expect(p).toBeLessThan(hi);
    // 학습 데이터 R²는 statsmodels rsquared와 같다
    expect(asMap(okCells(metrics))["결정계수 R²"] as number).toBeCloseTo(
      py.runPython("float(ols.rsquared)") as number,
      10,
    );
  });
});

describe("scikit-learn → 시트", () => {
  test("LinearRegression 계수표(절편 포함)·예측(원본 열과 함께)", () => {
    const [coef, pred] = run(["_pygrid_coef_table(lr)", "_pygrid_predict(lr, test, keep=True)"]);
    expect(okCells(coef).slice(1).map((r) => r[0].v)).toEqual(["(절편)", "x1", "x2"]);
    expect(okCells(coef)[1][1].v as number).toBeCloseTo(py.runPython("float(lr.intercept_)") as number, 12);
    expect(okCells(pred)[0].map((c) => c.v).slice(-4)).toEqual(["x1", "x2", "y", "예측"]);
    expect(okCells(pred).length).toBe(6);
  });

  test("Pipeline(StandardScaler→Ridge)도 계수표·예측", () => {
    const [coef, pred] = run(["_pygrid_coef_table(pipe)", "_pygrid_predict(pipe, test)"]);
    expect(coef.ok).toBe(true);
    expect(okCells(coef).slice(2).map((r) => r[0].v)).toEqual(["x1", "x2"]);
    expect(pred.ok).toBe(true);
  });

  test("sklearn 적합 통계·구간 예측은 안내 오류 (그 출력만 실패)", () => {
    const [stats, interval, ok] = run([
      "_pygrid_fit_stats(lr)",
      '_pygrid_predict(lr, test, interval="prediction")',
      "lr.intercept_",
    ]);
    expect(stats.ok).toBe(false);
    expect(interval.ok).toBe(false);
    expect(ok.ok).toBe(true);
  });

  test("학습 때 열이 없으면 KeyError로 알려 준다", () => {
    const [bad] = run(['_pygrid_predict(lr, test[["x1"]])']);
    expect(bad.ok).toBe(false);
    expect(bad.etype).toBe("KeyError");
  });
});

describe("모델별로 다른 속성 카탈로그 (코드·미리보기 포함)", () => {
  const info = (name: string) => inspect().find((v) => v.name === name)!.model!;
  const labels = (name: string) => info(name).members.map((m) => m.label);

  test("모델 종류마다 항목이 다르다", () => {
    expect(labels("ols")).toContain("결정계수 R²");
    expect(labels("ols")).not.toContain("오즈비 exp(β)");
    expect(labels("logit")).toContain("오즈비 exp(β)");
    expect(labels("logit")).toContain("McFadden 의사 R²");
    expect(labels("clf")).toContain("클래스");
    expect(info("clf").proba).toBe(true);
    expect(info("lr").proba).toBe(false);
    expect(labels("rf")).toContain("특성 중요도");
    expect(labels("km")).toEqual(expect.arrayContaining(["군집 중심", "관성 (군집 내 제곱합)"]));
    expect(info("pca").predict).toBe(false);
    expect(labels("pca")).toContain("설명 분산 비율");
    expect(labels("gs")).toEqual(expect.arrayContaining(["최적 파라미터", "교차검증 결과표"]));
  });

  test("미리보기 값과 코드가 채워지고, 보여 주는 코드는 그대로 실행된다", () => {
    for (const name of ["ols", "logit", "lr", "pipe", "clf", "rf", "km", "pca", "gs"]) {
      for (const m of info(name).members) {
        expect(m.preview, `${name} ${m.label}`).not.toBe("");
        const ok = py.runPython(
          `import numpy as np, pandas as pd
try:
    eval(${JSON.stringify(m.code)})
    _r = True
except Exception as _e:
    _r = repr(_e)
_r`,
        );
        expect(ok, `${name}: ${m.code}`).toBe(true);
      }
    }
    const r2 = info("ols").members.find((m) => m.expr === "ols.rsquared")!;
    expect(Number(r2.preview)).toBeCloseTo(py.runPython("float(ols.rsquared)") as number, 5);
  });

  test("분류 모델 클래스 확률·특성 중요도가 셀 값으로 펼쳐진다", () => {
    const [proba, imp, cv] = run([
      "_pygrid_predict(clf, test, proba=True)",
      "_pygrid_importances(rf)",
      "_pygrid_cv_table(gs)",
    ]);
    expect(okCells(proba)[0].map((c) => c.v)).toEqual(["P(0)", "P(1)"]);
    const row = okCells(proba)[1].map((c) => c.v as number);
    expect(row[0] + row[1]).toBeCloseTo(1, 10);
    expect(okCells(imp)[0].map((c) => c.v)).toEqual(["특성", "중요도"]);
    expect(okCells(cv).length).toBe(3);
  });
});
