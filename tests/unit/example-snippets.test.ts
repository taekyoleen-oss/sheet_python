// 부록 N — 예제 코드 라이브러리의 구조·규칙 검증(실행은 tests/pyodide/example-snippets.test.ts).
//
// 지키는 규칙:
//  · 상위 카테고리는 통계분석 / 위험률 산출 두 갈래이고 하위 카테고리 구성이 고정되어 있다.
//  · 모든 조각의 id는 전역 유일 — 코드 삽입 팝업이 id로 선택 상태를 들고 다닌다.
//  · 마지막 줄은 print가 아니라 값(식) — 그래야 셀에 결과가 나온다(wrangleSnippets와 같은 규칙).
//  · 모델 범위는 Ridge·Lasso·ElasticNet + polynomial·log로 한정 (트리 계열·베이지안 최적화 제외).

import { describe, expect, test } from "vitest";

import { substitutePlaceholders } from "@/lib/grid/snippet-placeholders";
import { EXAMPLE_SNIPPET_FAMILIES, findExampleSnippet } from "@/lib/reference/exampleSnippets";
import { snippetInsertCode } from "@/lib/reference/wrangleSnippets";

const ALL = EXAMPLE_SNIPPET_FAMILIES.flatMap((f) =>
  f.groups.flatMap((g) => g.snippets.map((s) => ({ family: f.id, group: g.id, ...s }))),
);

test("상위 카테고리 2종 — 통계분석 · 위험률 산출", () => {
  expect(EXAMPLE_SNIPPET_FAMILIES.map((f) => f.label)).toEqual(["통계분석", "위험률 산출"]);
});

test("하위 카테고리 구성 — 사용자가 지정한 5개 + 계리 절차 4개", () => {
  const [stat, risk] = EXAMPLE_SNIPPET_FAMILIES;
  expect(stat.groups.map((g) => g.label)).toEqual([
    "통계분석",
    "전처리 과정",
    "특성공학",
    "데이터 분석 (회귀 모델)",
    "모델 평가",
  ]);
  expect(risk.groups.map((g) => g.label)).toEqual([
    "위험률 산출 (조율·평활)",
    "생명표·계산기수",
    "다중탈퇴·다급부",
    "보험료·준비금",
  ]);
  for (const g of [...stat.groups, ...risk.groups]) {
    expect(g.snippets.length, g.label).toBeGreaterThanOrEqual(3);
  }
});

test("id는 전역 유일하고 findExampleSnippet으로 찾힌다", () => {
  const ids = ALL.map((s) => s.id);
  expect(new Set(ids).size).toBe(ids.length);
  for (const id of ids) expect(findExampleSnippet(id)?.id).toBe(id);
  expect(findExampleSnippet("없는-id")).toBeUndefined();
});

test("모델 범위 — 트리 계열·베이지안 최적화·외부 패키지 없음", () => {
  // Pyodide에 없거나(lightgbm·xgboost·shap) 이번 범위에서 뺀(트리·BO) 것들
  const banned =
    /lightgbm|LGBM|xgboost|RandomForest|GradientBoosting|DecisionTree|bayes_opt|BayesianOptimization|import shap|optbinning|lifelines/i;
  for (const s of ALL) {
    expect(s.code, `${s.id}: 범위 밖 모델/패키지`).not.toMatch(banned);
  }
  // 회귀 모델 그룹은 지정된 5종을 모두 다룬다
  const model = EXAMPLE_SNIPPET_FAMILIES[0].groups.find((g) => g.id === "model")!;
  const code = model.snippets.map((s) => s.code).join("\n");
  for (const kw of ["RidgeCV", "LassoCV", "ElasticNetCV", "PolynomialFeatures", "np.log("]) {
    expect(code, `회귀 모델 그룹에 ${kw} 없음`).toContain(kw);
  }
});

describe.each(ALL.map((s) => [s.id, s] as const))("%s", (_id, s) => {
  test("라벨·설명·코드가 채워져 있다", () => {
    expect(s.label.trim().length).toBeGreaterThan(0);
    expect(s.desc.trim().length).toBeGreaterThan(0);
    expect(s.code.trim().length).toBeGreaterThan(0);
  });

  test("마지막 문장은 print가 아니라 값(식)", () => {
    // 마지막 문장의 첫 줄 = 뒤에서부터 찾은 '들여쓰기 없는' 줄 (여러 줄 식이 흔하다).
    // 실제로 '식'인지는 실행으로 확정된다 — tests/pyodide/example-snippets.test.ts의 empty 검사.
    const lines = s.code.trimEnd().split("\n");
    const head = [...lines].reverse().find((L) => L.trim() && L === L.trimStart());
    expect(head, "들여쓰기 없는 줄이 없다").toBeDefined();
    expect(head!.startsWith("print("), `${s.id}: print로 끝나면 셀에 결과가 없다`).toBe(false);
    expect(head!.startsWith("#")).toBe(false);
  });

  test("삽입 텍스트에 설명 헤더가 붙는다", () => {
    const text = snippetInsertCode(s);
    expect(text.startsWith(`# ▸ ${s.label}\n# ${s.desc}\n`)).toBe(true);
  });

  test("자리표시자 {{range}}는 xl() 호출 전체로 치환된다", () => {
    // 앱은 {{range}}를 `xl("A1:K535", headers=True)` 호출 '전체'로 바꾼다(xl-ref.ts).
    // 따옴표 안에 쓰면 xl("xl(...)")가 되어 깨지므로 금지한다.
    expect(s.code, `${s.id}: {{range}}를 따옴표 안에 쓰면 안 된다`).not.toContain('"{{range}}"');
    const subbed = substitutePlaceholders(s.code, { "{{range}}": 'xl("wage!A1:K535", headers=True)' });
    expect(subbed).not.toContain("{{range}}");
  });
});
