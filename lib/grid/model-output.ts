// 부록 O.5: "모델 결과 → 시트" 선택을 출력 바인딩 명세로 바꾼다 (순수 함수).
// 각 명세의 selection.variable은 런타임 _pygrid_resolve가 평가하는 출력 식이고,
// code는 같은 결과를 내는 라이브러리 원래 코드(사용자가 자기 코드로 옮겨 쓸 때)다.

import type { ModelInfo, VariableInfo } from "@/lib/runtime/protocol";
import type { OutputSelection } from "@/types/workbook";

export type Interval = "none" | "confidence" | "prediction";

export interface ModelOutputChoice {
  /** 모델 변수 이름 */
  model: string;
  /** 고른 카탈로그 항목의 식 (ModelMember.expr) */
  members: string[];
  /** 예측·평가에 쓸 데이터셋(DataFrame) 변수 이름 */
  data?: string;
  /** 예측값 내보내기 (data 필요) */
  predict?: { interval: Interval; level: number; keep: boolean };
  /** 클래스 확률 predict_proba (data 필요, 분류 모델) */
  proba?: boolean;
  /** 평가 지표(R²·RMSE·MAE)의 목표 열 (data 필요) */
  target?: string;
}

export interface OutputSpec {
  selection: OutputSelection;
  label: string;
  /** 라이브러리 원래 코드 */
  code: string;
  /** 예상 spill 너비 (가로 배치 간격) */
  width: number;
}

/** JSON 문자열 리터럴은 그대로 유효한 Python 문자열 리터럴이다 */
const pyStr = (s: string): string => JSON.stringify(s);

/** 예측 입력 표현 — sklearn은 학습 때 열 순서로 고르고, statsmodels 공식 모델은 표 그대로 */
const designCode = (model: string, data: string, info: ModelInfo): string =>
  info.kind === "sklearn" && info.featureNames ? `${data}[${model}.feature_names_in_]` : data;

export function buildModelOutputSpecs(
  choice: ModelOutputChoice,
  info: ModelInfo,
  dataVar?: VariableInfo,
): OutputSpec[] {
  const { model, data } = choice;
  const specs: OutputSpec[] = [];
  for (const m of info.members) {
    if (!choice.members.includes(m.expr)) continue;
    specs.push({ selection: { variable: m.expr }, label: `${model}: ${m.label}`, code: m.code, width: m.shape[1] });
  }
  if (data && choice.predict && info.predict) {
    const { interval, level, keep } = choice.predict;
    const args = [model, data];
    if (interval !== "none") args.push(`interval=${pyStr(interval)}`, `level=${level}`);
    if (keep) args.push("keep=True");
    const tag = interval === "prediction" ? " + 예측구간" : interval === "confidence" ? " + 신뢰구간" : "";
    const X = designCode(model, data, info);
    specs.push({
      selection: { variable: `_pygrid_predict(${args.join(", ")})` },
      label: `${model}: 예측 ← ${data}${tag}`,
      code:
        interval === "none"
          ? `${model}.predict(${X})`
          : `${model}.get_prediction(${X}).summary_frame(alpha=${Number((1 - level).toFixed(4))})`,
      width: 2 + (interval !== "none" ? 2 : 0) + (keep ? (dataVar?.columns?.length ?? 0) : 0),
    });
  }
  if (data && choice.proba && info.proba) {
    specs.push({
      selection: { variable: `_pygrid_predict(${model}, ${data}, proba=True)` },
      label: `${model}: 클래스 확률 ← ${data}`,
      code: `${model}.predict_proba(${designCode(model, data, info)})`,
      width: 1 + (info.members.find((m) => m.code.endsWith(".classes_"))?.shape[0] ?? 2),
    });
  }
  if (data && choice.target && info.predict) {
    specs.push({
      selection: { variable: `_pygrid_eval_metrics(${model}, ${data}, ${pyStr(choice.target)})` },
      label: `${model}: 평가 지표 (${data} · ${choice.target})`,
      code: `_pygrid_eval_metrics(${model}, ${data}, ${pyStr(choice.target)})`,
      width: 2,
    });
  }
  return specs;
}

/** 코드에 필요한 import를 앞에 붙인다 (새 블록으로 보낼 때) */
export function withImports(code: string): string {
  const lines: string[] = [];
  if (/\bnp\./.test(code)) lines.push("import numpy as np");
  if (/\bpd\./.test(code)) lines.push("import pandas as pd");
  return [...lines, code].join("\n");
}
