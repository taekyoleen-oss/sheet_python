# Python 결과 → 시트 (부록 O.5) — 모델 객체의 계수·적합 통계·예측을 "출력 식"으로 고른다.
# 출력 선택(OutputSelection.variable)은 변수명 또는 이 헬퍼를 부르는 식이고,
# 블록 실행 직후 _pygrid_resolve가 평가해 기존 spill 경로(§3.3 변환)로 셀에 놓인다.
# bootstrap.py 규칙을 따른다: _pygrid_ 접두사, import는 함수 안에서만.
# statsmodels·scikit-learn은 import하지 않고 덕 타이핑으로 판별한다(사용자 코드가 이미 불러온 경우만 해당).

_PYGRID_SM_STATS = [
    ("nobs", "관측치 수"),
    ("df_model", "모형 자유도"),
    ("df_resid", "잔차 자유도"),
    ("rsquared", "결정계수 R²"),
    ("rsquared_adj", "수정 R²"),
    ("prsquared", "McFadden 의사 R²"),
    ("fvalue", "F 통계량"),
    ("f_pvalue", "F 검정 p값"),
    ("llr", "우도비 통계량"),
    ("llr_pvalue", "우도비 검정 p값"),
    ("aic", "AIC"),
    ("bic", "BIC"),
    ("llf", "로그우도"),
    ("llnull", "영모형 로그우도"),
    ("deviance", "이탈도"),
    ("null_deviance", "영모형 이탈도"),
    ("pearson_chi2", "Pearson χ²"),
    ("scale", "척도(σ²)"),
    ("mse_resid", "잔차 MSE"),
    ("ssr", "잔차제곱합 SSR"),
    ("ess", "회귀제곱합 ESS"),
    ("condition_number", "조건수"),
]

# 모델별 카탈로그: (그룹, 라벨, 출력 식, 사용자에게 보여 줄 코드). {m}=모델 변수, {e}=최종 추정기
# 출력 식은 블록 전역에서 평가된다(np·pd import 여부와 무관하게 돌도록 헬퍼를 쓴다).
# 코드는 사용자가 자기 코드에 그대로 옮겨 쓸 수 있는 라이브러리 원래 표현이다.
_PYGRID_SM_CATALOG = [
    ("요약표", "계수표 (계수·표준오차·t·p·95% 신뢰구간)", "_pygrid_coef_table({m})", "{m}.summary2().tables[1]"),
    ("요약표", "적합 통계표 (R²·AIC 등 한 표)", "_pygrid_fit_stats({m})", "_pygrid_fit_stats({m})"),
    ("계수·추론", "계수 β", "{m}.params", "{m}.params"),
    ("계수·추론", "표준오차", "{m}.bse", "{m}.bse"),
    ("계수·추론", "t(z) 통계량", "{m}.tvalues", "{m}.tvalues"),
    ("계수·추론", "p값", "{m}.pvalues", "{m}.pvalues"),
    ("계수·추론", "95% 신뢰구간", "{m}.conf_int()", "{m}.conf_int()"),
    ("계수·추론", "오즈비 exp(β)", "_pygrid_odds_ratio({m})", "np.exp({m}.params)"),
    ("적합값·잔차", "적합값", "{m}.fittedvalues", "{m}.fittedvalues"),
    ("적합값·잔차", "잔차", "{m}.resid", "{m}.resid"),
    ("적합값·잔차", "피어슨 잔차", "{m}.resid_pearson", "{m}.resid_pearson"),
    ("적합값·잔차", "이탈도 잔차", "{m}.resid_deviance", "{m}.resid_deviance"),
    ("모델 정보", "공식", "{m}.model.formula", "{m}.model.formula"),
    ("모델 정보", "종속변수", "{m}.model.endog_names", "{m}.model.endog_names"),
    ("모델 정보", "설명변수", "{m}.model.exog_names", "{m}.model.exog_names"),
]

_PYGRID_SK_CATALOG = [
    ("요약표", "계수표 (절편 + 특성별 계수)", "_pygrid_coef_table({m})", "pd.Series({e}.coef_.ravel(), index={m}.feature_names_in_)"),
    ("요약표", "특성 중요도", "_pygrid_importances({m})", "pd.Series({e}.feature_importances_, index={m}.feature_names_in_)"),
    ("요약표", "모델 설정 (하이퍼파라미터)", "_pygrid_dict_table({e}.get_params())", "{e}.get_params()"),
    ("계수", "절편", "{e}.intercept_", "{e}.intercept_"),
    ("계수", "계수 (배열)", "{e}.coef_", "{e}.coef_"),
    ("교차검증", "선택된 규제 강도 α", "{e}.alpha_", "{e}.alpha_"),
    ("교차검증", "선택된 l1_ratio", "{e}.l1_ratio_", "{e}.l1_ratio_"),
    ("교차검증", "최적 파라미터", "_pygrid_dict_table({m}.best_params_)", "{m}.best_params_"),
    ("교차검증", "최고 교차검증 점수", "{m}.best_score_", "{m}.best_score_"),
    ("교차검증", "교차검증 결과표", "_pygrid_cv_table({m})", "pd.DataFrame({m}.cv_results_)"),
    ("모델 정보", "특성 수", "{e}.n_features_in_", "{e}.n_features_in_"),
    ("모델 정보", "특성 이름", "{m}.feature_names_in_", "{m}.feature_names_in_"),
    ("모델 정보", "클래스", "{e}.classes_", "{e}.classes_"),
    ("모델 정보", "반복 횟수", "{e}.n_iter_", "{e}.n_iter_"),
    ("군집", "군집 중심", "{e}.cluster_centers_", "{e}.cluster_centers_"),
    ("군집", "관성 (군집 내 제곱합)", "{e}.inertia_", "{e}.inertia_"),
    ("군집", "군집 레이블", "{e}.labels_", "{e}.labels_"),
    ("차원 축소", "설명 분산 비율", "{e}.explained_variance_ratio_", "{e}.explained_variance_ratio_"),
    ("차원 축소", "주성분 계수", "{e}.components_", "{e}.components_"),
]

def _pygrid_model_kind(v):
    """'statsmodels' | 'sklearn' | None — 적합이 끝난 모델만."""
    mod = type(v).__module__ or ""
    if mod.startswith("statsmodels") and hasattr(v, "params") and hasattr(v, "model"):
        return "statsmodels"
    if mod.startswith("sklearn") and hasattr(v, "get_params") and hasattr(v, "fit"):
        try:
            from sklearn.utils.validation import check_is_fitted

            check_is_fitted(v)
            return "sklearn"
        except Exception:
            return None
    return None


def _pygrid_scalar_attr(v, attr):
    """스칼라 숫자 속성 값 또는 None (없음·계산 실패·비스칼라)."""
    import numbers

    try:
        x = getattr(v, attr)
    except Exception:
        return None
    if callable(x) or isinstance(x, bool) or not isinstance(x, numbers.Number):
        try:
            import numpy as np

            if isinstance(x, np.ndarray) and x.size == 1:
                return float(x.reshape(-1)[0])
        except Exception:
            pass
        return None
    x = float(x)
    return x if x == x else None  # NaN 제외


def _pygrid_feature_names(m, n):
    """sklearn 모델(또는 Pipeline)의 입력 특성 이름 n개."""
    names = None
    if hasattr(m, "steps") and len(m.steps) > 1:
        try:
            names = list(m[:-1].get_feature_names_out())
        except Exception:
            names = None
    if names is None:
        names = list(getattr(m, "feature_names_in_", []))  # ndarray라 `or` 금지
    if len(names) != n:
        names = [f"x{i}" for i in range(n)]
    return [str(x) for x in names]


def _pygrid_coef_table(m):
    """계수표 DataFrame. statsmodels: 계수·표준오차·t·p·95% CI / sklearn: 절편 + 특성별 계수."""
    import numpy as np
    import pandas as pd

    kind = _pygrid_model_kind(m)
    if kind == "statsmodels":
        params = np.asarray(m.params, dtype=float).reshape(-1)
        names = getattr(m.model, "exog_names", None) or [f"x{i}" for i in range(len(params))]
        ci = np.asarray(m.conf_int(), dtype=float).reshape(len(params), 2)
        return pd.DataFrame(
            {
                "계수": params,
                "표준오차": np.asarray(m.bse, dtype=float).reshape(-1),
                "t(z)": np.asarray(m.tvalues, dtype=float).reshape(-1),
                "p값": np.asarray(m.pvalues, dtype=float).reshape(-1),
                "하한 95%": ci[:, 0],
                "상한 95%": ci[:, 1],
            },
            index=pd.Index([str(x) for x in names], name="변수"),
        ).reset_index()
    if kind == "sklearn":
        est = m.steps[-1][1] if hasattr(m, "steps") else m
        if not hasattr(est, "coef_"):
            raise TypeError(f"{type(est).__name__}에는 계수(coef_)가 없습니다")
        coef = np.atleast_2d(np.asarray(est.coef_, dtype=float))
        names = _pygrid_feature_names(m, coef.shape[1])
        intercept = np.atleast_1d(np.asarray(getattr(est, "intercept_", 0.0), dtype=float))
        cols = ["계수"] if coef.shape[0] == 1 else [f"계수[{k}]" for k in range(coef.shape[0])]
        body = pd.DataFrame(coef.T, index=names, columns=cols)
        head = pd.DataFrame([np.resize(intercept, coef.shape[0])], index=["(절편)"], columns=cols)
        out = pd.concat([head, body])
        out.index.name = "변수"
        return out.reset_index()
    raise TypeError(f"계수표를 만들 수 없는 객체입니다: {type(m).__name__}")


def _pygrid_label_table(head, rows):
    """{라벨: 값} → [라벨 | 값] 2열 표. 라벨을 index가 아닌 열로 둬 includeIndex 규칙과 무관하게 보인다."""
    import pandas as pd

    return pd.DataFrame({head: [str(k) for k in rows], "값": list(rows.values())})


def _pygrid_fit_stats(m):
    """적합 통계 한 표(Series). statsmodels 결과에 있는 항목만 싣는다."""
    import pandas as pd

    if _pygrid_model_kind(m) != "statsmodels":
        raise TypeError(
            "적합 통계는 statsmodels 결과에서만 바로 계산됩니다 — "
            "scikit-learn 모델은 '평가 지표(데이터셋·목표 열)'를 고르세요"
        )
    rows = {}
    for attr, label in _PYGRID_SM_STATS:
        x = _pygrid_scalar_attr(m, attr)
        if x is not None:
            rows[label] = x
    return _pygrid_label_table("통계", rows)


def _pygrid_design(m, data):
    """예측 입력 정렬: 공식 모델은 표 그대로, 그 외는 학습 때 열 이름·순서에 맞춘다."""
    import numpy as np
    import pandas as pd

    kind = _pygrid_model_kind(m)
    if not isinstance(data, pd.DataFrame):
        return data
    if kind == "statsmodels":
        if getattr(m.model, "formula", None) is not None:
            return data  # patsy가 공식대로 설계행렬을 만든다
        names = list(getattr(m.model, "exog_names", []) or [])
        X = data.copy()
        if "const" in names and "const" not in X.columns:
            X.insert(0, "const", 1.0)
        missing = [c for c in names if c not in X.columns]
        if missing:
            raise KeyError(f"예측 데이터에 학습 때 쓴 열이 없습니다: {missing}")
        return X[names]
    names = getattr(m, "feature_names_in_", None)
    if names is not None:
        missing = [c for c in names if c not in data.columns]
        if missing:
            raise KeyError(f"예측 데이터에 학습 때 쓴 열이 없습니다: {missing}")
        return data[list(names)]
    return data.select_dtypes(include=[np.number])


def _pygrid_predict(m, data, interval=None, level=0.95, keep=False, proba=False):
    """예측값(Series '예측'). interval='confidence'|'prediction'이면 하한·상한 열을 붙인 표.

    R의 predict(fit, newdata, interval=...)에 대응한다. keep=True면 원본 열 오른쪽에 붙인다.
    """
    import numpy as np
    import pandas as pd

    if not hasattr(m, "predict"):
        raise TypeError(f"예측할 수 없는 객체입니다: {type(m).__name__}")
    X = _pygrid_design(m, data)
    index = data.index if isinstance(data, (pd.DataFrame, pd.Series)) else None
    if proba:
        if not hasattr(m, "predict_proba"):
            raise TypeError("이 모델은 클래스 확률(predict_proba)을 지원하지 않습니다")
        P = np.asarray(m.predict_proba(X), dtype=float)
        classes = getattr(m, "classes_", range(P.shape[1]))
        out = pd.DataFrame(P, index=index, columns=[f"P({c})" for c in classes])
    elif interval:
        if _pygrid_model_kind(m) != "statsmodels" or not hasattr(m, "get_prediction"):
            raise TypeError("구간 예측은 statsmodels 결과에서만 지원합니다")
        fr = m.get_prediction(X).summary_frame(alpha=1 - float(level))
        key = "obs_ci" if interval == "prediction" else "mean_ci"
        if f"{key}_lower" not in fr.columns:
            raise TypeError("이 모델은 예측구간을 지원하지 않습니다 — 신뢰구간을 고르세요")
        tag = "예측구간" if interval == "prediction" else "신뢰구간"
        pct = f"{float(level):.0%}"
        out = pd.DataFrame(
            {
                "예측": fr["mean"].to_numpy(),
                f"{tag} 하한 {pct}": fr[f"{key}_lower"].to_numpy(),
                f"{tag} 상한 {pct}": fr[f"{key}_upper"].to_numpy(),
            },
            index=index,
        )
    else:
        y = np.asarray(m.predict(X), dtype=float)
        if y.ndim == 2 and y.shape[1] == 1:
            y = y[:, 0]
        if y.ndim != 1:
            out = pd.DataFrame(y, index=index, columns=[f"예측[{k}]" for k in range(y.shape[1])])
        else:
            out = pd.Series(y, index=index, name="예측")
    if keep and isinstance(data, pd.DataFrame):
        out = pd.concat([data, out.to_frame() if isinstance(out, pd.Series) else out], axis=1)
    return out


def _pygrid_eval_metrics(m, data, target):
    """평가 지표(Series): 관측치 수·R²·RMSE·MAE — 주어진 데이터셋과 목표 열 기준."""
    import numpy as np
    import pandas as pd

    if target not in data.columns:
        raise KeyError(f"목표 열 '{target}'이 데이터에 없습니다")
    pred = _pygrid_predict(m, data)
    p = np.asarray(pred, dtype=float).reshape(-1)
    y = data[target].to_numpy(dtype=float)
    ok = ~(np.isnan(y) | np.isnan(p))
    y, p = y[ok], p[ok]
    if len(y) == 0:
        raise ValueError("평가할 관측치가 없습니다 (결측 제외 후 0행)")
    sse = float(np.sum((y - p) ** 2))
    sst = float(np.sum((y - y.mean()) ** 2))
    return _pygrid_label_table(
        "지표",
        {
            "관측치 수": float(len(y)),
            "결정계수 R²": 1 - sse / sst if sst > 0 else float("nan"),
            "RMSE": (sse / len(y)) ** 0.5,
            "MAE": float(np.mean(np.abs(y - p))),
        },
    )


def _pygrid_odds_ratio(m):
    """오즈비 exp(β) — 로지스틱·이항 GLM 해석용."""
    import numpy as np

    return np.exp(m.params)


def _pygrid_importances(m):
    """특성 중요도 표 (트리·앙상블) — 큰 순."""
    import pandas as pd

    est = m[-1] if hasattr(m, "steps") else m
    imp = est.feature_importances_
    names = _pygrid_feature_names(m, len(imp))
    out = pd.DataFrame({"특성": names, "중요도": imp})
    return out.sort_values("중요도", ascending=False, ignore_index=True)


def _pygrid_dict_table(d):
    """dict → [항목 | 값] 표 (값은 문자열·숫자 그대로)."""
    import numbers

    return _pygrid_label_table(
        "항목", {k: (v if isinstance(v, numbers.Number) and not isinstance(v, bool) else str(v)) for k, v in d.items()}
    )


def _pygrid_cv_table(m):
    """교차검증 결과표 — 순위·평균 점수·파라미터 열만, 순위순."""
    import pandas as pd

    cv = pd.DataFrame(m.cv_results_)
    keep = [c for c in cv.columns if c.startswith("param_") or c in ("rank_test_score", "mean_test_score", "std_test_score")]
    return cv[keep].sort_values("rank_test_score", ignore_index=True).astype({c: str for c in keep if c.startswith("param_")})


def _pygrid_shape_preview(x):
    """값 → (예상 spill 크기 [행, 열], 짧은 미리보기 문자열)."""
    import numbers

    import numpy as np
    import pandas as pd

    if isinstance(x, bool) or isinstance(x, (numbers.Number, np.generic)):
        try:
            f = float(x)
            return [1, 1], f"{f:.6g}"
        except Exception:
            return [1, 1], str(x)
    if isinstance(x, str):
        return [1, 1], x[:60]
    if isinstance(x, pd.DataFrame):
        return [x.shape[0] + 1, x.shape[1] + 1], f"표 {x.shape[0]}×{x.shape[1]}"
    if isinstance(x, pd.Series):
        head = ", ".join(f"{k}={v:.4g}" if isinstance(v, float) else f"{k}={v}" for k, v in list(x.items())[:3])
        more = " …" if len(x) > 3 else ""
        return [len(x) + 1, 2], f"{len(x)}개: {head}{more}"
    if isinstance(x, (list, tuple)):
        x = np.asarray(x, dtype=object)
    if isinstance(x, np.ndarray):
        if x.ndim <= 1:
            n = x.size
            head = ", ".join(f"{v:.4g}" if isinstance(v, float) else str(v) for v in x.reshape(-1)[:3])
            return [max(n, 1), 1], f"{n}개: {head}{' …' if n > 3 else ''}"
        return [x.shape[0], x.shape[1] if x.ndim > 1 else 1], f"배열 {x.shape[0]}×{x.shape[1]}"
    return [1, 1], type(x).__name__


def _pygrid_model_info(name, v):
    """inspect용 모델 카탈로그 — 이 모델에 실제로 있는 항목만(평가가 성공한 것) 식·코드·미리보기·크기와 함께."""
    kind = _pygrid_model_kind(v)
    if kind is None:
        return None
    g = globals()
    est = v[-1] if hasattr(v, "steps") else v
    e = f"{name}[-1]" if est is not v else name
    members = []
    seen = set()

    def add(group, label, expr, code):
        expr = expr.format(m=name, e=e)
        if expr in seen:
            return
        try:
            x = eval(compile(expr, "<pygrid>", "eval"), g)
        except Exception:
            return  # 이 모델에는 없는 항목
        if x is None or callable(x):
            return
        if "오즈비" in label:
            # 로지스틱·이항 계열에서만 의미가 있다
            if type(v.model).__name__ not in ("Logit", "Probit") and "Binomial" not in str(getattr(v.model, "family", "")):
                return
        shape, preview = _pygrid_shape_preview(x)
        seen.add(expr)
        members.append(
            {"group": group, "label": label, "expr": expr, "code": code.format(m=name, e=e), "shape": shape, "preview": preview}
        )

    if kind == "statsmodels":
        for group, label, expr, code in _PYGRID_SM_CATALOG[:2]:
            add(group, label, expr, code)
        for attr, label in _PYGRID_SM_STATS:
            if _pygrid_scalar_attr(v, attr) is not None:
                add("적합 통계", label, "{m}." + attr, "{m}." + attr)
        for group, label, expr, code in _PYGRID_SM_CATALOG[2:]:
            add(group, label, expr, code)
    else:
        for group, label, expr, code in _PYGRID_SK_CATALOG:
            add(group, label, expr, code)
        # 목록에 없는 학습 속성(끝이 _)도 보여 준다 — 모델마다 다르다
        known = {m["expr"] for m in members}
        for attr in sorted(getattr(est, "__dict__", {})):
            if not attr.endswith("_") or attr.startswith("_"):
                continue
            expr = f"{e}.{attr}"
            if expr not in known:
                add("기타 학습 속성", attr, expr, expr)
    return {
        "kind": kind,
        "formula": kind == "statsmodels" and getattr(v.model, "formula", None) is not None,
        "intervals": kind == "statsmodels" and hasattr(v, "get_prediction"),
        "predict": hasattr(v, "predict"),
        "proba": hasattr(v, "predict_proba"),
        "featureNames": getattr(v, "feature_names_in_", None) is not None,
        "members": members,
    }
