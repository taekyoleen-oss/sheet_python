// 모델적합 워크북 가이드 (부록 H.3) — 마법사 선택(데이터·모형·후보 분포)으로
// [마크다운 설명 + 코드] 단계 블록 시퀀스를 만든다. 코드는 데이터를 내장하지 않고
// 항상 sheet("범위", headers=True) 참조를 쓴다(데이터의 원본은 그리드).
// 순수 코어: buildFitGuideBlocks / 불순 진입점: runFitGuide (한 스토어 트랜잭션 = 한 undo).
// 분포 파라미터화는 lib/reference/fitPython.ts·distributions.ts pySpec와 일치
// (로그정규 s/scale·floc=0, 지수 scale=1/lambda, 음이항 프로파일 우도 등). fitPython은 수정하지 않는다.

import { toast } from "sonner";
import { cellKey, type Sheet } from "@/types/workbook";
import { colToLetter, formatA1, parseA1 } from "./a1";
import { createReferenceBlocks, type SendSection } from "./import-blocks";
import { useWorkbookStore } from "./model";
import { saveSettings } from "@/lib/storage/db";

/* ─────────────────────────── 카탈로그 ─────────────────────────── */

export type FitModel = "severity" | "frequency" | "compound";

export const FIT_MODEL_LABEL: Record<FitModel, string> = {
  severity: "심도(개별 손해액)",
  frequency: "빈도(사고 건수)",
  compound: "합성(총손해 S)",
};

/** 가이드가 코드를 생성할 수 있는 심도(연속형) 분포 — def: 마법사 기본 체크 */
export const FIT_GUIDE_SEV_DISTS = [
  { id: "lognormal", name: "로그정규", def: true },
  { id: "gamma", name: "감마", def: true },
  { id: "weibull", name: "와이블", def: true },
  { id: "exponential", name: "지수", def: false },
  { id: "pareto2", name: "파레토(2모수)", def: false },
  { id: "normal", name: "정규", def: false },
] as const;

/** 빈도(이산형) 분포 */
export const FIT_GUIDE_FREQ_DISTS = [
  { id: "poisson", name: "포아송", def: true },
  { id: "negbinom", name: "음이항", def: true },
] as const;

/** 마법사 1단계 샘플 데이터셋 (public/samples) — sevCol/freqCol은 0-based 열 번호 */
export const FIT_GUIDE_SAMPLES: {
  file: string;
  label: string;
  sevCol: number;
  freqCol?: number;
}[] = [
  { file: "claims.xlsx", label: "claims.xlsx — 청구·손해액 600행 (claim_amt·claim_cnt)", sevCol: 7, freqCol: 8 },
  { file: "policy.xlsx", label: "policy.xlsx — 계약·보험료 600행 (premium·n_contracts)", sevCol: 7, freqCol: 12 },
  { file: "experience.xlsx", label: "experience.xlsx — 경험데이터 800행 (duration_years)", sevCol: 4 },
  { file: "triangle.xlsx", label: "triangle.xlsx — 삼각형 1년차 지급 8행 (dev_1)", sevCol: 1 },
  { file: "mortality_table.xlsx", label: "mortality_table.xlsx — 남자 사망률 101행 (qx_male)", sevCol: 1 },
];

/* ─────────────────────────── 순수 코어 ─────────────────────────── */

export interface FitGuideRange {
  /** sheet() 인수 문자열 — 시트 접두어 포함, 헤더 행 포함 (예: "청구액!A1:A301") */
  ref: string;
  /** 실제 헤더(1~2개) — 마지막이 값(손해액·건수) 열 */
  headers: string[];
}

export interface FitGuideSpec {
  model: FitModel;
  /** 심도·빈도 분포 id 혼합 — 카탈로그로 구분한다 */
  dists: string[];
  sev?: FitGuideRange;
  freq?: FitGuideRange;
}

export interface GuideBlock {
  kind: "markdown" | "code";
  title: string;
  body: string;
  /** 코드 블록 출력 모드 (기본 values) */
  outputMode?: "values" | "object";
  /** 아래로 비워둘 행 수(spill 여유) */
  reserve?: number;
}

const nameOfSev = (id: string) => FIT_GUIDE_SEV_DISTS.find((d) => d.id === id)?.name ?? id;
const nameOfFreq = (id: string) => FIT_GUIDE_FREQ_DISTS.find((d) => d.id === id)?.name ?? id;

/** 심도 분포별 MLE 스탠자 — fits.append((이름, frozen, k, 파라미터 문자열)) */
const SEV_FIT_STANZA: Record<string, string> = {
  lognormal: `s, _, sc = stats.lognorm.fit(x, floc=0)      # 로그정규: floc=0, s=sigma, scale=e^mu
fits.append(("로그정규", stats.lognorm(s, 0, sc), 2, f"mu={np.log(sc):.4g}, sigma={s:.4g}"))`,
  gamma: `a, _, sc = stats.gamma.fit(x, floc=0)        # 감마: alpha=형상, theta=척도 (평균=alpha*theta)
fits.append(("감마", stats.gamma(a, 0, sc), 2, f"alpha={a:.4g}, theta={sc:.4g}"))`,
  weibull: `c, _, sc = stats.weibull_min.fit(x, floc=0)  # 와이블: k=형상(고장률 증감), lambda=척도
fits.append(("와이블", stats.weibull_min(c, 0, sc), 2, f"k={c:.4g}, lambda={sc:.4g}"))`,
  exponential: `_, sc = stats.expon.fit(x, floc=0)           # 지수: lambda = 1/scale = 1/표본평균
fits.append(("지수", stats.expon(0, sc), 1, f"lambda={1/sc:.4g}"))`,
  pareto2: `c, _, sc = stats.lomax.fit(x, floc=0)        # 2모수 파레토(Lomax) — 두꺼운 꼬리 손해분포
fits.append(("파레토(2모수)", stats.lomax(c, 0, sc), 2, f"alpha={c:.4g}, theta={sc:.4g}"))`,
  normal: `mu, sg = stats.norm.fit(x)                   # 정규: mu=평균, sigma=표준편차
fits.append(("정규", stats.norm(mu, sg), 2, f"mu={mu:.4g}, sigma={sg:.4g}"))`,
};

/** 빈도 분포별 MLE 스탠자 — counts에 적합 */
const FREQ_FIT_STANZA: Record<string, string> = {
  poisson: `lam = counts.mean()                          # 포아송 MLE: lambda = 평균 건수
fits.append(("포아송", stats.poisson(lam), 1, f"lambda={lam:.4g}"))`,
  negbinom: `from scipy.optimize import minimize_scalar   # 음이항: p=r/(r+평균), r만 수치최적화(프로파일 우도)
_m = counts.mean()
_nll = lambda lr: -float(np.sum(stats.nbinom.logpmf(counts, np.exp(lr), np.exp(lr) / (np.exp(lr) + _m))))
_r = float(np.exp(minimize_scalar(_nll, bounds=(np.log(1e-2), np.log(1e4)), method="bounded").x))
fits.append(("음이항", stats.nbinom(_r, _r / (_r + _m)), 2, f"r={_r:.4g}, p={_r/(_r+_m):.4g}"))`,
};

const sevStanzas = (ids: string[]) => ids.map((id) => SEV_FIT_STANZA[id]).filter(Boolean).join("\n");
const freqStanzas = (ids: string[]) => ids.map((id) => FREQ_FIT_STANZA[id]).filter(Boolean).join("\n");

/** sheet() 로드 줄 — dfVar 이름은 합성에서 df_x/df_n으로 분리 */
const loadSev = (r: FitGuideRange, dfVar = "df") =>
  `${dfVar} = sheet("${r.ref}", headers=True)
x = ${dfVar}["${r.headers[r.headers.length - 1]}"].dropna().astype(float).to_numpy()`;

const loadFreq = (r: FitGuideRange, dfVar = "df") =>
  `${dfVar} = sheet("${r.ref}", headers=True)
counts = ${dfVar}["${r.headers[r.headers.length - 1]}"].dropna().astype(int).to_numpy()`;

/** 심도 비교표 공통 꼬리 — DataFrame이 마지막 표현식(값 모드 spill) */
const SEV_TABLE_TAIL = `
rows = []
for name, dist, k, par in fits:
    logL = float(np.sum(dist.logpdf(x)))          # 로그우도 — 클수록 좋음
    ks = stats.kstest(x, dist.cdf)                # KS: 경험 CDF와의 최대 거리
    rows.append({"분포": name, "파라미터": par, "logL": round(logL, 2),
                 "AIC": round(2 * k - 2 * logL, 2),
                 "BIC": round(k * np.log(len(x)) - 2 * logL, 2),
                 "KS D": round(float(ks.statistic), 4), "KS p": round(float(ks.pvalue), 4)})
pd.DataFrame(rows).sort_values("AIC").reset_index(drop=True)`;

/** 빈도 비교표 공통 꼬리 — 이산분포라 KS 대신 카이제곱 적합도 */
const FREQ_TABLE_TAIL = `
rows = []
K = int(counts.max())
obs = np.bincount(counts, minlength=K + 1).astype(float)
for name, dist, k, par in fits:
    logL = float(np.sum(dist.logpmf(counts)))
    E = len(counts) * np.array([float(dist.pmf(j)) for j in range(K + 1)])
    E[-1] += len(counts) * float(dist.sf(K))      # K 초과 꼬리는 마지막 칸에 합산
    ok = E > 0
    chi2 = float(np.sum((obs[ok] - E[ok]) ** 2 / E[ok]))
    df_ = int(ok.sum()) - 1 - k
    rows.append({"분포": name, "파라미터": par, "logL": round(logL, 2),
                 "AIC": round(2 * k - 2 * logL, 2),
                 "BIC": round(k * np.log(len(counts)) - 2 * logL, 2),
                 "chi2": round(chi2, 2),
                 "chi2 p": round(float(stats.chi2.sf(chi2, df_)), 4) if df_ > 0 else None})
pd.DataFrame(rows).sort_values("AIC").reset_index(drop=True)`;

/** AIC 최소 심도 분포 선택 + PDF/CDF 오버레이 + QQ — fig 마지막 표현식 */
const sevVerifyCode = (r: FitGuideRange, ids: string[]) => `from scipy import stats
import matplotlib.pyplot as plt
${loadSev(r)}

fits = []
${sevStanzas(ids)}

def _aic(t):
    return 2 * t[2] - 2 * float(np.sum(t[1].logpdf(x)))
name, dist, k, par = min(fits, key=_aic)
print(f"AIC 최소: {name} ({par})")

xs = np.sort(x); n = len(xs)
xg = np.linspace(float(xs[0]), float(xs[-1]), 300)
fig, ax = plt.subplots(1, 3, figsize=(10.5, 3.4))
ax[0].hist(x, bins=40, density=True, alpha=0.35, color="#4A90C2", label="empirical")
ax[0].plot(xg, dist.pdf(xg), "r-", label=name); ax[0].set_title("PDF"); ax[0].legend()
ax[1].plot(xs, np.arange(1, n + 1) / n, drawstyle="steps-post", color="#4A90C2", label="ECDF")
ax[1].plot(xg, dist.cdf(xg), "r-", label=name); ax[1].set_title("CDF"); ax[1].legend()
pp = (np.arange(1, n + 1) - 0.5) / n
ax[2].scatter(dist.ppf(pp), xs, s=8, color="#4A90C2")
lim = [float(xs[0]), float(xs[-1])]
ax[2].plot(lim, lim, "k--", lw=1)
ax[2].set_title("Q-Q"); ax[2].set_xlabel("이론 분위수"); ax[2].set_ylabel("관측값")
fig.tight_layout()
fig`;

/** 몬테카를로 합성 시뮬레이션 앞부분(최적 분포 선택 + S 생성) — ⑤a/⑤b 공용 */
const compoundSimHead = (
  sev: FitGuideRange,
  freq: FitGuideRange,
  sevIds: string[],
  freqIds: string[],
) => `from scipy import stats
${loadSev(sev, "df_x")}
${loadFreq(freq, "df_n")}

# 심도·빈도 각각 AIC 최소 분포를 다시 적합해 선택합니다(블록 단독 실행 가능)
fits = []
${sevStanzas(sevIds)}
def _aic_x(t):
    return 2 * t[2] - 2 * float(np.sum(t[1].logpdf(x)))
sev_name, sev, _k1, sev_par = min(fits, key=_aic_x)

fits = []
${freqStanzas(freqIds)}
def _aic_n(t):
    return 2 * t[2] - 2 * float(np.sum(t[1].logpmf(counts)))
freq_name, freq, _k2, freq_par = min(fits, key=_aic_n)
print(f"심도 {sev_name} ({sev_par}) × 빈도 {freq_name} ({freq_par})")

rng = np.random.default_rng(42)              # seed 고정 = 재현 가능
n_years = 10_000                             # 시뮬레이션 연수
N = freq.rvs(n_years, random_state=rng)      # 그 해의 사고건수 N
total = int(N.sum())
all_x = sev.rvs(total, random_state=rng)     # 사고 1건당 손해액 X
idx = np.r_[0, np.cumsum(N)]
S = np.array([all_x[idx[i]:idx[i + 1]].sum() for i in range(n_years)])`;

/* 단계 마크다운 — 각 단계의 목적·확인 포인트(보험 용어) */

const md1 = (ref: string, col: string) => `## 1단계 — 데이터 확인

sheet() 참조로 그리드의 \`${ref}\` 범위를 불러옵니다 — 데이터의 원본은 항상 시트입니다.
행·열 크기와 값 열("${col}")이 의도한 범위와 일치하는지, 결측이나 이상값이 섞여 있지 않은지 확인하세요.`;

const MD2_SEV = `## 2단계 — 경험적 분석

분포를 가정하기 전에 데이터 자체의 생김새를 봅니다. 손해액(심도) 분포는 보통 오른쪽 꼬리가 길어(왜도>0)
로그정규·감마 같은 양수 분포가 후보가 됩니다. 히스토그램의 꼬리 두께와 90/95/99% 분위수를 확인하세요.`;

const MD2_FREQ = `## 2단계 — 경험적 분석

건수(빈도) 데이터의 평균과 분산을 비교합니다. 분산이 평균보다 크면 과산포 —
포아송(평균=분산)보다 음이항이 맞을 가능성이 큽니다. 건수 분포 막대에서 0건 비중과 꼬리도 함께 보세요.`;

const md3 = (names: string[], freqNames?: string[]) => `## 3단계 — 후보 분포 적합

선택한 후보 분포(${names.join("·")}${freqNames ? ` / 빈도 ${freqNames.join("·")}` : ""})를 scipy
최대우도추정(MLE)으로 적합하고 logL·AIC·BIC·적합도 통계로 비교합니다.
AIC·BIC는 작을수록 좋고, 비교표는 AIC 오름차순으로 정렬되어 셀에 깔립니다(값 모드 spill).`;

const MD4 = `## 4단계 — 최적 분포 검증

AIC 최소 분포를 경험 분포 위에 겹쳐 그려 순위표만으로 놓치는 것을 확인합니다.
PDF·CDF가 잘 얹히는지, Q-Q 점이 45° 선에 가까운지 보세요 —
오른쪽 끝이 위로 휘면 실제 꼬리가 모형보다 두껍다는(대형 손해 과소평가) 신호입니다.`;

const MD5 = `## 5단계 — 몬테카를로 합성

빈도 N과 심도 X의 최적 적합으로 '가상의 1년'을 수만 번 반복해 연간 총손해 S = X₁+…+X_N 분포를 만듭니다.
VaR는 분위수("이 확률로는 이 이하"), TVaR는 그 분위수를 넘는 최악 연도들의 평균 손해 —
요구자본·재보험 초과손해(XL) 층 설계의 근거가 됩니다.`;

/**
 * 순수 코어 — 마법사 선택을 [마크다운|코드] 블록 시퀀스로 변환한다.
 * 첫 블록은 제목 마크다운(# 모델적합 — …). 코드는 블록 단독 실행이 가능하도록
 * 각자 sheet() 로드·적합을 포함한다(블록 간 변수 공유에 의존하지 않는다).
 */
export function buildFitGuideBlocks(spec: FitGuideSpec): GuideBlock[] {
  const sevIds = spec.dists.filter((d) => SEV_FIT_STANZA[d]);
  const freqIds = spec.dists.filter((d) => FREQ_FIT_STANZA[d]);
  const out: GuideBlock[] = [];
  const md = (title: string, body: string) => out.push({ kind: "markdown", title, body });
  const code = (title: string, body: string, outputMode: "values" | "object", reserve?: number) =>
    out.push({ kind: "code", title, body, outputMode, ...(reserve ? { reserve } : {}) });

  if (spec.model === "severity" || spec.model === "compound") {
    const sev = spec.sev;
    if (!sev) throw new Error("심도 범위가 필요합니다");
    const col = sev.headers[sev.headers.length - 1];
    const sevNames = sevIds.map(nameOfSev);

    if (spec.model === "severity") {
      md(
        "제목",
        `# 모델적합 — ${FIT_MODEL_LABEL.severity}

\`${sev.ref}\` 범위의 개별 손해액을 후보 분포(${sevNames.join("·")})에 적합하고 AIC 기준으로 최적 분포를 고르는
단계별 가이드입니다. 각 단계의 설명을 읽고 코드 블록을 순서대로 실행하세요(목차 = 진행 가이드).`,
      );
      md("1단계", md1(sev.ref, col));
      code(
        "데이터 확인",
        `${loadSev(sev)}
print("행·열:", df.shape)
print("값 열:", "${col}", "· 결측 제외 n =", len(x))
df.head()`,
        "values",
        10,
      );
      md("2단계", MD2_SEV);
      code(
        "경험적 분석 — 요약·히스토그램",
        `import matplotlib.pyplot as plt
${loadSev(sev)}
print(pd.Series(x).describe().round(2))
print("분위수 90/95/99%:", [round(float(np.quantile(x, q)), 2) for q in (0.90, 0.95, 0.99)])
print("왜도:", round(float(pd.Series(x).skew()), 3), "— 양수면 오른쪽 꼬리(대형 손해)")
fig, ax = plt.subplots(figsize=(7, 3.4))
ax.hist(x, bins=40, color="#4A90C2", alpha=0.75)
ax.set_title("${col} 분포 — 경험적 히스토그램")
ax.set_xlabel("${col}")
fig`,
        "object",
      );
      md("3단계", md3(sevNames));
      code(
        "후보 분포 적합 · 비교표",
        `from scipy import stats
${loadSev(sev)}

fits = []
${sevStanzas(sevIds)}
${SEV_TABLE_TAIL}`,
        "values",
        Math.max(4, sevIds.length + 3) + 2,
      );
      md("4단계", MD4);
      code("최적 분포 검증", sevVerifyCode(sev, sevIds), "object");
      return out;
    }

    // ── compound: 심도+빈도 두 범위 ──
    const freq = spec.freq;
    if (!freq) throw new Error("빈도 범위가 필요합니다");
    const nCol = freq.headers[freq.headers.length - 1];
    const freqNames = freqIds.map(nameOfFreq);

    md(
      "제목",
      `# 모델적합 — ${FIT_MODEL_LABEL.compound}

심도 \`${sev.ref}\` × 빈도 \`${freq.ref}\` 범위로 집합손해모형(연간 총손해 S = X₁+…+X_N)을 만드는
단계별 가이드입니다. 심도·빈도를 각각 적합한 뒤 몬테카를로로 합성해 VaR·TVaR를 산출합니다.`,
    );
    md("1단계", `## 1단계 — 데이터 확인

심도(개별 손해액) \`${sev.ref}\`와 빈도(건수) \`${freq.ref}\` 두 범위를 sheet() 참조로 불러옵니다.
두 범위의 크기·값 열("${col}", "${nCol}")이 맞는지, 결측이 섞여 있지 않은지 확인하세요.`);
    code(
      "데이터 확인",
      `${loadSev(sev, "df_x")}
${loadFreq(freq, "df_n")}
print("심도 범위:", df_x.shape, "· 값 열: ${col} · n =", len(x))
print("빈도 범위:", df_n.shape, "· 건수 열: ${nCol} · n =", len(counts))
df_x.head()`,
      "values",
      10,
    );
    md("2단계", `## 2단계 — 경험적 분석

심도는 히스토그램으로 꼬리 두께를, 빈도는 평균 대비 분산(과산포)을 봅니다.
분산/평균이 1보다 크면 음이항 후보, 손해액 왜도가 크면 로그정규·파레토 계열 후보입니다.`);
    code(
      "경험적 분석 — 심도·빈도",
      `import matplotlib.pyplot as plt
${loadSev(sev, "df_x")}
${loadFreq(freq, "df_n")}
print("심도:", pd.Series(x).describe().round(2).to_dict())
m, v = float(counts.mean()), float(counts.var(ddof=1))
print("빈도 평균:", round(m, 3), "· 분산:", round(v, 3), "· 과산포 지수:", round(v / m, 3))
K = int(counts.max())
obs = np.bincount(counts, minlength=K + 1)
fig, ax = plt.subplots(1, 2, figsize=(9, 3.4))
ax[0].hist(x, bins=40, color="#4A90C2", alpha=0.75); ax[0].set_title("심도 ${col}")
ax[1].bar(np.arange(K + 1), obs / len(counts), width=0.6, color="#4A90C2", alpha=0.75)
ax[1].set_title("빈도 P(N=k)"); ax[1].set_xlabel("건수 k")
fig.tight_layout()
fig`,
      "object",
    );
    md("3단계", md3(sevNames, freqNames));
    code(
      "심도 후보 적합 · 비교표",
      `from scipy import stats
${loadSev(sev)}

fits = []
${sevStanzas(sevIds)}
${SEV_TABLE_TAIL}`,
      "values",
      Math.max(4, sevIds.length + 3) + 2,
    );
    code(
      "빈도 후보 적합 · 비교표",
      `from scipy import stats
${loadFreq(freq)}

fits = []
${freqStanzas(freqIds)}
${FREQ_TABLE_TAIL}`,
      "values",
      Math.max(4, freqIds.length + 3) + 2,
    );
    md("4단계", MD4);
    code("최적 심도 분포 검증", sevVerifyCode(sev, sevIds), "object");
    md("5단계", MD5);
    code(
      "몬테카를로 합성 — VaR·TVaR 요약",
      `${compoundSimHead(sev, freq, sevIds, freqIds)}

var95, var99 = float(np.quantile(S, 0.95)), float(np.quantile(S, 0.99))
pd.DataFrame({
    "지표": ["평균", "표준편차", "VaR 95%", "TVaR 95%", "VaR 99%", "TVaR 99%"],
    "값": [round(float(S.mean()), 1), round(float(S.std()), 1),
           round(var95, 1), round(float(S[S >= var95].mean()), 1),
           round(var99, 1), round(float(S[S >= var99].mean()), 1)],
})`,
      "values",
      11,
    );
    code(
      "연간 총손해 S 분포",
      `import matplotlib.pyplot as plt
${compoundSimHead(sev, freq, sevIds, freqIds)}

fig, ax = plt.subplots(figsize=(7, 3.4))
ax.hist(S, bins=60, color="#4A90C2", alpha=0.75)
ax.axvline(float(np.quantile(S, 0.99)), color="r", ls="--", lw=1, label="VaR 99%")
ax.set_title("연간 총손해 S — 몬테카를로"); ax.legend()
fig`,
      "object",
    );
    return out;
  }

  // ── frequency ──
  const freq = spec.freq;
  if (!freq) throw new Error("빈도 범위가 필요합니다");
  const nCol = freq.headers[freq.headers.length - 1];
  const freqNames = freqIds.map(nameOfFreq);

  md(
    "제목",
    `# 모델적합 — ${FIT_MODEL_LABEL.frequency}

\`${freq.ref}\` 범위의 건수 데이터를 빈도 분포(${freqNames.join("·")})에 적합하는 단계별 가이드입니다.
빈도 모형은 "단위 기간(연도·계약)에 사고가 몇 건 나는가"의 분포로, 요율산정과 집합손해모형의 재료가 됩니다.`,
  );
  md("1단계", md1(freq.ref, nCol));
  code(
    "데이터 확인",
    `${loadFreq(freq)}
print("행·열:", df.shape)
print("건수 열:", "${nCol}", "· 결측 제외 n =", len(counts))
df.head()`,
    "values",
    10,
  );
  md("2단계", MD2_FREQ);
  code(
    "경험적 분석 — 과산포 확인",
    `import matplotlib.pyplot as plt
${loadFreq(freq)}
m, v = float(counts.mean()), float(counts.var(ddof=1))
print("평균:", round(m, 3), "· 분산:", round(v, 3))
print("과산포 지수(분산/평균):", round(v / m, 3), "— 1보다 크면 음이항 후보")
K = int(counts.max())
obs = np.bincount(counts, minlength=K + 1)
fig, ax = plt.subplots(figsize=(7, 3.4))
ax.bar(np.arange(K + 1), obs / len(counts), width=0.6, color="#4A90C2", alpha=0.75)
ax.set_title("건수 분포 P(N=k) — 경험적")
ax.set_xlabel("건수 k")
fig`,
    "object",
  );
  md("3단계", md3(freqNames));
  code(
    "후보 분포 적합 · 비교표",
    `from scipy import stats
${loadFreq(freq)}

fits = []
${freqStanzas(freqIds)}
${FREQ_TABLE_TAIL}`,
    "values",
    Math.max(4, freqIds.length + 3) + 2,
  );
  md("4단계", `## 4단계 — 최적 분포 검증

AIC 최소 분포의 PMF를 관측 비율 막대 위에 겹쳐 그립니다. 0건 비중과 꼬리(많이 나는 해)가
모형과 맞는지 보세요 — 관측 0건이 모형보다 훨씬 많으면 제로팽창(ZIP·ZINB) 계열 검토 신호입니다.`);
  code(
    "최적 분포 검증",
    `from scipy import stats
import matplotlib.pyplot as plt
${loadFreq(freq)}

fits = []
${freqStanzas(freqIds)}

def _aic(t):
    return 2 * t[2] - 2 * float(np.sum(t[1].logpmf(counts)))
name, dist, k, par = min(fits, key=_aic)
print(f"AIC 최소: {name} ({par})")

K = int(counts.max())
kk = np.arange(0, K + 3)
obs = np.bincount(counts, minlength=K + 1) / len(counts)
ecdf = np.array([float(np.mean(counts <= t)) for t in kk])
fig, ax = plt.subplots(1, 2, figsize=(9, 3.4))
ax[0].bar(np.arange(K + 1), obs, width=0.5, alpha=0.4, color="#4A90C2", label="관측")
ax[0].plot(kk, dist.pmf(kk), "ro-", ms=4, lw=1, label=name)
ax[0].set_title("PMF"); ax[0].set_xlabel("건수 k"); ax[0].legend()
ax[1].step(kk, ecdf, where="post", color="#4A90C2", label="관측 CDF")
ax[1].plot(kk, dist.cdf(kk), "r-", label=name)
ax[1].set_title("CDF"); ax[1].legend()
fig.tight_layout()
fig`,
    "object",
  );
  return out;
}

/* ─────────────────────────── 불순 진입점 ─────────────────────────── */

export type FitGuideInput = {
  model: FitModel;
  dists: string[];
  data:
    | { mode: "sample"; sample: string }
    | { mode: "selection"; sevRef?: string; freqRef?: string };
};

/** A1 참조 문자열 → FitGuideRange (헤더는 시트 첫 행에서 읽는다) */
function resolveRef(refText: string | undefined, label: string): FitGuideRange | { error: string } {
  if (!refText || !refText.trim())
    return { error: `${label} 범위를 입력하세요 (예: A1:A301, 헤더 행 포함)` };
  let parsed;
  try {
    parsed = parseA1(refText.trim());
  } catch (e) {
    return { error: (e as Error).message };
  }
  const st = useWorkbookStore.getState();
  const sheet = parsed.sheetName
    ? st.workbook.sheets.find((s) => s.name === parsed.sheetName)
    : st.workbook.sheets.find((s) => s.id === st.activeSheetId);
  if (!sheet) return { error: `시트를 찾을 수 없습니다: ${parsed.sheetName ?? ""}` };
  const { r0, c0, r1, c1 } = parsed.range;
  if (c1 - c0 + 1 > 2)
    return { error: `${label} 범위는 1~2열이어야 합니다 (마지막 열 = 값)` };
  if (r1 - r0 < 2)
    return { error: `${label} 범위가 너무 짧습니다 — 헤더 행 + 데이터 행이 필요합니다` };
  const headers: string[] = [];
  for (let c = c0; c <= c1; c++) {
    const cell = sheet.cells[cellKey(r0, c)];
    if (!cell || cell.t !== "s" || !String(cell.v ?? "").trim())
      return {
        error: `${label} 범위 첫 행은 헤더(열 이름)여야 합니다 — 문자열 헤더 행을 포함해 선택하세요`,
      };
    headers.push(String(cell.v).trim());
  }
  return { ref: formatA1(parsed.range, sheet.name), headers };
}

/** 샘플 시트의 한 열 전체(헤더 포함) → FitGuideRange */
function columnRange(sheet: Sheet, c: number): FitGuideRange {
  let maxR = 0;
  for (const key of Object.keys(sheet.cells)) {
    const r = Number(key.slice(0, key.indexOf(":")));
    if (r > maxR) maxR = r;
  }
  const h = String(sheet.cells[cellKey(0, c)]?.v ?? "").trim() || colToLetter(c);
  return { ref: formatA1({ r0: 0, c0: c, r1: maxR, c1: c }, sheet.name), headers: [h] };
}

/**
 * 마법사 완료 진입점 — 샘플이면 그리드로 가져온 뒤(importData, 시트만·블록 없음)
 * 단계 블록을 한 스토어 트랜잭션으로 생성하고 워크북 뷰로 전환한다. 블록은 자동 실행하지 않는다.
 */
export async function runFitGuide(
  input: FitGuideInput,
): Promise<{ ids: string[] } | { error: string }> {
  const needSev = input.model !== "frequency";
  const needFreq = input.model !== "severity";
  const sevIds = input.dists.filter((d) => SEV_FIT_STANZA[d]);
  const freqIds = input.dists.filter((d) => FREQ_FIT_STANZA[d]);
  if (needSev && sevIds.length === 0) return { error: "심도 후보 분포를 하나 이상 선택하세요" };
  if (needFreq && freqIds.length === 0) return { error: "빈도 후보 분포를 하나 이상 선택하세요" };

  let sev: FitGuideRange | undefined;
  let freq: FitGuideRange | undefined;

  if (input.data.mode === "sample") {
    const sampleFile = input.data.sample;
    const meta = FIT_GUIDE_SAMPLES.find((s) => s.file === sampleFile);
    if (!meta) return { error: "알 수 없는 샘플 데이터셋입니다" };
    if (needFreq && meta.freqCol === undefined)
      return { error: "이 샘플에는 건수 열이 없어 빈도·합성 모형을 쓸 수 없습니다" };
    let bytes: Uint8Array;
    try {
      const res = await fetch(`/samples/${meta.file}`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      bytes = new Uint8Array(await res.arrayBuffer());
    } catch (e) {
      return { error: `샘플을 불러오지 못했습니다: ${(e as Error).message}` };
    }
    const { importData } = await import("@/lib/io/data-import");
    const imported = await importData(meta.file, bytes, {
      toSheet: true,
      toFs: false,
      makeBlock: "none",
    });
    const st = useWorkbookStore.getState();
    const sheet = st.workbook.sheets.find((s) => s.name === imported.sheetNames[0]);
    if (!sheet) return { error: "샘플 시트를 만들지 못했습니다" };
    if (needSev) sev = columnRange(sheet, meta.sevCol);
    if (needFreq) freq = columnRange(sheet, meta.freqCol!);
  } else {
    if (needSev) {
      const r = resolveRef(input.data.sevRef, "심도");
      if ("error" in r) return r;
      sev = r;
    }
    if (needFreq) {
      const r = resolveRef(input.data.freqRef, "빈도");
      if ("error" in r) return r;
      freq = r;
    }
  }

  const guide = buildFitGuideBlocks({ model: input.model, dists: input.dists, sev, freq });
  const sections: SendSection[] = guide.map((g) =>
    g.kind === "markdown"
      ? { title: g.title, code: "", markdown: g.body }
      : { title: g.title, code: g.body, outputMode: g.outputMode, reserve: g.reserve },
  );
  const ids = createReferenceBlocks(null, sections);
  if (ids.length === 0) return { error: "블록을 만들 수 없습니다 (활성 시트 없음)" };

  const st = useWorkbookStore.getState();
  st.setView("workbook");
  void saveSettings({ view: "workbook" });
  st.setFocusBlock(ids[0]);
  st.setSelectedBlock(ids[0]);
  toast(`단계별 블록 ${ids.length}개 생성 — 목차에서 순서대로 실행하세요`);
  return { ids };
}
