// 예제 코드 라이브러리 — 상위 카테고리 2종(통계분석 · 위험률 산출) × 하위 카테고리.
//
// 배경: 코드 삽입 팝업의 기존 목록은 '핸들링'(pandas 조각)과 '그래프'뿐이라, 모델을
// 단계별로 만들어 가는 예제가 들어갈 자리가 없었다. 여기서 두 갈래로 나눈다.
//   · 통계분석  — 통계분석 / 전처리 과정 / 특성공학 / 데이터 분석 / 모델 평가
//   · 위험률 산출 — 위험률 산출 / 생명표·계산기수 / 다중탈퇴·다급부 / 보험료·준비금
// 두 갈래는 성격이 달라 한 목록에 섞지 않는다. 통계분석은 '표본 → 모형 → 평가'의
// 일반 절차이고, 위험률 산출은 '위험률 → 기수 → 급부 현가'의 계리 절차다.
//
// 코드 규칙은 wrangleSnippets와 같다 — 마지막 줄은 print가 아니라 결과 '값(식)'을 둔다.
// 통계분석 조각은 임금 예측 예제(wage 시트: EDUCATION·EXPERIENCE·AGE·WAGE 등 534행)를
// 기준으로 하지만, 상단 상수(TARGET·CAT·NUM)만 바꾸면 다른 표에도 그대로 쓰인다.
// 위험률 산출 조각은 표 없이도 돌도록 소형 인라인 데이터를 품고 있다.
//
// 모델 범위(의도적 제한): 선형회귀 Ridge·Lasso·ElasticNet + 비선형 polynomial·log 모형.
// 트리 계열·베이지안 최적화는 넣지 않는다 — 계수를 직접 읽어 해석하는 흐름에 집중한다.

import type { WrangleSnippet, WrangleSnippetGroup } from "@/lib/reference/wrangleSnippets";

export type { WrangleSnippet as ExampleSnippet };

/** 상위 카테고리 — 하위 카테고리(그룹) 묶음 */
export interface SnippetFamily {
  id: string;
  /** 좌측 목록의 섹션 제목 */
  label: string;
  /** 섹션 설명(툴팁·문서용) */
  desc: string;
  groups: WrangleSnippetGroup[];
}

/** 통계분석 조각 상단 공통 머리말 — 다른 표에 쓸 때 이 3줄만 바꾼다 */
/**
 * 통계분석 조각 공통 머리말(내보냄 — 테스트가 다른 표의 머리말로 갈아 끼운다) — **다른 표로 바꿀 때 고치는 곳은 이 네 줄뿐**이다.
 * NUM+CAT을 화이트리스트로 쓰므로 목록에 없는 열(ID·날짜·파생 등)은 자동으로 빠진다.
 * 조각 안에서 쓰는 다른 상수(비교할 범주 BY, 구간화할 열 COL)는 CAT·NUM에서 끌어 쓴다.
 */
export const STAT_HEAD = `df = {{range}}                               # 그리드에서 표를 선택해 넣으면 xl(...) 호출로 치환됩니다
TARGET = "WAGE"                              # 예측할 연속형 열
CAT = ["SOUTH", "SEX", "UNION", "RACE", "OCCUPATION", "SECTOR", "MARR"]   # 범주형(코드·문자)
NUM = ["EDUCATION", "EXPERIENCE", "AGE"]     # 수치형 설명변수 (여기 없는 열은 모형에서 빠진다)`;

/** 설계행렬 — 원-핫 + 결측 대치. NUM+CAT 화이트리스트라 ID 열이 섞여도 안전하다 */
const DESIGN = `X = pd.get_dummies(df[NUM + CAT], columns=CAT, drop_first=True, dtype=float)
X = X.fillna(X.median())                     # 모델은 NaN을 받지 못한다 — 수치형은 중앙값으로
y = df[TARGET].astype(float)`;

/** 표준 전처리(설계행렬 → 7:3 분할 → train 기준 표준화) — 모델·평가 조각이 공유한다 */
const PREP = `${DESIGN}
X_tr, X_te, y_tr, y_te = train_test_split(X, y, test_size=0.3, random_state=1234)
sc = StandardScaler().fit(X_tr)              # 표준화는 train에서만 fit (정보 누수 방지)
A, B = sc.transform(X_tr), sc.transform(X_te)`;

/** 로그 모형용 — 0·음수가 섞인 타깃도 깨지지 않게 하한을 둔다 */
const LOGY = `ly_tr, ly_te = np.log(y_tr.clip(lower=1e-9)), np.log(y_te.clip(lower=1e-9))`;

// ── 상위 카테고리 ①: 통계분석 ──────────────────────────────────────────────

const STAT_GROUPS: WrangleSnippetGroup[] = [
  {
    id: "stat",
    label: "통계분석",
    snippets: [
      {
        id: "stat-overview",
        label: "데이터 개요 — 형태·타입·결측·중복",
        desc: "행·열 수, 열별 dtype, 결측·중복·0 비율을 한 표로 모아 데이터의 기본 건강 상태를 봅니다.",
        code: `${STAT_HEAD}

# ① 형태와 타입 — 몇 행 몇 열인지, 열마다 어떤 형인지
print("shape:", df.shape)
print("중복 행:", int(df.duplicated().sum()))

# ② 결측·고유값·0 비율을 한 표로
#    고유값이 행 수와 같으면 ID, 0 비율이 높으면 '해당 없음' 플래그일 수 있다
chk = pd.DataFrame({
    "dtype": df.dtypes.astype(str),
    "결측": df.isna().sum(),
    "고유값": df.nunique(),
    "0비율": (df == 0).sum() / len(df),
})
chk["역할"] = np.where(chk.index == TARGET, "타깃",
                np.where(chk.index.isin(CAT), "범주형",
                  np.where(chk.index.isin(NUM), "수치형", "제외(ID 등)")))
chk.reset_index(names="열").round(4)`,
      },
      {
        id: "stat-describe",
        label: "기술통계 — 평균·분산에 왜도·첨도 추가",
        desc: "describe()에 왜도(치우침)·첨도(꼬리 두께)·변동계수를 더해 분포 모양까지 한눈에 봅니다.",
        code: `${STAT_HEAD}

num = df[[TARGET] + NUM].apply(pd.to_numeric, errors="coerce")
d = num.describe().T
d["왜도"] = num.skew()          # 0이면 대칭, +면 오른쪽 꼬리가 길다(고액 소수)
d["첨도"] = num.kurt()          # 0(정규) 대비 꼬리 두께 — 클수록 극단값이 잦다
d["변동계수"] = d["std"] / d["mean"]   # 단위가 다른 열끼리 산포를 비교할 때
d.reset_index(names="열").round(3)`,
      },
      {
        id: "stat-target",
        label: "타깃 분포·정규성 — 원척도 vs 로그",
        desc: "예측 대상의 치우침을 왜도·Shapiro 검정으로 재고, 로그 변환이 도움이 되는지 그림으로 비교합니다.",
        code: `import matplotlib.pyplot as plt
from scipy import stats
${STAT_HEAD}

y = df[TARGET].astype(float).dropna()
ly = np.log(y.clip(lower=1e-9))

# Shapiro-Wilk — p < 0.05면 "정규분포로 보기 어렵다"
p0 = stats.shapiro(y)[1] if len(y) <= 5000 else float("nan")
p1 = stats.shapiro(ly)[1] if len(ly) <= 5000 else float("nan")
print(f"원척도   왜도 {y.skew():7.3f}   Shapiro p = {p0:.3e}")
print(f"로그척도 왜도 {ly.skew():7.3f}   Shapiro p = {p1:.3e}")
print("→ 왜도가 0에 가까워지면 선형회귀의 잔차 정규성 가정에 유리하다")

fig, ax = plt.subplots(1, 2, figsize=(9, 3.4))
ax[0].hist(y, bins=30, color="#4A90C2", edgecolor="white", linewidth=0.4)
ax[0].set_title(f"{TARGET} (원척도)")
ax[1].hist(ly, bins=30, color="#D9A441", edgecolor="white", linewidth=0.4)
ax[1].set_title(f"log({TARGET})")
fig.tight_layout()
fig`,
      },
      {
        id: "stat-corr",
        label: "상관분석 — 상관행렬 히트맵·타깃 상관 순위",
        desc: "수치형 열끼리의 상관계수를 색으로 보고, 타깃과의 상관이 큰 순서와 고상관 쌍을 출력합니다.",
        code: `import matplotlib.pyplot as plt
${STAT_HEAD}

corr = df[[TARGET] + NUM].apply(pd.to_numeric, errors="coerce").corr()

fig, ax = plt.subplots(figsize=(6.4, 5.2))
im = ax.imshow(corr, vmin=-1, vmax=1, cmap="coolwarm")
ax.set_xticks(range(len(corr)), corr.columns, rotation=90, fontsize=8)
ax.set_yticks(range(len(corr)), corr.columns, fontsize=8)
for i in range(len(corr)):
    for j in range(len(corr)):
        ax.text(j, i, f"{corr.iat[i, j]:.2f}", ha="center", va="center", fontsize=6)
fig.colorbar(im, ax=ax, shrink=0.8)
fig.tight_layout()

# 타깃과의 상관 순위 — |r|이 클수록 단독 설명력이 크다(인과는 아님)
print(corr[TARGET].drop(TARGET).abs().sort_values(ascending=False).round(3).to_string())
# 설명변수끼리 |r| > 0.9면 다중공선성 경보 — 규제 회귀(Ridge)가 필요해지는 신호
hi = [(a, b, round(float(corr.at[a, b]), 3)) for a in NUM for b in NUM
      if a < b and abs(corr.at[a, b]) > 0.9]
print("설명변수 간 고상관:", hi if hi else "없음")
fig`,
      },
      {
        id: "stat-group",
        label: "집단 비교 — 그룹 평균·ANOVA/t 검정",
        desc: "범주별 타깃 평균 차이가 우연인지 검정합니다(3집단 이상 ANOVA, 2집단 t검정).",
        code: `from scipy import stats
${STAT_HEAD}
BY = CAT[0]                                  # 비교할 범주 열 (다른 열을 보려면 여기만 교체)

g = df.groupby(BY)[TARGET].agg(건수="size", 평균="mean", 표준편차="std", 중앙값="median")
groups = [v.astype(float).to_numpy() for _, v in df.groupby(BY)[TARGET] if len(v) > 1]

if len(groups) > 2:
    stat, p = stats.f_oneway(*groups)         # 3집단 이상 — 일원분산분석
    name = "ANOVA F"
elif len(groups) == 2:
    stat, p = stats.ttest_ind(*groups, equal_var=False)   # 2집단 — Welch t검정
    name = "t"
else:
    stat, p, name = float("nan"), float("nan"), "검정 불가(집단 1개)"
print(f"{BY}: {name} = {stat:.3f}   p = {p:.3e}")
print("→ p < 0.05면 집단 간 평균 차이를 우연으로 보기 어렵다")
g.reset_index().round(3)`,
      },
    ],
  },
];

const PREP_GROUPS: WrangleSnippetGroup[] = [
  {
    id: "prep",
    label: "전처리 과정",
    snippets: [
      {
        id: "prep-types",
        label: "열 역할 확정 — 타깃·수치형·범주형·제외",
        desc: "dtype만 보면 코드로 저장된 범주형이 숫자로 잡힙니다. NUM·CAT 목록으로 역할을 못박고 나머지는 뺍니다.",
        code: `${STAT_HEAD}

# dtype 기준 자동 분류 — 여기까지는 코드값(1·2·3)도 '숫자'로 잡힌다
auto_num = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
excluded = [c for c in df.columns if c not in NUM + CAT + [TARGET]]
print("dtype이 숫자인 열:", auto_num)
print("모형에서 빠지는 열:", excluded, " ← ID·중복 파생 등")
print("CAT의 값은 '코드'라 평균을 낼 수 없다 — 원-핫으로 펴야 한다")

rows = []
for c in [TARGET] + NUM + CAT:
    vals = df[c].dropna().unique()
    rows.append({
        "열": c,
        "역할": "타깃" if c == TARGET else ("수치형" if c in NUM else "범주형"),
        "dtype": str(df[c].dtype),
        "고유값": len(vals),
        "예시": ", ".join(map(str, sorted(vals, key=str)[:5])),
    })
pd.DataFrame(rows)`,
      },
      {
        id: "prep-missing",
        label: "결측·중복 처리 — 중앙값·최빈값 대치",
        desc: "완전 중복 행을 지우고 수치형은 중앙값, 범주형은 최빈값으로 채웁니다(대치 전후 비교).",
        code: `${STAT_HEAD}

before = len(df)
d = df.drop_duplicates().reset_index(drop=True)      # ① 완전 중복 제거
na_before = d.isna().sum()

num_cols = d.select_dtypes("number").columns
d[num_cols] = d[num_cols].fillna(d[num_cols].median())   # ② 수치형 → 중앙값(이상치에 둔감)
for c in d.columns.difference(num_cols):                 # ③ 그 밖 → 최빈값
    mode = d[c].mode()
    if len(mode):
        d[c] = d[c].fillna(mode.iloc[0])

print(f"중복 제거 {before} → {len(d)}행,  남은 결측 {int(d.isna().sum().sum())}개")
cmp = pd.DataFrame({"대치 전 결측": na_before, "대치 후 결측": d.isna().sum()})
cmp = cmp[cmp.iloc[:, 0] > 0] if (cmp.iloc[:, 0] > 0).any() else cmp
cmp.reset_index(names="열")`,
      },
      {
        id: "prep-outlier",
        label: "이상치 — IQR 경계·winsorize(자르기)",
        desc: "1.5×IQR 밖을 이상치로 세고, 삭제 대신 경계값으로 자른(winsorize) 열을 만듭니다.",
        code: `${STAT_HEAD}
COL = TARGET                                  # 점검할 열 (설명변수를 보려면 NUM[0] 등으로)

s = df[COL].astype(float)
q1, q3 = s.quantile([0.25, 0.75])
iqr = q3 - q1
lo, hi = q1 - 1.5 * iqr, q3 + 1.5 * iqr
out = s[(s < lo) | (s > hi)]
print(f"IQR 경계 [{lo:.4g}, {hi:.4g}]   이상치 {len(out)}건 ({len(out) / len(s):.1%})")

# 표본이 적을 때는 행을 버리는 것보다 경계로 자르는 편이 정보 손실이 적다
w = s.clip(lo, hi)
print(f"→ 자른 뒤 왜도 {w.skew():.3f} (원래 {s.skew():.3f})")
pd.DataFrame({COL: s, COL + "_winsorized": w}).describe().reset_index(names="통계").round(4)`,
      },
      {
        id: "prep-encode",
        label: "범주형 인코딩 — 원-핫(drop_first)",
        desc: "코드값을 그대로 쓰면 없는 순서가 모형에 들어갑니다. 원-핫으로 펼치고 기준범주 하나를 뺍니다.",
        code: `${STAT_HEAD}

# 코드 3이 코드 1의 3배라는 뜻이 아닌데, 숫자로 두면 모형은 그렇게 읽는다.
# 원-핫은 범주마다 0/1 열을 만들고, 기준범주 하나는 절편에 흡수시킨다(drop_first).
${DESIGN}
print(f"설명변수 {len(NUM) + len(CAT)}열 → {X.shape[1]}열 "
      f"(수치형 {len(NUM)} + 더미 {X.shape[1] - len(NUM)})")
print("추가된 더미:", [c for c in X.columns if c not in NUM][:8], "…")
X.head()`,
      },
      {
        id: "prep-split",
        label: "train/test 분할 + 표준화 (누수 방지)",
        desc: "7:3으로 나눈 뒤 표준화를 train에서만 fit합니다. test 통계를 쓰면 성능이 부풀려집니다.",
        code: `from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
${STAT_HEAD}

${PREP}

print(f"train {X_tr.shape}   test {X_te.shape}")
print("train은 평균≈0·표준편차≈1, test는 train 기준이라 정확히 0·1이 아니다 — 정상이다")
pd.DataFrame({
    "변수": X.columns,
    "train평균": A.mean(0).round(6),
    "train표준편차": A.std(0).round(6),
    "test평균": B.mean(0).round(3),
    "test표준편차": B.std(0).round(3),
}).head(12)`,
      },
    ],
  },
  {
    id: "fe",
    label: "특성공학",
    snippets: [
      {
        id: "fe-binning",
        label: "구간화(분위수) — 구간별 타깃 평균",
        desc: "연속 변수를 분위수로 잘라 구간별 타깃 평균이 단조로운지 봅니다(비선형 신호 탐지).",
        code: `${STAT_HEAD}
COL = NUM[0]                                  # 구간화할 연속 변수

d = df[[COL, TARGET]].dropna().copy()
d["구간"], bins = pd.qcut(d[COL], q=4, duplicates="drop", retbins=True)
g = d.groupby("구간", observed=True)[TARGET].agg(건수="size", 평균="mean")
g["전체대비"] = g["평균"] / d[TARGET].mean()
print(f"{COL} 구간 경계:", [round(float(b), 4) for b in bins])
print("→ 구간별 평균이 단조롭게 오르면 선형항으로 충분, 꺾이면 제곱항·구간더미가 필요하다")
g.reset_index().astype({"구간": str}).round(4)`,
      },
      {
        id: "fe-eta",
        label: "설명력 지표 η²(상관비) — 회귀판 IV",
        desc: "분류의 IV(정보값)에 대응하는 회귀 지표. 집단 간 분산 ÷ 전체 분산으로 범주형을 선별합니다.",
        code: `from scipy import stats
${STAT_HEAD}

# IV·WoE는 Target이 0/1인 분류 전용이다. 회귀에서는 상관비 eta^2를 쓴다.
#   eta^2 = 집단 간 제곱합 / 전체 제곱합  (0~1, 클수록 그 범주가 타깃을 잘 가른다)
y = df[TARGET].astype(float)
sst = float(((y - y.mean()) ** 2).sum())
rows = []
for c in CAT:
    grp = [y[df[c] == v].to_numpy() for v in df[c].dropna().unique()]
    grp = [g for g in grp if len(g) > 1]
    if len(grp) < 2:
        continue
    ssb = sum(len(g) * (g.mean() - y.mean()) ** 2 for g in grp)
    F, p = stats.f_oneway(*grp)
    eta2 = ssb / sst
    rows.append({"변수": c, "수준수": len(grp), "eta2": eta2, "F": F, "p": p,
                 "판정": "강함" if eta2 >= 0.10 else "보통" if eta2 >= 0.03 else "약함"})
pd.DataFrame(rows).sort_values("eta2", ascending=False).round(4).reset_index(drop=True)`,
      },
      {
        id: "fe-derive",
        label: "파생변수 — 제곱항·비율·상호작용",
        desc: "수확체감(제곱항), 단위를 맞춘 비율, 두 조건이 겹칠 때만 생기는 효과(상호작용)를 만듭니다.",
        code: `${STAT_HEAD}

a = NUM[0]                                    # 파생에 쓸 수치형 두 개 (필요하면 바꾼다)
b = NUM[1] if len(NUM) > 1 else NUM[0]        # 수치형이 하나뿐이면 교차항 = 제곱항이 된다
g = CAT[0]                                    # 상호작용에 쓸 범주형
top = df[g].mode().iloc[0]                    # 그 범주의 최빈 수준을 0/1 지표로

# 같은 열이 두 번 들어가면 d[a]가 DataFrame이 되어 대입이 깨진다 → 중복 제거
d = df[list(dict.fromkeys([TARGET, a, b, g]))].copy()
d[f"{a}^2"] = d[a] ** 2                       # ① 제곱항 — 오르다 꺾이는 수확체감
d[f"{a}x{b}"] = d[a] * d[b]                   # ② 교차항 — 둘이 함께 클 때만 생기는 효과
d[f"{a}/{b}"] = d[a] / d[b].replace(0, np.nan)  # ③ 비율 — 절대값보다 뜻이 뚜렷할 때
d[f"{g}={top}"] = (d[g] == top).astype(float)
d[f"{g}={top}x{a}"] = d[f"{g}={top}"] * d[a]  # ④ 범주 × 수치 상호작용

made = list(dict.fromkeys([f"{a}^2", f"{a}x{b}", f"{a}/{b}", f"{g}={top}x{a}"]))
base = d[list(dict.fromkeys([a, b]))].corrwith(d[TARGET].astype(float))
print("원본 상관:", base.round(4).to_dict())
print("파생 상관:", d[made].corrwith(d[TARGET].astype(float)).round(4).to_dict())
d[list(dict.fromkeys([a, b, *made, TARGET]))].head(10).round(4)`,
      },
      {
        id: "fe-poly",
        label: "다항 특성 생성 — PolynomialFeatures",
        desc: "수치형만 골라 제곱·교차항까지 펼칩니다. 더미까지 펼치면 열이 폭발하므로 분리합니다.",
        code: `from sklearn.preprocessing import PolynomialFeatures
${STAT_HEAD}

src = df[NUM].apply(pd.to_numeric, errors="coerce")
src = src.fillna(src.median())
pf = PolynomialFeatures(degree=2, include_bias=False)
Z = pd.DataFrame(pf.fit_transform(src), columns=pf.get_feature_names_out(NUM))
print(f"{len(NUM)}열 → {Z.shape[1]}열 (원항 + 제곱 + 교차항)")

r = Z.corrwith(df[TARGET].astype(float)).sort_values(key=abs, ascending=False)
print(r.round(3).to_string())
Z.head().round(4)`,
      },
      {
        id: "fe-transform",
        label: "변환 후보 비교 — log·sqrt·역수 왜도",
        desc: "치우친 변수에 어떤 변환이 가장 대칭에 가까워지는지 왜도로 한 번에 비교합니다.",
        code: `${STAT_HEAD}
COL = TARGET                                  # 변환을 검토할 열

s = df[COL].astype(float).dropna()
pos = s.clip(lower=1e-9)
out = pd.DataFrame({
    "변환": ["원척도", "log", "sqrt", "역수(1/x)"],
    "왜도": [s.skew(), np.log(pos).skew(), np.sqrt(s.clip(lower=0)).skew(), (1 / pos).skew()],
})
out["절대왜도"] = out["왜도"].abs()
print("절대왜도가 가장 작은 변환이 대칭에 가장 가깝다 — 다만 해석 가능성도 함께 본다")
out.sort_values("절대왜도").round(4).reset_index(drop=True)`,
      },
    ],
  },
];

/** 모델 조각 공통 import — 분할·표준화·지표 */
const MIMP = `from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import r2_score, mean_absolute_error, mean_squared_error`;

const MODEL_GROUPS: WrangleSnippetGroup[] = [
  {
    id: "model",
    label: "데이터 분석 (회귀 모델)",
    snippets: [
      {
        id: "model-ols",
        label: "기준선 — 다중선형회귀(OLS)·VIF 진단",
        desc: "규제 없는 OLS로 기준 성능을 잡고, VIF로 다중공선성을 확인해 규제 회귀의 필요를 판단합니다.",
        code: `import statsmodels.api as sm
from statsmodels.stats.outliers_influence import variance_inflation_factor
${STAT_HEAD}

${DESIGN}
res = sm.OLS(y, sm.add_constant(X)).fit()
print(f"R2 = {res.rsquared:.4f}   adj.R2 = {res.rsquared_adj:.4f}   F p = {res.f_pvalue:.3e}")

# VIF — 10을 넘으면 그 변수는 다른 변수들로 거의 설명된다(계수가 불안정해진다).
# 정의상 겹치는 변수(예: 나이 = 학력 + 경력 + 상수)가 있으면 여기서 튄다 → Ridge로 넘어갈 근거.
Xc = sm.add_constant(X).astype(float).to_numpy()
vif = [variance_inflation_factor(Xc, i) for i in range(Xc.shape[1])]
out = pd.DataFrame({"변수": ["const"] + list(X.columns),
                    "계수": res.params.to_numpy(),
                    "표준오차": res.bse.to_numpy(),
                    "p값": res.pvalues.to_numpy(),
                    "VIF": vif})
out["유의"] = np.where(out["p값"] < 0.05, "○", "·")
out["공선성"] = np.where(out["VIF"] > 10, "경보", "")
print("VIF 상위:", out.nlargest(4, "VIF")[["변수", "VIF"]].round(1).to_dict("records"))
out.round(4)`,
      },
      {
        id: "model-ridge",
        label: "Ridge — 계수 경로·RidgeCV로 alpha 선택",
        desc: "L2 규제가 계수를 줄이는 모습을 경로 그림으로 보고, 5겹 교차검증으로 alpha를 고릅니다.",
        code: `import matplotlib.pyplot as plt
from sklearn.linear_model import Ridge, RidgeCV
${MIMP}
${STAT_HEAD}

${PREP}

alphas = np.logspace(-2, 3, 40)
path = np.array([Ridge(alpha=a).fit(A, y_tr).coef_ for a in alphas])
cv = RidgeCV(alphas=alphas, cv=5).fit(A, y_tr)       # 5겹 교차검증으로 alpha 선택
p_tr, p_te = cv.predict(A), cv.predict(B)
print(f"선택 alpha = {cv.alpha_:.4g}")
print(f"R2   train {r2_score(y_tr, p_tr):.4f}   test {r2_score(y_te, p_te):.4f}")
print(f"RMSE train {mean_squared_error(y_tr, p_tr) ** 0.5:.4g}   "
      f"test {mean_squared_error(y_te, p_te) ** 0.5:.4g}")
print("Ridge는 계수를 0으로 만들지 않는다 — 변수를 버리지 않고 크기만 줄인다")

fig, ax = plt.subplots(figsize=(7, 3.6))
for j in range(path.shape[1]):
    ax.plot(alphas, path[:, j], lw=1)
ax.set_xscale("log")
ax.axvline(cv.alpha_, color="#4A90C2", ls="--", lw=1.5)
ax.set_xlabel("alpha (규제 강도)")
ax.set_ylabel("표준화 계수")
ax.set_title("Ridge 계수 경로 — 점선이 교차검증 선택값")
fig.tight_layout()
fig`,
      },
      {
        id: "model-lasso",
        label: "Lasso — 계수를 0으로 만드는 변수 선택",
        desc: "L1 규제는 쓸모가 적은 변수의 계수를 정확히 0으로 만듭니다. 살아남은 변수를 봅니다.",
        code: `from sklearn.linear_model import LassoCV
${MIMP}
${STAT_HEAD}

${PREP}

cv = LassoCV(alphas=np.logspace(-3, 1, 60) * np.std(y_tr), cv=5,
             max_iter=50000, random_state=0).fit(A, y_tr)
p_tr, p_te = cv.predict(A), cv.predict(B)
keep = int((cv.coef_ != 0).sum())
print(f"선택 alpha = {cv.alpha_:.4g}   살아남은 변수 {keep}/{X.shape[1]}개")
print(f"R2   train {r2_score(y_tr, p_tr):.4f}   test {r2_score(y_te, p_te):.4f}")
print(f"RMSE test {mean_squared_error(y_te, p_te) ** 0.5:.4g}")

coef = pd.DataFrame({"변수": X.columns, "계수": cv.coef_})
coef["선택"] = np.where(coef["계수"] != 0, "○", "· (제외)")
coef.reindex(coef["계수"].abs().sort_values(ascending=False).index).round(4).reset_index(drop=True)`,
      },
      {
        id: "model-enet",
        label: "ElasticNet — l1_ratio 격자로 L1·L2 배합",
        desc: "Ridge와 Lasso를 섞습니다. 상관이 높은 변수 묶음을 함께 살리면서 일부는 버립니다.",
        code: `from sklearn.linear_model import ElasticNetCV
${MIMP}
${STAT_HEAD}

${PREP}

cv = ElasticNetCV(l1_ratio=[0.1, 0.3, 0.5, 0.7, 0.9, 0.95, 1.0],
                  alphas=np.logspace(-3, 1, 40) * np.std(y_tr), cv=5,
                  max_iter=50000, random_state=0).fit(A, y_tr)
p_tr, p_te = cv.predict(A), cv.predict(B)
print(f"l1_ratio = {cv.l1_ratio_}  (1이면 Lasso, 0에 가까우면 Ridge)")
print(f"alpha = {cv.alpha_:.4g}   0이 아닌 계수 {int((cv.coef_ != 0).sum())}/{X.shape[1]}개")
print(f"R2   train {r2_score(y_tr, p_tr):.4f}   test {r2_score(y_te, p_te):.4f}")
print(f"RMSE test {mean_squared_error(y_te, p_te) ** 0.5:.4g}")

coef = pd.DataFrame({"변수": X.columns, "계수": cv.coef_})
coef.reindex(coef["계수"].abs().sort_values(ascending=False).index).round(4).reset_index(drop=True).head(12)`,
      },
      {
        id: "model-poly",
        label: "비선형 — 다항회귀(차수 1·2·3 비교)",
        desc: "수치형만 제곱·교차항으로 펼쳐 차수를 올려 봅니다. train만 좋아지는 지점이 과적합입니다.",
        code: `from sklearn.preprocessing import PolynomialFeatures
from sklearn.linear_model import RidgeCV
${MIMP}
${STAT_HEAD}

${PREP}

rows = []
for deg in (1, 2, 3):
    pf = PolynomialFeatures(degree=deg, include_bias=False)
    Ztr = np.hstack([pf.fit_transform(X_tr[NUM]), X_tr.drop(columns=NUM).to_numpy(float)])
    Zte = np.hstack([pf.transform(X_te[NUM]), X_te.drop(columns=NUM).to_numpy(float)])
    s2 = StandardScaler().fit(Ztr)
    m = RidgeCV(alphas=np.logspace(-2, 3, 40), cv=5).fit(s2.transform(Ztr), y_tr)
    q_tr, q_te = m.predict(s2.transform(Ztr)), m.predict(s2.transform(Zte))
    rows.append({"차수": deg, "특성수": Ztr.shape[1],
                 "R2_train": r2_score(y_tr, q_tr), "R2_test": r2_score(y_te, q_te),
                 "RMSE_test": mean_squared_error(y_te, q_te) ** 0.5})
out = pd.DataFrame(rows)
out["과적합폭"] = out["R2_train"] - out["R2_test"]
print("R2_test가 꺾이고 과적합폭이 커지는 직전 차수가 적정 복잡도")
out.round(4)`,
      },
      {
        id: "model-log",
        label: "비선형 — log 모형(반로그·로그-로그)·Duan 스미어링",
        desc: "log(y)를 예측합니다. 반로그와 로그-로그를 함께 적합해 되돌린 값으로 비교하고, 과소추정은 스미어링으로 보정합니다.",
        code: `from sklearn.linear_model import RidgeCV
${MIMP}
${STAT_HEAD}

${PREP}
${LOGY}

# 로그 모형은 두 가지 형태가 있고, 어느 쪽이 맞는지는 데이터가 정한다.
#  ① 반로그(semi-log)  log(y) = b0 + b·z(x)    계수×100 = "x가 1표준편차 늘 때 y 몇 %"
#  ② 로그-로그(log-log) log(y) = b0 + b·log(x)  계수 = 탄력성(%변화 ÷ %변화)
# 설명변수도 자릿수가 넓으면(오른쪽 꼬리가 긴 금액 등) ①은 되돌릴 때 지수적으로 튄다.
WIDE = [c for c in NUM if (df[c] > 0).all() and df[c].skew() > 1]   # 로그를 씌울 수치형
print("로그-로그에 쓸 설명변수:", WIDE if WIDE else "없음(반로그와 같아진다)")

Xl_tr, Xl_te = X_tr.copy(), X_te.copy()
for c in WIDE:
    Xl_tr[c] = np.log(Xl_tr[c].clip(lower=1e-9))
    Xl_te[c] = np.log(Xl_te[c].clip(lower=1e-9))
scl = StandardScaler().fit(Xl_tr)

rows = []
fits = {}
for name, tr, te in [("반로그", A, B),
                     ("로그-로그", scl.transform(Xl_tr), scl.transform(Xl_te))]:
    m = RidgeCV(alphas=np.logspace(-2, 3, 40), cv=5).fit(tr, ly_tr)
    lp_tr, lp_te = m.predict(tr), m.predict(te)
    # Duan 스미어링 — exp()로만 되돌리면 평균이 체계적으로 낮게 나온다(Jensen 부등식).
    #   고치는 것은 '편의'이지 RMSE가 아니다.
    smear = float(np.exp(ly_tr - lp_tr).mean())
    p_te = np.exp(lp_te) * smear
    fits[name] = (m, smear)
    rows.append({"형태": name,
                 "로그척도 R2_test": r2_score(ly_te, lp_te),
                 "원척도 R2_test": r2_score(y_te, p_te),
                 "원척도 RMSE_test": mean_squared_error(y_te, p_te) ** 0.5,
                 "smearing": smear,
                 "예측최대/실제최대": p_te.max() / y_te.max()})

out = pd.DataFrame(rows)
print(out.round(4).to_string(index=False))
best = out.loc[out["원척도 R2_test"].idxmax(), "형태"]
if (out["예측최대/실제최대"] > 2).any():
    print("⚠ 되돌린 예측이 실제 최대의 2배를 넘는 형태가 있다 — 그 형태는 미지정(misspecified)이다")
print(f"→ 원척도 기준으로는 '{best}'가 낫다. 로그척도 R2만 보고 고르면 안 된다 — "
      f"되돌린 값으로 비교해야 한다")

m, smear = fits[best]
unit = "1표준편차당 %" if best == "반로그" else "탄력성(%/%)"
pd.DataFrame({"변수": X.columns, "계수": m.coef_,
              unit: m.coef_ * (100 if best == "반로그" else 1)}) \\
  .pipe(lambda t: t.reindex(t["계수"].abs().sort_values(ascending=False).index)) \\
  .round(4).reset_index(drop=True).head(12)`,
      },
    ],
  },
];

const EVAL_GROUPS: WrangleSnippetGroup[] = [
  {
    id: "eval",
    label: "모델 평가",
    snippets: [
      {
        id: "eval-metrics",
        label: "성능 지표 — R²·RMSE·MAE train/test",
        desc: "한 모델의 train·test 지표를 나란히 놓고 과적합 폭을 봅니다(회귀는 R² 높게·RMSE 낮게).",
        code: `from sklearn.linear_model import RidgeCV
${MIMP}
${STAT_HEAD}

${PREP}

model = RidgeCV(alphas=np.logspace(-2, 3, 40), cv=5).fit(A, y_tr)
p_tr, p_te = model.predict(A), model.predict(B)

def score(t, p):
    return {"R2": r2_score(t, p), "RMSE": mean_squared_error(t, p) ** 0.5,
            "MAE": mean_absolute_error(t, p),
            "MAPE%": float(np.mean(np.abs((t - p) / np.clip(np.abs(t), 1e-9, None))) * 100)}

out = pd.DataFrame({"train": score(y_tr, p_tr), "test": score(y_te, p_te)})
out["차이"] = out["train"] - out["test"]
print("R2 차이가 크면 과적합 — train에만 맞춘 것이다")
out.reset_index(names="지표").round(4)`,
      },
      {
        id: "eval-cv",
        label: "교차검증 — KFold 평균±표준편차",
        desc: "분할 한 번의 운을 걷어냅니다. 5겹 점수의 평균과 표준편차로 성능의 안정성을 봅니다.",
        code: `from sklearn.linear_model import RidgeCV, LassoCV, ElasticNetCV
from sklearn.model_selection import KFold, cross_val_score
from sklearn.pipeline import make_pipeline
${MIMP}
${STAT_HEAD}

${DESIGN}
kf = KFold(n_splits=5, shuffle=True, random_state=1234)
sd = float(np.std(y))

# 표준화를 파이프라인에 넣어야 겹마다 train에서만 fit된다(교차검증 중 누수 방지)
cands = {
    "Ridge": RidgeCV(alphas=np.logspace(-2, 3, 40)),
    "Lasso": LassoCV(alphas=np.logspace(-3, 1, 40) * sd, cv=5, max_iter=50000, random_state=0),
    "ElasticNet": ElasticNetCV(l1_ratio=[0.3, 0.5, 0.7, 0.9, 1.0],
                               alphas=np.logspace(-3, 1, 30) * sd, cv=5,
                               max_iter=50000, random_state=0),
}
rows = []
for name, est in cands.items():
    s = cross_val_score(make_pipeline(StandardScaler(), est), X, y, cv=kf, scoring="r2")
    rows.append({"model": name, "R2_평균": s.mean(), "R2_표준편차": s.std(),
                 "최저": s.min(), "최고": s.max()})
print("표준편차가 크면 데이터를 어떻게 나누느냐에 성능이 휘둘린다는 뜻")
pd.DataFrame(rows).round(4)`,
      },
      {
        id: "eval-compare",
        label: "모델 비교표 — 5종 한 표로 정렬",
        desc: "Ridge·Lasso·ElasticNet·다항·로그 모형을 같은 분할·같은 지표로 비교해 최종 후보를 고릅니다.",
        code: `from sklearn.linear_model import RidgeCV, LassoCV, ElasticNetCV
from sklearn.preprocessing import PolynomialFeatures
${MIMP}
${STAT_HEAD}

${PREP}
${LOGY}
sd = float(np.std(y_tr))

rows = []
def add(name, p_tr, p_te, note=""):
    rows.append({"model": name,
                 "R2_train": r2_score(y_tr, p_tr), "R2_test": r2_score(y_te, p_te),
                 "RMSE_train": mean_squared_error(y_tr, p_tr) ** 0.5,
                 "RMSE_test": mean_squared_error(y_te, p_te) ** 0.5, "비고": note})

for name, est in [("Ridge", RidgeCV(alphas=np.logspace(-2, 3, 40), cv=5)),
                  ("Lasso", LassoCV(alphas=np.logspace(-3, 1, 60) * sd, cv=5,
                                    max_iter=50000, random_state=0)),
                  ("ElasticNet", ElasticNetCV(l1_ratio=[0.3, 0.5, 0.7, 0.9, 1.0],
                                              alphas=np.logspace(-3, 1, 30) * sd, cv=5,
                                              max_iter=50000, random_state=0))]:
    m = est.fit(A, y_tr)
    add(name, m.predict(A), m.predict(B), f"변수 {int((m.coef_ != 0).sum())}/{X.shape[1]}")

# 다항(2차) — 수치형만 펼친다
pf = PolynomialFeatures(degree=2, include_bias=False)
Ztr = np.hstack([pf.fit_transform(X_tr[NUM]), X_tr.drop(columns=NUM).to_numpy(float)])
Zte = np.hstack([pf.transform(X_te[NUM]), X_te.drop(columns=NUM).to_numpy(float)])
s2 = StandardScaler().fit(Ztr)
mp = RidgeCV(alphas=np.logspace(-2, 3, 40), cv=5).fit(s2.transform(Ztr), y_tr)
add("Polynomial(2)", mp.predict(s2.transform(Ztr)), mp.predict(s2.transform(Zte)),
    f"특성 {Ztr.shape[1]}")

# 로그 모형 — 설명변수도 꼬리가 길면 로그-로그로 적합해야 되돌릴 때 튀지 않는다.
# (반로그로 고정하면 금액 같은 변수에서 예측이 지수적으로 폭발한다 — '비선형 log 모형' 조각 참조)
WIDE = [c for c in NUM if (df[c] > 0).all() and df[c].skew() > 1]
Ltr, Lte = X_tr.copy(), X_te.copy()
for c in WIDE:
    Ltr[c] = np.log(Ltr[c].clip(lower=1e-9))
    Lte[c] = np.log(Lte[c].clip(lower=1e-9))
scl = StandardScaler().fit(Ltr)
ml = RidgeCV(alphas=np.logspace(-2, 3, 40), cv=5).fit(scl.transform(Ltr), ly_tr)
sm_ = float(np.exp(ly_tr - ml.predict(scl.transform(Ltr))).mean())
add("Log(Ridge)",
    np.exp(ml.predict(scl.transform(Ltr))) * sm_, np.exp(ml.predict(scl.transform(Lte))) * sm_,
    f"{'로그-로그' if WIDE else '반로그'} · smearing {sm_:.3f}")

out = pd.DataFrame(rows).sort_values("R2_test", ascending=False).reset_index(drop=True)
out["과적합폭"] = out["R2_train"] - out["R2_test"]
print("선택 기준: R2_test 최대 · RMSE_test 최소 · 과적합폭이 작을 것")
out.round(4)`,
      },
      {
        id: "eval-residual",
        label: "잔차 진단 — 적합값 대비·QQ·이분산 검정",
        desc: "잔차가 0 주변에 고르게 퍼지는지, 정규에 가까운지, 분산이 일정한지(BP 검정) 확인합니다.",
        code: `import matplotlib.pyplot as plt
import statsmodels.api as sm
from statsmodels.stats.diagnostic import het_breuschpagan
from scipy import stats
from sklearn.linear_model import RidgeCV
${MIMP}
${STAT_HEAD}

${PREP}

model = RidgeCV(alphas=np.logspace(-2, 3, 40), cv=5).fit(A, y_tr)
fit = model.predict(A)
resid = y_tr.to_numpy() - fit

lm_p = het_breuschpagan(resid, sm.add_constant(A))[1]
print(f"Breusch-Pagan p = {lm_p:.3e}  →  "
      f"{'분산이 일정하지 않다(이분산) — 로그 변환·가중회귀 검토' if lm_p < 0.05 else '등분산 가정 유지'}")
print(f"잔차 왜도 {stats.skew(resid):.3f}   평균 {resid.mean():.3e}")

fig, ax = plt.subplots(1, 3, figsize=(11, 3.3))
ax[0].scatter(fit, resid, s=10, alpha=0.5, color="#4A90C2")
ax[0].axhline(0, color="#C2504A", lw=1)
ax[0].set_xlabel("적합값"); ax[0].set_ylabel("잔차"); ax[0].set_title("잔차 vs 적합값")
stats.probplot(resid, dist="norm", plot=ax[1])
ax[1].set_title("정규 QQ")
ax[2].hist(resid, bins=30, color="#D9A441", edgecolor="white", linewidth=0.4)
ax[2].set_title("잔차 분포")
fig.tight_layout()
fig`,
      },
      {
        id: "eval-coef",
        label: "계수 해석 — 표준화 계수 크기 순 막대",
        desc: "표준화한 계수는 단위가 달라도 크기를 직접 비교할 수 있습니다. 부호와 크기를 함께 읽습니다.",
        code: `import matplotlib.pyplot as plt
from sklearn.linear_model import RidgeCV
${MIMP}
${STAT_HEAD}

${PREP}

model = RidgeCV(alphas=np.logspace(-2, 3, 40), cv=5).fit(A, y_tr)
coef = pd.DataFrame({"변수": X.columns, "계수": model.coef_})
coef = coef.reindex(coef["계수"].abs().sort_values(ascending=False).index).head(12)

fig, ax = plt.subplots(figsize=(7, 4))
colors = np.where(coef["계수"] >= 0, "#4A90C2", "#C2504A")
ax.barh(coef["변수"][::-1], coef["계수"][::-1], color=colors[::-1])
ax.axvline(0, color="#23272E", lw=0.8)
ax.set_xlabel(f"표준화 계수 (1표준편차 증가 시 {TARGET} 변화)")
ax.set_title("변수 중요도 — 크기 순")
fig.tight_layout()

print(f"절편 {model.intercept_:.4g} = 모든 변수가 평균일 때의 {TARGET}")
print(coef.round(4).to_string(index=False))
fig`,
      },
      {
        id: "eval-band",
        label: "예측 밴드 — 분위수로 구간 만들기",
        desc: "최종 모델의 예측값을 분위수로 잘라 실무용 밴드(하·중·상)를 만들고 실적과 대조합니다.",
        code: `from sklearn.linear_model import RidgeCV
${MIMP}
${STAT_HEAD}

${PREP}

model = RidgeCV(alphas=np.logspace(-2, 3, 40), cv=5).fit(A, y_tr)
pred = model.predict(B)
resid_sd = float(np.std(y_tr.to_numpy() - model.predict(A), ddof=1))

band = pd.DataFrame({"실적": y_te.to_numpy(), "예측": pred})
# rank로 자르면 예측값에 동점이 있어도 구간 경계가 겹치지 않는다
band["밴드"] = pd.qcut(band["예측"].rank(method="first"), q=3, labels=["하위", "중위", "상위"])
band["하한"] = band["예측"] - 1.96 * resid_sd     # 근사 95% 예측구간
band["상한"] = band["예측"] + 1.96 * resid_sd
hit = float(((band["실적"] >= band["하한"]) & (band["실적"] <= band["상한"])).mean())
print(f"잔차 표준편차 {resid_sd:.4g}   95% 예측구간 적중률 {hit:.1%} (목표 ≈ 95%)")

g = band.groupby("밴드", observed=True).agg(건수=("실적", "size"), 예측평균=("예측", "mean"),
                                           실적평균=("실적", "mean"), 실적중앙=("실적", "median"))
g["괴리%"] = (g["예측평균"] / g["실적평균"] - 1) * 100
g.reset_index().round(4)`,
      },
    ],
  },
];

// ── 상위 카테고리 ②: 위험률 산출 ───────────────────────────────────────────
//
// 통계분석과 달리 '표본에서 모형을 고르는' 절차가 아니라, 확정된 위험률에서
// 기수 → 급부 현가 → 보험료로 내려가는 계리 절차다. 그래서 목록을 따로 둔다.
// 각 조각은 표 없이도 돌도록 소형 위험률을 인라인으로 만들며, 실제 표가 있으면
// 맨 윗줄을 xl("위험률!A1:C112", headers=True)로 바꾸면 된다.

/** 예제용 위험률 — 실제 경험표가 있으면 이 3줄을 xl() 한 줄로 바꾼다 */
const RATE = `age = np.arange(0, 111)
qx = np.clip(0.00035 + 0.000028 * 1.0955 ** age, 0, 1)   # Makeham 근사 사망률(예제용)
# 실제 경험표를 쓰려면: qx = {{range}}["qx"].to_numpy(float)   (그리드 선택 범위로 치환)`;

const RISK_GROUPS: WrangleSnippetGroup[] = [
  {
    id: "risk-rate",
    label: "위험률 산출 (조율·평활)",
    snippets: [
      {
        id: "risk-crude",
        label: "조율(crude rate) — 발생자수 ÷ 노출",
        desc: "원시통계의 발생자수를 추계인구(노출)로 나눠 군단연령별 조율을 만들고 신뢰도를 함께 봅니다.",
        code: `import pandas as pd, numpy as np

# 군단연령별 원시통계 — 실제로는 xl("원시통계!A1:D20", headers=True)
raw = pd.DataFrame({
    "군단": ["0-4", "5-9", "10-19", "20-29", "30-39", "40-49", "50-59", "60-69", "70-79"],
    "중심연령": [2, 7, 15, 25, 35, 45, 55, 65, 75],
    "발생자수": [12, 9, 21, 68, 194, 613, 1487, 2760, 3312],
    "추계인구": [2_150_000, 2_310_000, 4_820_000, 6_640_000, 7_180_000,
                8_260_000, 8_010_000, 5_430_000, 3_120_000],
})
raw["조율"] = raw["발생자수"] / raw["추계인구"]
# 신뢰도 — 발생자수가 적은 군단은 조율이 크게 흔들린다(포아송 근사 변동계수 = 1/√n)
raw["변동계수"] = 1 / np.sqrt(raw["발생자수"])
raw["신뢰도"] = np.where(raw["발생자수"] >= 1082, "완전", "부분")   # 1082건 = ±5% 90% 신뢰
print("발생자수 1,082건이 흔히 쓰는 완전신뢰도 기준 — 미달 군단은 평활·보정이 더 중요하다")
raw.round(8)`,
      },
      {
        id: "risk-smooth",
        label: "평활 — 직선보간 + 그레빌 9항",
        desc: "군단연령 조율을 한 살 단위로 보간한 뒤 ±4세 9개 값에 대칭 가중치를 곱해 요철을 없앱니다.",
        code: `import matplotlib.pyplot as plt
import pandas as pd, numpy as np

MID = np.array([2, 7, 15, 25, 35, 45, 55, 65, 75])              # 군단 중심연령
CRUDE = np.array([5.6e-6, 3.9e-6, 4.4e-6, 1.02e-5, 2.70e-5,
                  7.42e-5, 1.856e-4, 5.083e-4, 1.0615e-3])      # 군단별 조율
TOP = 110

# 그레빌 9항 보정계수(r = 4 … −4, 합 = 1)와 양끝 외삽계수
GW = np.array([-0.040724, -0.009872, 0.118470, 0.266557, 0.331139,
               0.266557, 0.118470, -0.009872, -0.040724])
EW = np.array([-0.180078, -0.287231, 0.114696, 1.352613])

def greville(y):
    """양끝을 4개씩 외삽해 9칸 창을 채운 뒤 대칭 가중합. 음수는 0으로."""
    head = np.array([])
    for _ in range(4):
        head = np.r_[np.dot(np.r_[head, y][:4], EW[::-1]), head]
    tail = np.array([])
    for _ in range(4):
        tail = np.r_[tail, np.dot(np.r_[y, tail][-4:], EW)]
    ext = np.r_[head, y, tail]
    return np.maximum([np.dot(ext[j:j + 9], GW) for j in range(len(y))], 0.0)

ages = np.arange(TOP + 1)
lin = np.interp(ages, MID, CRUDE)          # 군단 사이는 직선, 양끝은 수평
sm = greville(lin)

fig, ax = plt.subplots(figsize=(7, 3.6))
ax.plot(ages, lin, lw=1, color="#D9A441", label="직선보간")
ax.plot(ages, sm, lw=1.6, color="#4A90C2", label="그레빌 평활")
ax.scatter(MID, CRUDE, s=18, color="#C2504A", zorder=3, label="군단 조율")
ax.set_yscale("log"); ax.set_xlabel("연령"); ax.set_ylabel("위험률"); ax.legend(fontsize=8)
fig.tight_layout()
print(f"가중치 합 {GW.sum():.6f} — 평활은 수준을 유지한 채 요철만 깎는다")
fig`,
      },
      {
        id: "risk-margin",
        label: "적용위험률 — 담보 좁히기·안전할증",
        desc: "전체 발생률에 담보 해당 비율을 곱해 좁히고, 통계 변동을 흡수할 안전할증을 얹습니다.",
        code: `import pandas as pd, numpy as np

ages = np.array([30, 40, 50, 60, 70])
smooth = np.array([2.70e-5, 7.42e-5, 1.856e-4, 5.083e-4, 1.0615e-3])   # 평활 후 전체 발생률
NARROW = np.array([0.31, 0.34, 0.37, 0.35, 0.32])   # 담보 조건 해당 비율(예: BMI 30 이상)
LOAD = 0.50                                          # 안전할증

t = pd.DataFrame({"연령": ages, "평활률": smooth, "담보비율": NARROW})
t["담보률"] = t["평활률"] * t["담보비율"]
t["적용률"] = t["담보률"] * (1 + LOAD)
t["할증분"] = t["적용률"] - t["담보률"]
print(f"안전할증 {LOAD:.0%} — 경험 변동·추세를 흡수하는 마진이며 책임준비금 산출 기초율과는 구분한다")
t.round(8)`,
      },
    ],
  },
  {
    id: "risk-life",
    label: "생명표·계산기수",
    snippets: [
      {
        id: "risk-lx",
        label: "생명표 — lx·dx·px·기대여명 ex",
        desc: "사망률 qx에서 생존자 lx, 사망자 dx, 기대여명 ex를 차례로 만듭니다(기수의 출발점).",
        code: `import pandas as pd, numpy as np

${RATE}

RADIX = 100_000.0
lx = np.empty(len(age)); lx[0] = RADIX
for i in range(1, len(age)):
    lx[i] = lx[i - 1] * (1 - qx[i - 1])              # 생존자 = 전년 생존자 × 생존률
dx = lx * qx                                          # 그 해 사망자

# 기대여명 — 연중 균등사망 가정(a = 0.5)
Lx = lx - 0.5 * dx
Tx = np.cumsum(Lx[::-1])[::-1]
ex = Tx / lx

lt = pd.DataFrame({"x": age, "qx": qx, "lx": lx, "dx": dx, "ex": ex})
print(lt[lt["x"].isin([0, 20, 40, 60, 80, 100])].round(4).to_string(index=False))
lt.round(6)`,
      },
      {
        id: "risk-comm",
        label: "계산기수 — Dx·Nx·Cx·Mx",
        desc: "할인율을 곱한 생존/사망 기수와 그 누계입니다. 보험료·연금·준비금이 전부 이 네 개의 비로 나옵니다.",
        code: `import pandas as pd, numpy as np

${RATE}
I = 0.025                                             # 예정이율

v = 1 / (1 + I)
lx = np.empty(len(age)); lx[0] = 100_000.0
for i in range(1, len(age)):
    lx[i] = lx[i - 1] * (1 - qx[i - 1])
dx = lx * qx

Dx = lx * v ** age                                    # 생존 기수
Cx = dx * v ** (age + 0.5)                            # 사망 기수 (연중 평균 0.5년 후 지급)
Nx = np.cumsum(Dx[::-1])[::-1]                        # Dx의 꼬리 누계 → 연금
Mx = np.cumsum(Cx[::-1])[::-1]                        # Cx의 꼬리 누계 → 사망보험

com = pd.DataFrame({"x": age, "Dx": Dx, "Nx": Nx, "Cx": Cx, "Mx": Mx})
x, n = 40, 20
print(f"{x}세 {n}년만기 정기보험 일시납 순보험료 = (M{x} - M{x + n}) / D{x}"
      f" = {(Mx[x] - Mx[x + n]) / Dx[x]:.6f}")
print(f"{x}세 {n}년납 연납 순보험료 = 위 값 × D{x} / (N{x} - N{x + n})"
      f" = {(Mx[x] - Mx[x + n]) / (Nx[x] - Nx[x + n]):.6f}")
com.round(4)`,
      },
      {
        id: "risk-law",
        label: "사망률 법칙 적합 — Gompertz·Makeham",
        desc: "사력 μx = −ln(1−qx)를 로그 축에서 적합해 표 밖 연령을 외삽하거나 요철을 다듬습니다.",
        code: `import matplotlib.pyplot as plt
from scipy.optimize import curve_fit
import pandas as pd, numpy as np

${RATE}

mu = -np.log(1 - np.clip(qx, 0, 0.999999))            # 사력
m = (age >= 30) & (mu > 0)                            # 유아·사고 구간은 지수법칙에서 벗어난다
x, ly = age[m].astype(float), np.log(mu[m])

b1, b0 = np.polyfit(x, ly, 1)                         # Gompertz는 log μ가 1차식 → 닫힌 해
Bg, cg = float(np.exp(b0)), float(np.exp(b1))
lm = lambda t, A, B, c: np.log(A + B * c ** t)
(Am, Bm, cm), _ = curve_fit(lm, x, ly, p0=[1e-4, Bg, cg],
                            bounds=([0, 1e-12, 1.0], [1.0, 1.0, 2.0]), maxfev=20000)

sse = lambda p: float(np.sum((ly - p) ** 2))
print(f"Gompertz  μ = {Bg:.3e}·{cg:.5f}^x        SSE {sse(np.log(Bg) + x * np.log(cg)):.4f}")
print(f"Makeham   μ = {Am:.3e} + {Bm:.3e}·{cm:.5f}^x   SSE {sse(lm(x, Am, Bm, cm)):.4f}")
print("Makeham의 상수항 A는 연령과 무관한 사고 사망을 담는다")

fig, ax = plt.subplots(figsize=(7, 3.6))
ax.scatter(x, np.exp(ly), s=6, color="#6B7280", label="실측 μx")
ax.plot(x, np.exp(np.log(Bg) + x * np.log(cg)), color="#D9A441", label="Gompertz")
ax.plot(x, np.exp(lm(x, Am, Bm, cm)), color="#4A90C2", label="Makeham")
ax.set_yscale("log"); ax.set_xlabel("연령"); ax.set_ylabel("사력"); ax.legend(fontsize=8)
fig.tight_layout()
fig`,
      },
    ],
  },
];

const RISK_GROUPS2: WrangleSnippetGroup[] = [
  {
    id: "risk-multi",
    label: "다중탈퇴·다급부",
    snippets: [
      {
        id: "risk-multidec",
        label: "다중탈퇴 생존자표 — 사망·발생·해지",
        desc: "탈퇴 원인이 둘 이상이면 lx가 원인별로 갈립니다. 경합을 제거한 결합확률로 생존자를 굴립니다.",
        code: `import pandas as pd, numpy as np

${RATE}
X0, N = 40, 20                                        # 가입연령·보험기간

a = age[X0:X0 + N + 1]
q_d = qx[X0:X0 + N + 1]                               # 사망률
q_i = 0.0012 * 1.075 ** (a - X0)                      # 담보 발생률(예: 진단)
q_w = np.full(len(a), 0.04)                           # 해지율

# 두 위험이 같은 해에 함께 있을 때: q = 1 - (1-q1)(1-q2)
# 단순 합은 두 위험이 겹치는 부분을 두 번 센다
q_all = 1 - (1 - q_d) * (1 - q_i) * (1 - q_w)
lx = np.empty(len(a)); lx[0] = 100_000.0
for i in range(1, len(a)):
    lx[i] = lx[i - 1] * (1 - q_all[i - 1])

# 원인별 탈퇴자 — 결합 탈퇴를 원인별 강도 비율로 배분
inten = np.c_[q_d, q_i, q_w]
share = inten / inten.sum(axis=1, keepdims=True)
dec = lx[:, None] * q_all[:, None] * share
t = pd.DataFrame({"x": a, "lx": lx, "d_사망": dec[:, 0],
                  "d_발생": dec[:, 1], "d_해지": dec[:, 2]})
simple = float((q_d + q_i + q_w)[0])
print(f"단순 합 {simple:.6f} vs 결합 {q_all[0]:.6f} — "
      f"차이 {(simple - q_all[0]) * 1e5:.1f}건/10만명 (겹침을 두 번 센 만큼)")
t.round(3)`,
      },
      {
        id: "risk-defer",
        label: "면책기간·감액 — 90일 3/4·50% 감액",
        desc: "첫 해에만 걸리는 면책기간과 가입 초기 감액을 기수에 반영하는 방법입니다.",
        code: `import pandas as pd, numpy as np

${RATE}
X0, N, I = 40, 20, 0.025

a = age[X0:X0 + N + 1]
q = 0.0012 * 1.075 ** (a - X0)                        # 담보 발생률
v = 1 / (1 + I)

# 1) 90일 면책 — 첫 해(k=0)는 365일 중 275일만 보장 → 계수 3/4
# 2) 가입 1년 미만 50% 감액 → 계수 1/2
k1 = np.ones(len(a)); k1[0] = 0.75
k2 = np.ones(len(a)); k2[0] = 0.50
adj = q * k1 * k2

lx = np.empty(len(a)); lx[0] = 100_000.0
for i in range(1, len(a)):
    lx[i] = lx[i - 1] * (1 - q[i - 1])
Dx = lx * v ** (a - X0)
Cx = lx * adj * v ** (a - X0 + 0.5)                   # 감액·면책을 반영한 발생 기수
Mx = np.cumsum(Cx[::-1])[::-1]

raw_Mx = np.cumsum((lx * q * v ** (a - X0 + 0.5))[::-1])[::-1]
base = raw_Mx[0] / Dx[0]
print(f"보정 전 급부 현가 {base:.6f}  ->  보정 후 {Mx[0] / Dx[0]:.6f}  "
      f"({Mx[0] / Dx[0] / base - 1:.2%})")
pd.DataFrame({"x": a, "q": q, "면책계수": k1, "감액계수": k2,
              "보정q": adj, "lx": lx, "Cx": Cx, "Mx": Mx}).round(6)`,
      },
      {
        id: "risk-sumx",
        label: "급부배율 SUMx — 급부 여러 개를 한 현가로",
        desc: "담보마다 기수를 따로 쌓고 급부금액(배율)을 곱해 더합니다. 보험료는 이 합 하나에서 나옵니다.",
        code: `import pandas as pd, numpy as np

${RATE}
X0, N, I = 40, 20, 0.025

a = age[X0:X0 + N + 1]
v = 1 / (1 + I)
q_d = qx[X0:X0 + N + 1]
lx = np.empty(len(a)); lx[0] = 100_000.0
for i in range(1, len(a)):
    lx[i] = lx[i - 1] * (1 - q_d[i - 1])
Dx = lx * v ** (a - X0)
disc = v ** (a - X0 + 0.5)

# 담보별 발생률과 급부배율(가입금액 대비)
BEN = {
    "사망": (q_d, 1.00),
    "암진단": (0.0022 * 1.070 ** (a - X0), 1.00),
    "수술": (0.0041 * 1.045 ** (a - X0), 0.20),
    "입원일당": (0.0180 * 1.030 ** (a - X0), 0.02),
}
rows, total = [], 0.0
for name, (q, mult) in BEN.items():
    Msum = float(np.sum(lx * q * disc))               # 그 담보의 Cx 합
    pv = Msum / Dx[0] * mult
    total += pv
    rows.append({"담보": name, "배율": mult, "기수합": Msum, "급부현가": pv})

t = pd.DataFrame(rows)
t["구성비"] = t["급부현가"] / total
print(f"SUMx (전체 급부 현가) = {total:.6f}")
print("보험료는 담보별로 따로 구하지 않는다 — 이 합 하나를 납입 연금 현가로 나눈다")
t.round(6)`,
      },
    ],
  },
  {
    id: "risk-prem",
    label: "보험료·준비금",
    snippets: [
      {
        id: "risk-net",
        label: "순보험료 — 일시납·연납·월납 N(m) 보정",
        desc: "급부 현가를 납입 연금 현가로 나눕니다. 월납은 연중 분할 납입을 N(m)으로 보정합니다.",
        code: `import pandas as pd, numpy as np

${RATE}
X0, N, I = 40, 20, 0.025
FACE = 100_000_000                                    # 가입금액

v = 1 / (1 + I)
lx = np.empty(len(age)); lx[0] = 100_000.0
for i in range(1, len(age)):
    lx[i] = lx[i - 1] * (1 - qx[i - 1])
Dx = lx * v ** age
Cx = lx * qx * v ** (age + 0.5)
Nx, Mx = np.cumsum(Dx[::-1])[::-1], np.cumsum(Cx[::-1])[::-1]

pv_ben = (Mx[X0] - Mx[X0 + N]) / Dx[X0]               # 급부 현가
ann = (Nx[X0] - Nx[X0 + N]) / Dx[X0]                  # 연납 기시급 연금 현가
rows = []
for m, label in [(1, "연납"), (2, "6개월납"), (4, "3개월납"), (12, "월납")]:
    # N(m) 근사: N - (m-1)/(2m)·D — 연중 분할 납입의 이자 손실 보정
    ann_m = ann - (m - 1) / (2 * m) * (Dx[X0] - Dx[X0 + N]) / Dx[X0]
    P = pv_ben / ann_m / m
    rows.append({"납입주기": label, "m": m, "연금현가": ann_m,
                 "1회 순보험료": P * FACE, "연간합계": P * m * FACE})
print(f"일시납 순보험료 = {pv_ben * FACE:,.0f}원   (급부현가 {pv_ben:.6f})")
pd.DataFrame(rows).round(2)`,
      },
      {
        id: "risk-gross",
        label: "영업보험료 — 사업비 alpha·beta·gamma 부가",
        desc: "신계약비·유지비·수금비를 얹어 영업보험료를 풉니다(보험료 비례 항은 분모로 이항).",
        code: `import pandas as pd, numpy as np

${RATE}
X0, N, I, M = 40, 20, 0.025, 12
FACE = 100_000_000

v = 1 / (1 + I)
lx = np.empty(len(age)); lx[0] = 100_000.0
for i in range(1, len(age)):
    lx[i] = lx[i - 1] * (1 - qx[i - 1])
Dx = lx * v ** age
Cx = lx * qx * v ** (age + 0.5)
Nx, Mx = np.cumsum(Dx[::-1])[::-1], np.cumsum(Cx[::-1])[::-1]

pv_ben = (Mx[X0] - Mx[X0 + N]) / Dx[X0]
ann = (Nx[X0] - Nx[X0 + N]) / Dx[X0] - (M - 1) / (2 * M) * (Dx[X0] - Dx[X0 + N]) / Dx[X0]
ann10 = (Nx[X0] - Nx[X0 + min(N, 10)]) / Dx[X0]

A1, A2 = 0.009, 0.90      # 신계약비: 가입금액 비례 / 순보험료 비례(10년 환산)
B1, B2 = 0.0015, 0.05     # 유지비: 가입금액 비례 / 영업보험료 비례
G = 0.025                 # 수금비(영업보험료 비례)

P_net = pv_ben / ann
# 영업보험료 G: G·ann·(1 - B2 - G) = 급부현가 + A1·ann10 + B1·ann + A2·P_net·ann10
num = pv_ben + A1 * ann10 + B1 * ann + A2 * P_net * ann10
P_gross = num / ann / (1 - B2 - G)
t = pd.DataFrame([
    {"항목": "순보험료", "연액": P_net * FACE, "월액": P_net * FACE / M},
    {"항목": "영업보험료", "연액": P_gross * FACE, "월액": P_gross * FACE / M},
    {"항목": "부가보험료", "연액": (P_gross - P_net) * FACE, "월액": (P_gross - P_net) * FACE / M},
])
t["비중"] = t["연액"] / (P_gross * FACE)
print(f"부가보험료율 = {P_gross / P_net - 1:.2%}  (영업보험료 비례 항은 분모로 이항해 푼다)")
t.round(0)`,
      },
      {
        id: "risk-reserve",
        label: "책임준비금·해약환급금 — 해지공제 상각",
        desc: "장래법 준비금(장래 급부 - 장래 순보험료)에서 미상각 신계약비를 뺀 환급금을 경과별로 냅니다.",
        code: `import matplotlib.pyplot as plt
import pandas as pd, numpy as np

${RATE}
X0, N, I = 40, 20, 0.025
FACE, SURR_YEARS = 100_000_000, 7                     # 해지공제 상각기간

v = 1 / (1 + I)
lx = np.empty(len(age)); lx[0] = 100_000.0
for i in range(1, len(age)):
    lx[i] = lx[i - 1] * (1 - qx[i - 1])
Dx = lx * v ** age
Cx = lx * qx * v ** (age + 0.5)
Nx, Mx = np.cumsum(Dx[::-1])[::-1], np.cumsum(Cx[::-1])[::-1]
P = (Mx[X0] - Mx[X0 + N]) / (Nx[X0] - Nx[X0 + N])     # 연납 순보험료(가입금액 1당)

alpha = 0.009 * min(N, 10)                            # 해지공제 재원(신계약비)
rows = []
for t in range(N + 1):
    x = X0 + t
    V = (Mx[x] - Mx[X0 + N]) / Dx[x] - P * (Nx[x] - Nx[X0 + N]) / Dx[x]   # 장래법
    unamort = alpha * max(0, SURR_YEARS - t) / SURR_YEARS                 # 미상각 신계약비
    ded = min(unamort, max(V, 0.0))
    rows.append({"경과": t, "책임준비금": V * FACE, "해지공제": ded * FACE,
                 "해약환급금": max(V - ded, 0.0) * FACE, "납입보험료": P * t * FACE})
res = pd.DataFrame(rows)
res["환급률"] = res["해약환급금"] / res["납입보험료"].replace(0, np.nan)

fig, ax = plt.subplots(figsize=(7, 3.6))
ax.plot(res["경과"], res["책임준비금"], color="#4A90C2", label="책임준비금")
ax.plot(res["경과"], res["해약환급금"], color="#D9A441", label="해약환급금")
ax.plot(res["경과"], res["납입보험료"], color="#6B7280", ls="--", lw=1, label="납입보험료")
ax.set_xlabel("경과년수"); ax.set_ylabel("원"); ax.legend(fontsize=8)
fig.tight_layout()
print(f"해지공제는 {SURR_YEARS}년 직선 상각 — 그 뒤로는 준비금 = 환급금")
fig`,
      },
      {
        id: "risk-ldf",
        label: "지급준비금 — 체인래더 LDF·IBNR",
        desc: "누적 지급 삼각형에서 개발계수를 뽑아 삼각형을 채우고 미보고발생손해액(IBNR)을 냅니다.",
        code: `import pandas as pd, numpy as np

# 누적 지급 삼각형 — 실제로는 xl("triangle!A1:I9", headers=True)
tri = np.array([
    [3580, 5240, 6010, 6390, 6570, 6660, 6700, 6720],
    [3810, 5580, 6400, 6810, 7000, 7100, 7150, np.nan],
    [4020, 5900, 6780, 7210, 7420, 7530, np.nan, np.nan],
    [4310, 6320, 7260, 7720, 7940, np.nan, np.nan, np.nan],
    [4560, 6690, 7690, 8170, np.nan, np.nan, np.nan, np.nan],
    [4880, 7160, 8220, np.nan, np.nan, np.nan, np.nan, np.nan],
    [5140, 7540, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan],
    [5470, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan],
], dtype=float)
n = tri.shape[0]

# 개발계수 LDF = 합(다음 열) / 합(같은 행이 다음 열까지 있는 현재 열) — 가중평균
ldf = np.array([np.nansum(tri[:n - j - 1, j + 1]) / np.nansum(tri[:n - j - 1, j])
                for j in range(n - 1)])
full = tri.copy()
for i in range(n):
    for j in range(n - 1):
        if np.isnan(full[i, j + 1]):
            full[i, j + 1] = full[i, j] * ldf[j]

ult = full[:, -1]
paid = np.array([tri[i, n - 1 - i] for i in range(n)])   # 사고연도별 최신 누적 지급
out = pd.DataFrame({"사고연도": np.arange(1, n + 1), "기지급": paid,
                    "최종손해액": ult, "IBNR": ult - paid})
print("개발계수:", np.round(ldf, 4).tolist())
print(f"IBNR 합계 = {out['IBNR'].sum():,.0f}")
out.round(1)`,
      },
    ],
  },
];

// ── 상위 카테고리 등록 ─────────────────────────────────────────────────────

export const EXAMPLE_SNIPPET_FAMILIES: SnippetFamily[] = [
  {
    id: "stat",
    label: "통계분석",
    desc: "표본에서 모형을 고르는 절차 — 통계분석 → 전처리 → 특성공학 → 회귀 모델 → 평가",
    groups: [...STAT_GROUPS, ...PREP_GROUPS, ...MODEL_GROUPS, ...EVAL_GROUPS],
  },
  {
    id: "risk",
    label: "위험률 산출",
    desc: "확정된 위험률에서 내려가는 계리 절차 — 위험률 → 기수 → 급부 현가 → 보험료·준비금",
    groups: [...RISK_GROUPS, ...RISK_GROUPS2],
  },
];

/** id로 예제 조각 찾기 (테스트·문서 생성용) */
export function findExampleSnippet(id: string): WrangleSnippet | undefined {
  for (const fam of EXAMPLE_SNIPPET_FAMILIES) {
    for (const g of fam.groups) {
      const s = g.snippets.find((x) => x.id === id);
      if (s) return s;
    }
  }
  return undefined;
}
