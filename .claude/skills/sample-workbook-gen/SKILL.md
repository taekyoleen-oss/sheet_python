---
name: sample-workbook-gen
description: 샘플 워크북 JSON(생명표·손해율·히스토그램 예제)과 코드 스니펫(data/snippets.json) 생성. 샘플·스니펫을 추가·수정할 때 사용.
---

# sample-workbook-gen

## 실행

```bash
python .claude/skills/sample-workbook-gen/scripts/build_samples.py
```

산출물 — 부록 K 계리 예제 5종은 `public/samples/*.xlsx` 원본을 시트로 내장한다(openpyxl 없으면 zipfile+xml 파서로 폴백):

- `life-table.pygrid.json` — 위험률·생명표(mortality_table 101행): lx·dx·ex, Gompertz·Makeham 적합. **첫 방문 기본 로드 워크북**.
- `premium-glm.pygrid.json` — 보험료 요인 분석(policy 600행): 교차표, 감마 GLM(로그 링크) 상대도, 예측 검증.
- `freq-severity.pygrid.json` — 빈도·심도(claims 600행): 심도 AIC 비교, 과산포 검정, 순보험료, MC VaR·TVaR.
- `survival-retention.pygrid.json` — 생존분석·유지율(experience 800행): Kaplan-Meier 자체 구현, 구간 해지율, 로그순위.
- `chain-ladder.pygrid.json` — 지급준비금(triangle 8×8): 개발계수·완성 삼각형·IBNR·Mack 표준오차.
- `premium-term.pygrid.json` — 보험료 산출·정기보험(부록 M): 계산기수 → 보험료(원본 검산) → 준비금·해약환급금 → 표준/적용 비교.
  원본은 `PREMIUM_SRC`(OneDrive 산출과정표 xlsx) — 없으면 생성을 건너뛰고 기존 JSON을 유지한다.
- `cancer-multi.pygrid.json` — 암보험 다중탈퇴(부록 M.4 #6): 위험률 11종(발생률 vs 비율) → 담보별 다중탈퇴 생존자표
  → Cx·Mx(90일 면책 3/4) → 급부배율 SUMX → 순·영업공제료 → 원본 `총괄` 대조 검산.
  원본은 `CANCER_SRC`(OneDrive 더블암진단특약.xlsx) — 없으면 생성을 건너뛰고 기존 JSON을 유지한다.
- `risk-rate.pygrid.json` — 위험률 산출(부록 M.4 #5): 발생자수÷추계인구 = 조율 → 비만동반 비율·안전할증 0.5
  → 직선보간 + 그레빌 9항 평활(양끝 4개 외삽) → 원본 `3. 산출결과` 111행 대조 검산 → 연령별 곡선.
  원본은 `RISK_SRC`(OneDrive 영양조사.xlsx) — 없으면 생성을 건너뛰고 기존 JSON을 유지한다.
- `nonsurrender.pygrid.json` — 무해지환급형(부록 M.4 #8): 치매 CDR 발생률 → 해지율 4% 이중탈퇴 생존자표
  → 현금흐름 PV(Dx·Cx·Nx·Mx 대체, 중증치매 연금 10년) → 순·영업보험료(12) → 무해지 vs 표준형 환급률 곡선
  → 원본 `P테이블` 10개 조합 대조 검산. 원본은 `NS_SRC`(OneDrive 치매보험 PV산출.xlsm — zip 구조라 xlsx와 같은 파서로 읽힌다).
- `whole-life-multi.pygrid.json` — 종신공제 다급부(부록 M.5 #9): 사망률·50%이상 장해율 → **이중탈퇴 생존자표
  `lx`(사망만) vs `lx′`(사망+장해 납입면제)** → 급부 5종의 기수 차분(Mx·Dx·Nx) → 급부배율 SUMX → N* 납입주기 보정
  → 순·영업공제료(단기납·일시납 사업비 세트) → 원본 `총괄` 28항목 대조 검산.
  원본은 `WL_SRC`(OneDrive 더블종신공제_주계약_엑셀.xlsx) — 없으면 생성을 건너뛰고 기존 JSON을 유지한다.
- `term-variants.pygrid.json` — 정기보험 변형(부록 M.5 #7): 평준형 요약 → **체증형**(`Rx = ΣMx`로 10년 거치 후
  체증 급부의 `M*` 구성, 원본 `P!A7`) → **미달체**(사망률 ×3 — 원본 `위험률!O·P`, 표준율 ×3은 `Q·R`) →
  3종 비교표 → 준비금·환급률 3종 곡선 → 원본 `P`·`(표준)P`·`미달P`·`미달표준P`·`V`·`미달V` 28항목 대조 검산.
  원본은 `PREMIUM_SRC`(premium-term과 같은 산출과정표) — 없으면 생성을 건너뛰고 기존 JSON을 유지한다.
  원본 `조회!C3`(체증)이 0%로 저장돼 있어 체증형은 원본 캐시 대신 `Rx` 항등식으로 검증한다(`TV_INC` = 예제 기본 체증률).
- `accident-class.pygrid.json` — 상해공제 직종축(부록 M.5 #10): **위험률의 축이 나이가 아니라 직종 A/B**
  (원본 `OFFSET(위험율!$B$13, 주계약!$C$4, sex)` — 나이 인수가 없다) → 경과 0~5년 기수표
  (`lx`와 80%장해 탈퇴가 하나 더 붙은 `lx(80)` 두 벌) → 급부 4종 배율(재해소득보장은 10년 확정연금 현가)
  → **순수보장형**(1년·RA 0·α2 17%·β2 8.5%) / **만기환급형 50%**(5년·α2 50%·β2 5.5% — 만기급부가
  보험료에 비례해 `RA·m·mm·Dx(n)/N*` 항이 **분모로** 이항된다, 원본 `주계약!I3·I5`) → 8칸 매트릭스
  → 원본 `PV테이블` 공제료 블록(`AD3:AX742`) 주계약 8행 대조 검산(영업·순·만기환급금 3열 모두 차이 0).
  원본은 `ACC_SRC`(OneDrive 무)Sh어업인상해공제1601_가입설계.xlsm) — 없으면 생성을 건너뛰고 기존 JSON을 유지한다.
  원본 `보험료` 시트(A1:AC2569)는 암 상품 템플릿 잔재(코드 `A7514`·`E2214`…, 나이 61~80)라 검산에 쓰지 않는다.
  `PTABLE.zip`·`VTABLE.zip`은 Fasoo DRM 암호화라 읽을 수 없다.
- `cancer-annuity.pygrid.json` — 암 진단 후 생활비 연금(부록 M.6 #11): 암발생률 q1 → **주요암 qc**
  (`N7 = G7-G7*(J7+K7+L7+M7)`, 소액암 4종은 발생률이 아니라 **비율**) → **암발생후사망률 2차원표**
  (진단연령 × 진단 후 경과 0~10, 남 C6:W101 / 여 C106:W203) → 진단 후 코호트 `lsck(x)+k`
  → 확정지급분(3배 일시금 + 2년 확정 분기 생활비, 계수 `CL14=3+0.03·ä(4)·ä2`)과
  생존조건 지급분(10년 반기 연금 `BE=SUMPRODUCT(AI:AR, D7:D16)·v^(t+0.5)`) 분리
  → 급부 A~F 조립(`k1=3/4×1/2` 90일 면책×50% 감액, `k2=1/2`) → 순·영업공제료(`N(#)x` 월납 보정)
  → 원본 45항목 대조 검산(차이 ≤ 1.2e-8) → 연령별 공제료 곡선·급부 구성비.
  원본은 `CA_SRC`(OneDrive `보험료산출(생활비주는암)-주계약(올인원).xlsx`) — 없으면 생성을 건너뛴다.
  원본 `기수표(생활비암)`의 lx는 **일반사망 탈퇴가 없는 상수 10만**이고(Z8=Z7), 준비금 블록(CT6:DF104)은
  급부 현가로 `Ax(일반암)`만 써서 `Sumx`와 맞지 않아(경과 0년 −15,221원) 준비금 단계는 재현하지 않는다.
  같은 계열의 `MG올인원암공제_주계약엑셀(161012)-1.xlsx`는 갱신형·건강관리형이 섞여 더 복잡해 쓰지 않았다.
- `ci-whole-life.pygrid.json` — CI종신 3중탈퇴(부록 M.6 #12): CI 11종(`w1`~`w10` + `b` 중대한 화상·부식)
  **경합 제거 결합 `w = 1−∏(1−wʲ)` · `q^CI = w+b−w·b`**(k=0은 90일 면책 3/4) → LTC 결합
  `q^HD = q^CI+q^LT−q^CI·q^LT` → **3중탈퇴 생존자표 `L`·`L¹`·`L²`(납입면제)·`L³`(1년미만 유방암)**
  → 급부 기수 `M̄`·`M̄1`·`M̄2`·`M̄3`·`C̄4`·`M̄5` → `SUMx` → `N(m)*`·`N(m)#` → `10Px`(순환참조 닫힌 해)
  → 월납 순·영업보험료 → **방법서 수식 ↔ 코드 대조·자체 정합성 검증 29행** → 생존자 곡선·탈퇴 원인별 기여도.
  원본은 `CI_SRC`(OneDrive 교보 CI종신 산출방법서 PDF)지만 **PDF는 산식만 싣고 발생률 수치가 없어**
  (전부 보험개발원 문서번호 인용) 생성에 쓰지 않는다 — **원본 파일 없이도 생성된다.**
  발생률 출처: LTC 3종(`kLx`·`kDx`·`rx`)만 같은 폴더 `위험률(CI)세부내역.xls`의 실데이터를 모듈 상수
  `CI_LTC`(20~79세 남/여)로 내장했고, 사망률·암발생률은 `premium-term`·`cancer-multi` 워크북에서 읽어
  쓰며(`sample_rate_cols()`), 나머지 CI 9종 + `bx`는 **사망률 × 고정계수 가정**이다.
  그래서 검산이 아니라 **자체 정합성 검증**을 쓴다 — 단 `kL′ = (1−rx)×kLx`만은 원본 xls에 결과 열이
  따로 있어 실제 대조가 가능하다(최대차 4.9e-07 = 원본 6자리 반올림 수준).
  스킵 조건은 `premium-term`·`cancer-multi` JSON이 없을 때뿐이라 **`main()`에서 그 둘보다 뒤에 와야 한다.**
- `wage-regression.pygrid.json` — 임금 회귀 예측(부록 N, **통계분석 카테고리**): 데이터 품질 → 타깃 분포·로그 변환
  → 상관·다중공선성(AGE = EDUCATION + EXPERIENCE + {2,6}) → 범주형 ANOVA → 코드북 기준 열 구분
  → IQR 이상치 → **원-핫(10열 → 16열)·7:3 분할·train 전용 표준화** → 분위수 구간화 → **상관비 eta^2**(회귀판 IV)
  → 파생변수·PolynomialFeatures → OLS·VIF → **Ridge·Lasso·ElasticNet 계수 비교** → **다항(차수 1·2·3)·로그 모형
  + Duan 스미어링** → 모델 비교표·5겹 교차검증 → 잔차 진단(Breusch-Pagan) → 계수 해석·임금 밴드. 16단계.
  원본 데이터는 `public/samples/wage.xlsx`(CPS 1985 임금 534행 + `meta` 코드북 시트) — **리포에 있어 항상 생성된다.**
  `read_xlsx(name, sheet)`로 두 시트를 모두 읽으므로 시트 인자를 지우면 안 된다.
  원본 노트북(패스트캠퍼스 Ch03)과 다르게 한 곳은 부록 N.5 표에 있다 — 특히 **코드 범주형을 숫자로 쓰지 않는다**.
  13단계는 스미어링 보정이 train에서는 편의·RMSE를 함께 줄이지만 **이 분할의 test에서는 RMSE를 악화**시키는
  실측값을 그대로 싣는다(부록 N.6). 테스트가 이 관계를 고정하므로 숫자를 바꾸려면 테스트도 같이 고친다.
- `loss-ratio.pygrid.json` · `claim-severity.pygrid.json` — 소형 기본 예제(부록 H.2).
- `data/snippets.json` — 초보자용 스니펫(기술통계·그룹 집계·피벗·히스토그램·선형회귀·생명표 lx 계산).
  코드 삽입 팝업의 **예제 코드 카테고리**는 이 스크립트가 아니라 `lib/reference/exampleSnippets.ts`가 갖는다(부록 N).

## 규칙

- 워크북 JSON은 설계서 §3.1 스키마(`types/workbook.ts`)를 정확히 따른다. 스키마 변경 시 이 스크립트를 같이 고친다.
- 블록 코드는 `xl()` 참조가 실제 시트 범위와 일치해야 한다(로드 직후 전체 실행이 성공해야 함).
- 스니펫의 `{{range}}` 자리표시자는 삽입 시 현재 선택 범위의 `xl()` 참조로 치환된다.
- 사용 라이브러리는 Pyodide 가용 범위(numpy·pandas·scipy·statsmodels·matplotlib)로 한정 — lifelines 등은 자체 구현.
- 워크북 하나당 2MB 이하. 데이터 셀에는 서식(`f`)을 붙이지 않는다.
- 생성 후 검증: `npm test -- tests/unit/sample-workbooks.test.ts` + `npm run test:e2e -- tests/e2e/sample-workbooks.spec.ts`.
- 임금 회귀 예제는 `npx vitest run -c vitest.pyodide.config.ts tests/pyodide/wage-regression.test.ts`로
  16단계를 실제 시트 데이터에 돌려 본다(블록 하나하나가 혼자서도 도는지까지 확인한다).
