# 시트기반 파이썬 (Sheet Python)

> 내부 프로젝트명은 PyGrid Studio이며 파일 확장자(.pygrid.json)·코드 식별자는 호환성을 위해 유지합니다.

브라우저에서 서버 없이 실행되는 **Python in Excel** 스타일 워크북. 표를 붙여넣고, PY 블록에서 `xl("A1:C10", headers=True)` 참조로 시트 데이터를 받아 Pyodide(WASM)로 실행하고, 결과를 spill·객체 카드·이미지로 확인합니다. 셀에는 미니 수식(`=A1+B2*2`, `=SUM(A1:A10)` — 4칙연산·집계 5종)도 쓸 수 있고, 수식 값이 바뀌면 이를 참조하는 Python 블록이 자동 재실행됩니다.

- 설계서: `docs/pygrid-studio-design.md` (모든 규칙의 마스터)
- Excel과 다르게 동작하는 부분: `docs/domain/python-in-excel-parity.md`
- 변환 규칙·경계 케이스: `docs/domain/conversion-rules.md`
- 워커 프로토콜: `output/runtime-protocol.md` · 계산 엔진 계약: `output/calc-engine-api.md`

## 실행

```bash
npm install        # 의존성 (SheetJS는 cdn.sheetjs.com tarball — npm xlsx 설치 금지)
npm run dev        # http://localhost:3000 (COOP/COEP 헤더 포함 — 실행 중단 기능에 필요)
npm run build && npm start   # 프로덕션
```

첫 방문 시 위험률·생명표 샘플 워크북이 열리고, Pyodide(numpy·pandas)는 jsDelivr CDN에서 백그라운드 로드됩니다(약 10초, 재방문은 브라우저 캐시). matplotlib 등은 첫 import 때 지연 로드됩니다.

## 예제 워크북 (데이터 + 코드가 한 파일)

**샘플 워크북** 메뉴의 예제는 데이터가 시트에 이미 들어 있고, 그 데이터를 `xl()`로 읽는 단계별 [설명 마크다운 + 코드 블록]이 함께 저장되어 있습니다 — 파일을 열고 **전체 실행**만 누르면 모델 산출까지 끝납니다(외부 파일·네트워크 의존 없음).

메뉴는 절차가 다른 두 갈래로 나뉘어 있습니다.

### 통계분석 — 표본에서 모형을 고르는 절차

| 예제 | 내장 데이터 | 산출 모델 |
|------|------------|----------|
| 임금 회귀 예측 (5단계) | `wage` 534행 + `meta` 코드북 | 원-핫·표준화 → η² 변수 선별 → **Ridge·Lasso·ElasticNet·다항·로그 회귀** → 교차검증·잔차 진단·임금 밴드 (16단계) |
| 보험료 요인 분석 (GLM) | `policy` 600행 | 기술통계·교차표, 감마 GLM(로그 링크) 계수·상대도, 예측 검증 |
| 빈도·심도 모형 | `claims` 600행 | 심도 분포 AIC 비교, 포아송·음이항 과산포, 순보험료, 몬테카를로 VaR·TVaR |
| 생존분석·유지율 | `experience` 800행 | Kaplan-Meier 생존곡선(자체 구현), 구간별 해지율, 로그순위 검정 |

### 위험률 산출 — 위험률에서 보험료로 내려가는 계리 절차

| 하위 구분 | 예제 |
|---|---|
| 위험률·생명표 | 위험률·생명표(Gompertz·Makeham) · 위험률 산출(원시통계 → 조율 → 보간·평활) |
| 보험료 산출 | 정기보험(계산기수) · 정기보험 변형(체증형·미달체) · 상해공제(직종축) |
| 다중탈퇴·다급부 | 암보험(다중탈퇴) · 암보험(진단 후 생활비) · 종신공제(다급부) · CI종신(3중탈퇴) |
| 준비금·환급금 | 무해지환급형(해지율 PV) · 지급준비금(체인래더) |

원본 데이터(`public/samples/*.xlsx`)를 그리드로 따로 가져오려면 **파일 > 데이터 불러오기 > 샘플 데이터셋**을 쓰세요. 워크북 JSON은 `.claude/skills/sample-workbook-gen/scripts/build_samples.py`가 생성합니다.

## 예제 코드 (코드 삽입 팝업)

PY 블록의 **코드 삽입** 버튼은 같은 2단 카테고리로 조각을 제공합니다 — 그리드에서 범위를 선택하고 넣으면 `{{range}}`가 그 범위의 `xl()` 호출로 치환됩니다.

- **통계분석** — 통계분석 / 전처리 과정 / 특성공학 / 데이터 분석(회귀 모델) / 모델 평가
- **위험률 산출** — 위험률 산출(조율·평활) / 생명표·계산기수 / 다중탈퇴·다급부 / 보험료·준비금
- **핸들링**(pandas 조각 12그룹) · **그래프**(3그룹) · **모델적합 가이드**

회귀 모델은 **Ridge·Lasso·ElasticNet + polynomial·log**로 한정했습니다(계수를 직접 읽어 해석하는 흐름에 집중 — 트리 계열·베이지안 최적화 제외).

## 브라우저 파이썬의 한계

- **실행 지연**: 첫 접속 시 런타임 로드(~10초)와 각 패키지의 첫 import에 시간이 걸립니다(재방문은 캐시로 단축).
- **제한적인 패키지**: Pyodide가 제공하는 패키지와 순수 파이썬 패키지(micropip)만 사용할 수 있습니다 — xgboost·plotly·requests 등은 불가.
- **메모리**: 데이터 크기는 브라우저 탭 메모리 한도에 종속됩니다.
- **격리 환경**: 임의 네트워크 접근·로컬 파일 시스템 접근이 제한됩니다(데이터는 앱의 불러오기 기능 사용).

## AI 설정 (선택)

✦ AI 코드 지원(블록 생성·제안·변수 반영·에러분석)을 쓰려면 **파일 > AI 설정**에서 Anthropic API 키를 입력하세요(발급: [console.anthropic.com](https://console.anthropic.com)). 키는 이 브라우저의 IndexedDB에만 저장되며 Anthropic API 호출에만 사용됩니다 — 서버 라우트가 없고, 워크북 파일·내보내기·git 어디에도 포함되지 않습니다. AI 제안은 자동 실행되지 않으며 항상 확인 후 직접 적용·실행합니다.

## 테스트

```bash
npx tsc --noEmit   # 타입 검사
npm test           # vitest 단위 (a1·모델·클립보드 픽스처 30종·spill·계산 엔진·IO)
npm run test:py    # Node용 Pyodide (§3.3 전 행 + G2 골든 + 예제 조각 40개 + 임금 회귀 16단계 실행)
npm run test:e2e   # Playwright (dev 서버 자동 기동): G3·G4·G5·G6 골든 포함
```

골든 테스트 결과: `output/golden-results.json`

### G1 수동 검증 (실제 Excel 왕복)

자동화 불가 — Windows에서 Excel을 열고 다음을 확인하세요.

1. Excel에서 헤더 + 정수·소수·`1,234`·`12.5%`·`2026-09-02`·한글·빈 셀이 섞인 표(약 21×6)를 복사해 PyGrid 그리드에 붙여넣기 → 붙여넣기 미리보기에서 열 유형 추론 확인 후 적용
2. 같은 범위를 PyGrid에서 복사해 Excel 빈 시트에 붙여넣기
3. 확인: 숫자는 숫자로, `12.5%`는 퍼센트 서식 숫자로, 날짜는 날짜 셀로 인식되고 한글·빈 셀이 보존되는지

## 배포

Vercel: `vercel.json`에 COOP/COEP 헤더가 이미 설정되어 있습니다. `NEXT_PUBLIC_PYODIDE_INDEX_URL`로 Pyodide 셀프 호스팅(`public/pyodide/`) 전환 가능.

## 구조

`CLAUDE.md`(에이전트 지침·버전 고정), `components/`(shell·grid·python·panels), `lib/`(grid·runtime·storage·io), `workers/pyodide.worker.ts`, `data/`(샘플·스니펫), `tests/`(unit·pyodide·e2e).
