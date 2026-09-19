# PyGrid Studio — 에이전트 지침

## 1. 프로젝트 개요

브라우저 전용 "Python in Excel" 워크북. 설계서: `docs/pygrid-studio-design.md` (모든 규칙의 마스터).
원칙: **Excel의 Python in Excel과 다르게 동작하는 부분은 반드시 `docs/domain/python-in-excel-parity.md`에 문서화한다.**
구현 계획(마일스톤 M1~M8, 검증된 버전): 설계서 §5.7 + 아래 버전 고정 목록.

## 2. 버전 고정 (2026-09-02 검증)

| 항목 | 값 | 비고 |
|------|----|------|
| Pyodide | **314.0.6** | CDN: `https://cdn.jsdelivr.net/pyodide/v314.0.6/full/` — 모듈 워커 필수(클래식 워커 미지원). jsDelivr는 CORP+ACAO 헤더를 보내 COEP require-corp 호환 |
| Next.js | 15.5.24 (webpack — `--turbopack` 금지) | 16.x 업그레이드 금지 |
| React | 19.1.0 | |
| 그리드 | `@glideapps/glide-data-grid@6.0.4-alpha24` | React 19 peer 지원 버전. import는 `components/grid/SheetGrid.tsx`에서만. 실패 시 폴백: `react-data-grid@7.0.0-beta.61` |
| SheetJS | 0.20.3 (cdn.sheetjs.com tarball) | **npm `xlsx` 설치 금지**(0.18.5 동결) |
| 상태 | zustand 5 + immer + zundo 2.3 | |
| 기타 | idb 8, codemirror 6 + @codemirror/lang-python, vitest 4, @playwright/test 1.62 | `.npmrc save-exact` |

## 3. 워커 프로토콜 계약

`lib/runtime/protocol.ts`가 유일한 계약이며 `/output/runtime-protocol.md`에 문서화한다.
변경 절차: protocol.ts 수정 → runtime-protocol.md 갱신 → grid-ui 측 소비 코드 수정 → vitest/Playwright 통과 확인. **PyProxy를 메인 스레드로 넘기지 않는다** — 워커에서 `convert.py`로 JSON-safe 변환 후 postMessage.

## 4. 변환·spill·순환 참조 정책 (최다 결함 위험 영역)

- Excel ↔ Python 값 변환: 설계서 §3.3 표가 전부다. 행 단위로 vitest 테스트 존재해야 함(`tests/pyodide/`).
- `xl()`은 문자열 리터럴 인수만 허용(ast 정적 분석). 비리터럴 → `#PYTHON! xl() 인수는 문자열 리터럴이어야 합니다`.
- spill 충돌: 다른 블록/비어있지 않은 셀과 겹치면 `#SPILL!`, 이전 spill은 성공 시에만 교체(한 스토어 트랜잭션 = 한 undo 단계).
- 순환 참조: 순환 구성원 전부 `#PYTHON! 순환 참조`, 실행 큐에서 제외.
- 계산 순서: 의존성 우선 위상 정렬, 동순위는 (시트 순, 앵커 행, 열).

## 5. Pyodide 구현 주의점

- 모듈 워커: `new Worker(new URL(...), { type: 'module' })`. Pyodide는 워커 안에서 `import(/* webpackIgnore: true */ indexURL + 'pyodide.mjs')`로 CDN 로드(번들 제외).
- matplotlib: 기본 백엔드가 webagg이므로 초기화 스크립트에서 `matplotlib.use("Agg")` 강제.
- 한글 폰트: 부트 시 `public/fonts/Pretendard-Regular.otf`를 Pyodide FS에 쓰고 `font_manager.fontManager.addfont()` + `rcParams['font.family']` 등록. 안 하면 □ 깨짐.
- 패키지: 부트 시 numpy·pandas만 선로드, 나머지는 `loadPackagesFromImports`.
- COOP/COEP: `next.config.ts` `headers()`(dev/start) + `vercel.json`(배포). 인터럽트는 SAB, 비격리 환경은 워커 terminate+재부트 폴백.

## 6. 브랜드 토큰

`app/globals.css`의 §4.2 토큰이 마스터. Sky Blue `#4A90C2`는 "Python이 관여한 곳"(실행·선택·spill·활성 탭)에만. 폰트: Pretendard(UI)·JetBrains Mono(코드·셀 주소)·Fraunces(로고·제목). 라이트 전용.

## 6.5 예제 코드 카테고리 (부록 N)

예제는 **상위 카테고리 2종**으로 갈라 둔다 — 섞지 않는다.

| 상위 | 하위 | 정본 |
|---|---|---|
| **통계분석** | 통계분석 · 전처리 과정 · 특성공학 · 데이터 분석(회귀 모델) · 모델 평가 | `lib/reference/exampleSnippets.ts` |
| **위험률 산출** | 위험률 산출(조율·평활) · 생명표·계산기수 · 다중탈퇴·다급부 · 보험료·준비금 | 같은 파일 |

- 회귀 모델 범위는 **Ridge·Lasso·ElasticNet + polynomial·log**로 고정. 트리 계열·베이지안 최적화 금지(단위 테스트가 막는다).
- 조각 코드 규칙은 `wrangleSnippets`와 동일 — 마지막 줄은 `print`가 아니라 **값(식)**.
- 자리표시자 `{{range}}`는 **`xl(...)` 호출 전체**로 치환된다 → `df = {{range}}`로 쓰고 따옴표 안에 넣지 않는다.
- 코드 삽입 팝업(`SnippetInsertDialog`)과 샘플 워크북 메뉴(`FileMenu`)가 같은 2단 구분을 쓴다.
- 조각을 고치면 `tests/pyodide/example-snippets.test.ts`가 **전부 실제로 실행**해 확인한다.

## 7. 서브에이전트

- `py-runtime`(.claude/agents/py-runtime/AGENT.md): workers/, lib/runtime/*. 산출물 `/output/runtime-protocol.md`, `/output/conversion-report.json`
- `grid-ui`(.claude/agents/grid-ui/AGENT.md): components/, lib/grid/*, lib/storage/*, lib/io/*
- 서브에이전트 간 직접 호출 금지. 계약은 파일로 전달. 순서: M1∥M3 → M2 → M4 → M5 → M6 → M7 → M8.

## 8. 검증 명령

```bash
npx tsc --noEmit          # 타입 검사
npm run build             # 프로덕션 빌드
npm test                  # vitest 빠른 단위 (tests/unit)
npm run test:py           # Node용 Pyodide 변환 테스트 (tests/pyodide, 느림)
npm run test:e2e          # Playwright (dev 서버 자동 기동, COOP/COEP 포함)
```

골든 테스트 G1~G6(설계서 §5.9): G2=test:py, G3·G4·G5=e2e, G6=e2e 스모크+수동 글리프, G1=수동(실제 Excel 왕복). 결과는 `/output/golden-results.json`.

## 9. 커밋 규칙

- 마일스톤 단위 커밋(Phase 0, M1…M8), 메시지는 `M3: 런타임 부트 — Pyodide 모듈 워커 + 진행률` 형식.
- main 직접 커밋(단독 리포). 커밋 전 `tsc --noEmit` + 해당 마일스톤 테스트 통과 필수.
