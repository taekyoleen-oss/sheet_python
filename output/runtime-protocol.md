# 런타임 워커 프로토콜 (M4)

계약 원본: `lib/runtime/protocol.ts` — 이 문서는 그 해설이다. protocol.ts가 바뀌면 이 문서를 함께 갱신한다.
소비 측(grid-ui)은 워커를 직접 다루지 말고 **`lib/runtime/client.ts`의 `getRuntimeClient()`**를 쓴다.

## 구성 요소

| 파일 | 역할 |
|------|------|
| `workers/pyodide.worker.ts` | 모듈 워커. CDN에서 Pyodide 로드, 스냅샷 주입, 코드 실행+변환, stdout/stderr 스트리밍 |
| `lib/runtime/client.ts` | 메인 스레드 클라이언트: 부트, 준비 전 큐잉, 요청/응답 매핑, 타임아웃·인터럽트·재부트 |
| `lib/runtime/py/bootstrap.py` | 워커 내부 헬퍼(`_pygrid_exec_capture`·`_pygrid_run`·`_pygrid_inspect`·`_pygrid_reset`·`_pygrid_mpl_setup`) |
| `lib/runtime/py/xl.py` | `xl()` 브리지: 참조 추출(ast) + 스냅샷 캐시 → §3.3 입력 변환(DataFrame/스칼라) |
| `lib/runtime/py/model_out.py` | 부록 O.5 모델 결과 헬퍼: 카탈로그(`_pygrid_model_info`)·계수표·적합 통계·예측(구간)·평가 지표. 부트 시 convert.py 다음에 로드 |
| `lib/runtime/py/convert.py` | §3.3 출력 변환: 값 모드 `cells`, 객체 모드 preview/PNG. 단일 출력 `_pygrid_run_convert`, 다중 출력 `_pygrid_run_convert_multi` |
| `lib/runtime/converters.ts` | `OutCell[][]` → `Cell[][]`(`toCells`), 앵커+shape → `spillRange`. 순수 함수 |
| `lib/runtime/py/init_default.py` | 기본 초기화 스크립트. `client.ts`가 `DEFAULT_INIT_SCRIPT`로 재수출 |

원칙: **PyProxy는 postMessage로 넘기지 않는다.** 모든 페이로드는 JSON-safe(이미지 등 ArrayBuffer는 transferable).

## 부트 시퀀스

1. `RuntimeClient.boot({ initScript?, indexURL?, timeoutSec? })` — 멱등(중복 호출 시 같은 Promise).
   - indexURL 기본값: `NEXT_PUBLIC_PYODIDE_INDEX_URL` env → 없으면 `DEFAULT_PYODIDE_INDEX_URL`(jsDelivr v314.0.6).
2. 워커에 `{t:'boot', indexURL, packages:['numpy','pandas'], initScript, fontUrl:'/fonts/Pretendard-Regular.otf'}` 전송.
   `crossOriginIsolated`면 곧바로 `{t:'setInterruptBuffer', buffer: SharedArrayBuffer(1)}`도 전송(워커가 부트 완료 후 적용).
3. 워커 진행 단계 — `{t:'progress', pct, label}` (pct는 대략치):
   - 5% 스크립트 로드 → 20% 런타임 초기화 → 45% 패키지 로드(numpy·pandas) → 75% 폰트 등록 → 90% 초기화 스크립트 → 100% 완료
4. 폰트: `fontUrl`을 fetch해 Pyodide FS `/fonts/Pretendard-Regular.otf`에 기록. 실패해도 부트는 계속(stderr id 0으로 경고).
5. 초기화 스크립트 실행. **실패해도 부트는 성공** — 오류는 `{t:'stderr', id:0, ...}`로 스트리밍.
6. `{t:'ready', pyVersion, pyodideVersion}` → 클라이언트 상태 `ready`, 큐 flush.
   실패 시 `{t:'bootError', message}` → 상태 `error`, 대기 중 요청 전부 reject.

### matplotlib 지연 로드 정책 (선택한 구현)

- 부트 시 matplotlib을 내려받지 않는다. 초기화 스크립트의 `import matplotlib`은 `try/except ImportError`로 감싸져 있다.
- 워커는 **모든 run/repl 전에** `loadPackagesFromImports(code)`를 돌리고, 직후 `_pygrid_mpl_setup()`(멱등)을 호출한다.
  `loadPackagesFromImports`는 다운로드만 하므로, `_pygrid_mpl_setup`이 다운로드된 matplotlib을 **직접 import**해
  사용자 코드 실행 전에 Agg 백엔드 + Pretendard 폰트(OTF 내부 실명으로 등록)를 적용한다 — 첫 실행부터 차트 한글이 렌더된다.
- 초기화 스크립트에는 `loadPackagesFromImports`를 적용하지 않는다(guarded import를 부트 시 내려받지 않기 위함).
  **제약**: 커스텀 초기화 스크립트가 scipy 등 미로드 패키지를 import하면 ImportError가 stderr(id 0)로 보고되고 부트는 계속된다. — TODO(후속): 초기화 스크립트용 패키지 선로드 옵션.

## 준비 전 큐잉

`run`/`repl`/`analyze`/`inspect`/`reset`은 ready 전에 호출해도 되고, 클라이언트가 큐에 넣었다가 ready 때 순서대로 보낸다(스펙 §4.7 "런타임 준비 후 실행됩니다"). 워커 내부도 메시지를 프라미스 체인으로 **직렬 처리**한다 — 동시에 두 실행이 인터리빙되지 않는다.

## 메시지 목록

### 메인 → 워커 (`MainToWorker`)

| t | 필드 | 설명 |
|---|------|------|
| `boot` | indexURL, packages, initScript, fontUrl | 부트. 워커당 1회(중복 무시) |
| `setInterruptBuffer` | buffer: SharedArrayBuffer | 인터럽트 버퍼. 체인 밖에서 즉시 적용(부트 전이면 보관) |
| `analyze` | id, code | `xl()` 참조 추출(ast). 리터럴 위반 → `analyzeError` |
| `run` | id, blockId, code, snapshots, outputMode, includeIndex, **output?**, **outputs?** | 블록 실행: 스냅샷 주입 → 실행 → 출력 선택 → §3.3 변환. `outputs`가 있으면 다중 출력 경로(아래) |
| `repl` | id, code | 콘솔 실행(공유 네임스페이스) |
| `inspect` | id | 전역 변수 목록 |
| `listDir` | id, path? | 부록 P 속성 창 탐색기: 폴더 목록 (path 생략 = Python 작업 폴더) → `dirListing` / `fileError` |
| `resetRuntime` | id, initScript | 워커 안 best-effort 리셋(사용자 전역 삭제 → 초기화 스크립트 재실행) |
| `writeFile` | id, path, bytes: ArrayBuffer | Pyodide FS에 파일 기록(`FS.writeFile`). bytes는 **transferable**. path 검증은 클라이언트 담당 |
| `readFile` | id, path | FS 파일 읽기 → `fileRead` / 실패 시 `fileError` |

### 워커 → 메인 (`WorkerToMain`)

| t | 필드 | 설명 |
|---|------|------|
| `progress` | pct, label | 부트 진행률 |
| `ready` | pyVersion, pyodideVersion | 부트 완료 |
| `bootError` | message | 부트 실패 |
| `stdout` / `stderr` | id, chunk | 라인 단위 스트리밍. id = 진행 중인 run/repl id, **유휴(부트·리셋 출력)는 id 0** |
| `analyzed` | id, refs | `xl()` 문자열 리터럴 인수를 그대로(중복 제거·순서 유지) |
| `analyzeError` | id, message | `xl() 인수는 문자열 리터럴이어야 합니다` / `xl() headers 인수는 True/False 리터럴이어야 합니다` / `구문 오류: …` |
| `result` | id, blockId + RunPayload | run 결과(아래 "run 변환 결과" 참조). 실패: errorType(예외 클래스명)·message·traceback·durationMs. stdout/stderr = 그 run이 스트리밍한 출력 전체 사본(카드 아래 결과 표시용) |
| `replResult` | id, repr, stdout:'', stderr:'', traceback? | repr = 마지막 표현식의 `repr()`, None이면 null. stdout/stderr 필드는 빈 문자열(이미 스트리밍됨) |
| `dirListing` | id, cwd, path, entries: DirEntry[] | `{name, dir, size}` — 폴더 먼저·이름순. path는 워커가 절대 경로로 정규화 |
| `variables` | id, vars: VariableInfo[] | name·type·shape(2D만)·summary(repr ≤80자). 모듈·함수·클래스·`_` 이름 제외. 부록 O.5: 2D 표는 `columns`(≤500), 적합된 statsmodels 결과·scikit-learn 추정기(비지도 포함)는 `model: ModelInfo`(kind·formula·intervals·predict·proba·featureNames·members[{group,label,expr,code,preview,shape}] — 모델에 실제로 있는 항목만, expr=출력 식, code=라이브러리 원래 코드) |
| `resetDone` | id | 리셋 완료 |
| `fileWritten` | id | writeFile 성공 |
| `fileRead` | id, bytes: ArrayBuffer | readFile 성공. bytes는 **transferable**(Emscripten readFile의 새 버퍼) |
| `fileError` | id, message | writeFile/readFile 실패. `파일을 찾을 수 없습니다: …` / `파일 쓰기 실패(…): …` / 준비 전 호출 |

## 실행 의미론 (run/repl 공통)

0. **엑셀 엔진 지연 설치** (R5): 코드가 `EXCEL_CODE_RE`(protocol.ts — `read_excel|to_excel|ExcelWriter|\.xlsx|\.xls\b`)에
   걸리면 세션당 1회 `loadPackage("openpyxl")`, 실패 시 micropip(PyPI) 폴백(pyimport로 실행 — 사용자 전역 오염 없음).
   **314.0.6 배포판에는 openpyxl이 없어 실경로는 micropip 폴백이다** — `tests/pyodide/excel.test.ts`가 체인 전체를 검증.
   프라미스 캐시라 이후 실행은 대기 없이 통과. 설치 실패는 stderr 한 줄
   (`엑셀 지원(openpyxl) 설치 실패 — …`)로만 알리고 실행은 계속된다(그때 사용자 코드가 원래의 ImportError를 낸다).
1. `loadPackagesFromImports(code)` (오류 무시 — 구문 오류는 실행 단계에서 더 나은 트레이스백으로 보고)
2. `_pygrid_mpl_setup()` (멱등)
3. **run만**: `msg.snapshots`(JSON)를 `_pygrid_xl_load`로 워커 내부 캐시에 주입. 실행이 끝나면(finally) 캐시를 비운다 — 스냅샷은 요청 단위로만 유효하다.
4. 본문을 exec하고 **마지막 문장이 표현식이면 eval해 값 캡처**. repl은 repr(None → null), run은 §3.3 변환.
   예외 시 워커 래퍼 프레임을 지운 트레이스백(`<pygrid>` 프레임부터) + 예외 클래스명(errorType).
   `KeyboardInterrupt`(중단)도 같은 실패 경로로 보고된다 → errorType `"KeyboardInterrupt"`.
5. 제약: 콘솔/블록 코드의 **top-level await 미지원**(ast exec/eval 캡처 방식) — 필요해지면 `PyCF_ALLOW_TOP_LEVEL_AWAIT`로 확장.

## xl() 스냅샷 주입 계약 (calc engine → run)

- `run.snapshots`: `Record<참조 문자열, RangeSnapshot>` — 키는 `analyze`가 돌려준 refs **그대로**(예: `"A1:F21"`, `"Sheet2!A1"`). 코드 안 `xl("...")` 리터럴과 문자 단위로 일치해야 한다.
- `RangeSnapshot.values`는 항상 2D(단일 셀도 1×1), 빈 셀은 `null`. `scalar:true`면 `xl()`이 스칼라를 돌려준다.
- 주입 안 된 참조를 실행 중 `xl()`이 만나면 `RuntimeError: xl(): 참조 '…'의 데이터가 준비되지 않았습니다` (안전망 — calc engine이 항상 선주입해야 한다).
- `xl(ref, headers=False)` 입력 변환(§3.3): 열 전체 `'n'` 정수·빈 셀 없음 → `int64`, `'n'` 그 외 → `float64`+NaN, `'d'` → `datetime64[ns]`+NaT, `'b'` 빈 셀 없음 → `bool`, 그 외/혼합/빈 셀 포함 → `object`+None. `headers=True`면 첫 행이 컬럼명(빈 헤더는 `Unnamed: {j}`), `False`면 정수 컬럼 0..n-1. 경계 케이스 결정: `docs/domain/conversion-rules.md`.

## 출력 선택 계약 (`run.output`, v1.1)

`OutputSelection = { variable?, columns?, rowLimit? }` (`types/workbook.ts`). 필드가 없으면 **기존 동작 그대로**(마지막 표현식, 전체 열·행).
워커는 `JSON.stringify(msg.output ?? null)`을 `_pygrid_output_sel`로 넣고 `_pygrid_run_convert(code, mode, includeIndex, output_sel)`을 호출한다 — PyProxy는 넘기지 않는다.

| 필드 | 의미 | 경계 |
|------|------|------|
| `variable` | 마지막 표현식 대신 **그 이름의 전역 변수**를 결과로 쓴다. 식별자가 아니면 **출력 식**으로 전역에서 `eval`한다(부록 O.5 — 예: `model.rsquared`, `_pygrid_coef_table(model)`, `_pygrid_predict(model, test, interval="prediction")`; 판정·평가는 `_pygrid_resolve`). 식 오류는 그 출력만의 실패(예외 클래스명 그대로). 본문은 끝까지 실행된다(마지막 표현식의 부작용 유지) | 없는 이름 → RunFailure `errorType:"NameError"`, message `출력 변수 '…'가 정의되지 않았습니다` (한국어 요약은 `errors-ko.ts`가 붙인다) |
| `columns` | DataFrame 결과에서 **요청한 순서대로** 열만 남긴다. 열 이름은 `str()` 기준 비교 | 없는 열은 조용히 무시. 요청 열이 **하나도 없으면 전체 열 유지**(빈 spill 방지). DataFrame이 아니면 무시 |
| `rowLimit` | DataFrame·Series·list/tuple·1D·2D ndarray에서 **상위 N 데이터 행**만 남긴다 | 헤더 행(DataFrame 열 이름·Series name)은 N에 포함되지 않는다. `N ≤ 0`은 무제한. 스칼라·dict·Figure·기타 객체는 무시 |

적용 시점: **§3.3 변환 *전***(`convert.py`의 `_pygrid_select_output`). 따라서 값 모드 `cells`/`shape`와 객체 모드 `preview`/`shape`/`typeName`이 같은 선택을 반영한다.

## 다중 출력 계약 (`run.outputs` → `RunSuccess.outputs`, v1.2)

한 블록의 코드를 **1회만 실행**하고 결과의 여러 부분을 서로 다른 셀 위치에 배치하기 위한 경로다.
`msg.outputs`가 있으면 `outputMode`/`includeIndex`/`output`(레거시 단일 출력 필드)은 **무시된다**.

```ts
OutputRequest = { id, mode, includeIndex, selection? }   // selection = OutputSelection (위 표와 동일 의미)
OutputItemSuccess = { id, ok: true, kind, cells?, typeName?, shape?, preview?, imagePng? }
OutputItemFailure = { id, ok: false, errorType, message, traceback? }
RunSuccess.outputs?: OutputItem[]                        // 요청 순서 그대로
```

워커 동작: `JSON.stringify(msg.outputs)`를 `_pygrid_outputs`로 넣고
`_pygrid_run_convert_multi(code, outputs_json)` 호출 → 본문 1회 exec(마지막 표현식 캡처) →
요청마다 `variable` 선택 → `_pygrid_select_output` → `_pygrid_convert`. PNG는 출력마다 base64로 받아
`ArrayBuffer`로 디코드하고 **전부 transferable로 넘긴다**. `RunSuccess.kind`는 첫 성공 출력의 kind(없으면 `"object"`) —
다중 출력에서 의미 있는 값은 `outputs[]`뿐이므로 최상위 `cells`/`preview`/`imagePng`는 채우지 않는다.

### 실패 격리 (conversion-rules #20~#23)

| 무엇이 실패 | 결과 |
|---|---|
| 코드 본문(예외·`KeyboardInterrupt`) | **run 전체 실패** `RunFailure` — `outputs` 없음. grid-ui는 블록의 모든 출력에 같은 오류를 표시 |
| 출력의 `selection.variable`이 없는 이름 | 그 출력만 `OutputItemFailure` `NameError` — 나머지 출력은 정상 배치, `payload.ok`는 여전히 `true` |
| 값 모드 출력에 Figure | 그 출력만 `PyGridImageError` `이미지는 값으로 펼칠 수 없습니다` |
| 변환 자체의 예상 밖 실패 | 그 출력만 `ConversionError` |

`columns`/`rowLimit`/`includeIndex`는 출력마다 독립이다 — 같은 변수를 값 모드 상위 3행과 객체 모드 카드로 동시에 낼 수 있다.

## run 변환 결과 (§3.3 출력)

- **값 모드**(`outputMode:'values'`): `cells: OutCell[][]` + `shape` + `kind`(1×1이면 `'scalar'` 아니면 `'table'`).
  - 스칼라→1×1(None/NaN/NaT/inf → `{v:null,t:'s'}` 빈 셀), bool→`'b'`, 날짜→`'d'` ISO+`f` 서식, list/1D ndarray→세로 1열, Series→name 있으면 헤더 셀+값, 2D→2D, dict→2열(키·값), DataFrame→헤더 행+값(includeIndex `'auto'`: 기본 RangeIndex 제외·그 외 첫 열, `'always'`/`'never'` 강제), 그 외 객체→`str()` 1셀.
  - Figure는 값 모드 불가 → RunFailure `errorType:"PyGridImageError"`, message `이미지는 값으로 펼칠 수 없습니다` (grid-ui는 `#PYTHON!` 셀로 표시).
- **객체 모드**(`'object'`): `kind` 분류 — DataFrame/Series→`'table'`(preview: columns·dtypes·상위 100행·shape, NaN/NaT→null), Figure→`'image'`(`imagePng: ArrayBuffer` transferable, PNG dpi150), 스칼라류→`'scalar'`(preview repr), 그 외→`'object'`(preview repr, 2000자 절단).
- 변환 자체가 예상 밖으로 실패하면 `errorType:"ConversionError"`.

## 파일 I/O (R4) — `writeFile` / `readFile` / `listFiles`

- `client.writeFile(path, bytes: Uint8Array): Promise<void>` — Pyodide FS(작업 디렉터리)에 파일 기록.
  `client.readFile(path): Promise<Uint8Array>` — FS 파일 읽기(**파이썬 코드가 만든 파일도 읽는다**).
  둘 다 다른 요청처럼 **준비 전 큐잉**된다.
- **path는 경로 없는 파일 이름만**: `/`·`\`·`..`·빈 문자열은 클라이언트가 즉시 reject
  (`파일 이름은 경로 없이 지정해야 합니다('/'·'..' 금지): …`) — 워커로 나가지 않는다.
- **재부트 재기록**: `terminateAndReboot()`는 워커 FS를 날린다. 클라이언트가 `writeFile`한 파일의
  사본을 `Map<path, Uint8Array>`로 캐시해 두고 **매 ready 때 큐 flush보다 먼저** 전부 다시 쓴다
  (재기록 응답은 추적하지 않는 best-effort) — 인터럽트·타임아웃 재부트 후에도 데이터가 살아남는다.
  진행 중이던 writeFile/readFile **요청 자체**는 재부트 시 다른 요청처럼 reject된다.
- `client.listFiles(): string[]` — 캐시된(클라이언트가 쓴) 파일 이름 목록. 워커 왕복 없음(UI 칩 표시용).
  파이썬이 만든 파일은 목록에 없다. **`_`로 시작하는 이름은 내부용**(모델 적합 등) — UI에서 걸러도 된다.
- 전송: writeFile의 bytes(클라이언트 사본에서 뜬 복사본)와 fileRead의 bytes 모두 ArrayBuffer **transferable**.

### 모델 적합 어댑터 (R5, `lib/reference/fit-runner.ts`)

`runDistributionFit(payload, onPhase)` — FIT_SCRIPT(`lib/reference/pyFit.ts`, 단일 원본)를 워커 런타임에서 실행:
`boot`(멱등 부트) → `pkg`(`import scipy; del scipy` repl — 워커의 loadPackagesFromImports가 scipy 다운로드) →
`run`(`_fit_input.json`·`_fit_script.py` writeFile → `exec(compile(open(...).read()), {빈 dict})` repl → `_fit_output.json` readFile → JSON.parse).
FIT_SCRIPT 꼬리만 소스와 다르다: `json.dumps(OUT)` 반환 대신 `_fit_output.json`에 json.dump(repl repr 파싱은 취약).
빈 dict exec라 사용자 전역에 이름이 남지 않는다. 검증: `tests/pyodide/fit.test.ts`(`npm run test:py`).

## 타임아웃 · 인터럽트 · 재부트

- run/repl마다 타임아웃(기본 60초, `timeoutSec` 인자·`defaultTimeoutSec`으로 조정). 만료 시 `client.interrupt()`.
- `interrupt()`: `crossOriginIsolated`면 SAB[0]=2(SIGINT) → 실행 중 코드에 `KeyboardInterrupt`.
  **2초 안에 결과가 없거나 비격리 환경이면 `terminateAndReboot()`** — 워커 종료 후 같은 initScript로 재부트.
- `terminateAndReboot()`: 대기 중 요청 전부 `"런타임이 재설정되어 실행이 중단되었습니다"`로 reject, 상태 `rebooting` → `loading` → `ready`, `'reboot'` 이벤트 발생(UI 토스트: "런타임이 재설정되어 변수가 초기화되었습니다").
- `reset()`(워커 안 리셋)은 변수만 지운다. `sys.modules`에 남은 모듈 상태까지 지우려면 `terminateAndReboot()`.

## 클라이언트 API 요약 (`RuntimeClient`)

```ts
getRuntimeClient(): RuntimeClient            // 앱 전역 싱글턴 (브라우저 전용)
boot(opts?): Promise<void>                   // 멱등
run(blockId, code, snapshots?, outputMode?, includeIndex?, timeoutSec?, output?, outputs?): Promise<RunPayload>
repl(code, timeoutSec?): Promise<{ repr: string | null; traceback?: string }>
analyze(code): Promise<string[]>             // 비리터럴 인수 → reject(Error(한국어 메시지))
inspect(): Promise<VariableInfo[]>
reset(initScript?): Promise<void>            // 새 스크립트로 리셋 가능. 주면 이후 재부트에도 그 스크립트 사용
writeFile(path, bytes: Uint8Array): Promise<void>  // FS 기록 + 캐시(재부트 후 자동 재기록)
readFile(path): Promise<Uint8Array>          // FS 읽기. 없으면 reject(한국어 메시지)
listFiles(): string[]                        // 캐시된 파일 이름(워커 왕복 없음). '_' 접두 = 내부용
interrupt(): void
terminateAndReboot(): Promise<void>
on(event, fn): () => void                    // 'progress' | 'status' | 'stdout' | 'stderr' | 'reboot'
getStatus(): 'idle'|'loading'|'ready'|'running'|'error'|'rebooting'
getVersions(): { pyVersion, pyodideVersion } | null
DEFAULT_INIT_SCRIPT                          // init_default.py 원문 (UI 기본값)
```

grid-ui 소비 헬퍼(`lib/runtime/converters.ts`): `toCells(cells)` → `Cell[][]`(v:null=빈 셀), `spillRange(anchor, rows, cols)` → `CellRange`.

검증: §3.3 행 단위 테스트 `tests/pyodide/convert.test.ts`(`npm run test:py`) → `output/conversion-report.json`. M5(calc engine)가 snapshots 직렬화·spill 반영을 잇는다.
