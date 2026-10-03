# PyGrid 워커 내부 헬퍼 — 부트 시 pyodide.globals에 정의된다.
# 사용자 네임스페이스 오염을 막기 위해 모든 이름은 _pygrid_ 접두사를 쓰고,
# 모듈 import는 함수 안에서만 한다(_pygrid_reset이 접두사 없는 전역을 지우기 때문).


def _pygrid_mpl_setup():
    """matplotlib이 사용 가능해진 순간 Agg 백엔드·Pretendard 폰트를 적용한다.

    선택한 방식(작업 지침 §3): matplotlib은 부트 시 선로드하지 않는다(지연 로드).
    이 함수는 멱등이며, 워커가 초기화 스크립트 직후와 매 loadPackagesFromImports
    직후에 호출한다. loadPackagesFromImports는 패키지를 *다운로드*만 하고 import하지는
    않으므로, 다운로드된 상태라면 여기서 직접 import해 사용자 코드가 실행되기 *전에*
    폰트·백엔드를 적용한다 — 안 그러면 첫 matplotlib 실행의 차트 한글이 □로 깨진다.
    """
    import sys

    if "matplotlib" not in sys.modules:
        import importlib.util

        if importlib.util.find_spec("matplotlib") is None:
            return  # 아직 다운로드 전(지연 로드 유지) — 부트·초기화 경로에서 무해
    import matplotlib

    matplotlib.use("Agg")
    try:
        import os

        from matplotlib import font_manager

        _font_path = "/fonts/Pretendard-Regular.otf"
        if os.path.exists(_font_path):
            # OTF 내부 실명으로 등록·지정 (하드코딩 이름 불일치 방지)
            _name = font_manager.FontProperties(fname=_font_path).get_name()
            if not any(f.name == _name for f in font_manager.fontManager.ttflist):
                font_manager.fontManager.addfont(_font_path)
            matplotlib.rcParams["font.family"] = _name
        matplotlib.rcParams["axes.unicode_minus"] = False
    except Exception:
        pass  # 폰트 실패는 실행을 막지 않는다(차트 한글이 □로 보일 뿐)


def _pygrid_exec_capture(code, variable=None):
    """코드 본문을 exec하고 마지막 문장이 표현식이면 eval해 그 값을 돌려준다.

    _pygrid_run(REPL, repr)과 convert.py의 _pygrid_run_convert(블록 실행)가 공유한다.
    variable(출력 선택)을 주면 마지막 표현식 대신 그 이름의 **전역 변수** 값을 돌려준다.
    본문은 항상 끝까지 실행한다(마지막 표현식의 부작용 유지). 없는 이름이면 NameError.
    """
    import ast

    g = globals()
    tree = ast.parse(code, filename="<pygrid>", mode="exec")
    last = None
    if tree.body and isinstance(tree.body[-1], ast.Expr):
        last = tree.body.pop()
    exec(compile(tree, "<pygrid>", "exec"), g)
    value = None
    if last is not None:
        value = eval(compile(ast.Expression(last.value), "<pygrid>", "eval"), g)
    if variable:
        return _pygrid_resolve(variable)
    return value


def _pygrid_resolve(selector):
    """출력 선택 → 값. 변수명이면 전역 조회, 아니면 식으로 평가한다(부록 O.5).

    식 예: `model.rsquared`, `_pygrid_coef_table(model)`, `_pygrid_predict(model, test)`.
    사용자 코드와 같은 전역에서 평가하므로 블록 코드 이상의 권한은 없다.
    """
    g = globals()
    if selector.isidentifier():
        if selector not in g:
            raise NameError(f"출력 변수 '{selector}'가 정의되지 않았습니다")
        return g[selector]
    return eval(compile(selector, "<pygrid>", "eval"), g)


def _pygrid_format_exc(e):
    """래퍼 프레임을 지우고 사용자 코드("<pygrid>")부터 보여주는 트레이스백 문자열."""
    import traceback

    if isinstance(e, SyntaxError):
        return "".join(traceback.format_exception_only(e))
    tb = e.__traceback__
    while tb is not None and tb.tb_frame.f_code.co_filename != "<pygrid>":
        tb = tb.tb_next
    return "".join(traceback.format_exception(e.with_traceback(tb or e.__traceback__)))


def _pygrid_run(code):
    """REPL용 실행: 마지막 표현식의 repr 캡처.

    반환(JSON 문자열):
      성공 {"ok": true, "repr": str|null, "type": str|null}  (None 결과 → repr null)
      실패 {"ok": false, "etype": 예외 클래스명, "msg": str, "tb": 트레이스백}
    """
    import json

    try:
        value = _pygrid_exec_capture(code)
        if value is None:
            return json.dumps({"ok": True, "repr": None, "type": None})
        return json.dumps(
            {"ok": True, "repr": repr(value), "type": type(value).__name__},
            ensure_ascii=False,
        )
    except BaseException as e:  # KeyboardInterrupt(중단)도 실행 실패로 보고한다
        return json.dumps(
            {"ok": False, "etype": type(e).__name__, "msg": str(e), "tb": _pygrid_format_exc(e)},
            ensure_ascii=False,
        )


def _pygrid_inspect():
    """전역 변수 목록(JSON 문자열). 모듈·함수·클래스·언더스코어 이름은 제외."""
    import json
    import types

    out = []
    for name, v in list(globals().items()):
        if name.startswith("_"):
            continue
        if isinstance(
            v, (types.ModuleType, types.FunctionType, types.BuiltinFunctionType, type)
        ):
            continue
        info = {"name": name, "type": type(v).__name__}
        shape = getattr(v, "shape", None)
        if (
            isinstance(shape, tuple)
            and len(shape) == 2
            and all(isinstance(n, int) for n in shape)
        ):
            info["shape"] = list(shape)
        try:
            info["summary"] = repr(v)[:80]
        except Exception:
            info["summary"] = "<repr 실패>"
        # 부록 O.5: 표는 열 이름(목표 열 선택용), 적합된 모델은 시트로 보낼 항목 카탈로그
        cols = getattr(v, "columns", None)
        if info.get("shape") and cols is not None:
            try:
                info["columns"] = [str(c) for c in list(cols)[:500]]
            except Exception:
                pass
        try:
            model = _pygrid_model_info(name, v)
            if model:
                info["model"] = model
        except Exception:
            pass  # 카탈로그 실패가 변수 목록을 막지 않는다
        out.append(info)
    return json.dumps(out, ensure_ascii=False, default=str)


def _pygrid_reset():
    """사용자 전역을 제거한다(워커 안 best-effort 리셋).

    ponytail: sys.modules에 남은 모듈 상태(예: matplotlib rcParams)는 지우지 못한다.
    완전 초기화가 필요하면 client.ts의 terminateAndReboot(워커 재시작) 경로를 쓴다.
    """
    keep = {
        "__name__",
        "__doc__",
        "__package__",
        "__loader__",
        "__spec__",
        "__builtins__",
        "__annotations__",
    }
    g = globals()
    for name in list(g):
        if name in keep or name.startswith("_pygrid"):
            continue
        del g[name]
    # xl()은 공개 이름이라 위에서 지워지므로 복구한다 (xl.py가 _pygrid_xl로도 정의)
    if "_pygrid_xl" in g:
        g["xl"] = g["_pygrid_xl"]
