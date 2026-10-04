# xl() 브리지 — 설계서 §2.4(참조 추출)·§3.3(입력 변환).
# 워커가 부트/리셋 시 bootstrap.py 다음에 로드한다.
# 공개 이름은 `sheet`(정식)와 `xl`(Python in Excel 호환 별칭). 나머지는 _pygrid_ 접두사(리셋에서 살아남는다).

# ref 문자열 → RangeSnapshot dict {"values": 2D, "types": 2D, "scalar": bool}
# 워커가 run 직전에 _pygrid_xl_load로 주입하고 finally에서 clear한다.
_pygrid_xl_cache = {}


def _pygrid_xl_load(snapshots_json):
    import json

    _pygrid_xl_cache.clear()
    _pygrid_xl_cache.update(json.loads(snapshots_json))


def _pygrid_extract_refs(code):
    """ast로 xl(...) 호출의 참조를 추출한다(§2.4).

    반환(JSON 문자열):
      {"ok": true, "refs": [문자열 리터럴 그대로, 중복 제거·순서 유지]}
      {"ok": false, "message": 한국어 오류}  — 비리터럴 인수 등
    """
    import ast
    import json

    try:
        tree = ast.parse(code)
    except SyntaxError as e:
        return json.dumps({"ok": False, "message": f"구문 오류: {e.msg}"}, ensure_ascii=False)

    refs = []
    for node in ast.walk(tree):
        if not (isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id in ("sheet", "xl")):
            continue
        if not node.args or not (
            isinstance(node.args[0], ast.Constant) and isinstance(node.args[0].value, str)
        ):
            return json.dumps(
                {"ok": False, "message": f"{node.func.id}() 인수는 문자열 리터럴이어야 합니다"},
                ensure_ascii=False,
            )
        for kw in node.keywords:
            if kw.arg == "headers" and not (
                isinstance(kw.value, ast.Constant) and isinstance(kw.value.value, bool)
            ):
                return json.dumps(
                    {"ok": False, "message": f"{node.func.id}() headers 인수는 True/False 리터럴이어야 합니다"},
                    ensure_ascii=False,
                )
        refs.append(node.args[0].value)
    return json.dumps({"ok": True, "refs": list(dict.fromkeys(refs))}, ensure_ascii=False)


def _pygrid_xl_scalar(v, t):
    """단일 셀 값 → Python 스칼라 (§3.3 입력 행)."""
    import pandas as pd

    if v is None:
        return None
    if t == "d":
        return pd.Timestamp(v)
    return v  # n → int/float, s → str, b → bool (JSON이 이미 그 타입)


def _pygrid_xl_column(vals, types):
    """열 하나를 §3.3 입력 규칙으로 변환한다.

    all 'n' 정수·빈 셀 없음 → int64 / 'n' 그 외 → float64+NaN /
    all 'd' → datetime64[ns]+NaT / all 'b' 빈 셀 없음 → bool /
    그 외(혼합·'s'·빈 열) → object+None  (경계 케이스 로그 #1)
    """
    import pandas as pd

    nonempty = {t for v, t in zip(vals, types) if v is not None}
    has_empty = any(v is None for v in vals)
    if nonempty == {"n"}:
        if not has_empty and all(isinstance(v, int) and not isinstance(v, bool) for v in vals):
            return pd.array(vals, dtype="int64")
        return pd.array(
            [float(v) if v is not None else float("nan") for v in vals], dtype="float64"
        )
    if nonempty == {"d"}:
        # pandas 3.x는 ISO 문자열에서 us 해상도를 추론하므로 §3.3의 ns로 고정한다
        return pd.to_datetime(list(vals), errors="coerce").as_unit("ns")  # None → NaT
    if nonempty == {"b"} and not has_empty:
        return pd.array(vals, dtype="bool")
    # §3.3: 's'·혼합·빈 열 → object (pandas 3.x의 전용 str dtype 대신. None 유지)
    return pd.Series(list(vals), dtype="object")


def _pygrid_xl(ref, headers=False):
    """xl("A1:C10", headers=True) — 주입된 스냅샷을 §3.3 규칙으로 변환해 돌려준다."""
    import pandas as pd

    snap = _pygrid_xl_cache.get(ref)
    if snap is None:
        raise RuntimeError(
            f"sheet(): 참조 {ref!r}의 데이터가 준비되지 않았습니다 (계산 엔진이 스냅샷을 주입하지 않음)"
        )
    values, types = snap["values"], snap["types"]
    if snap.get("scalar"):
        return _pygrid_xl_scalar(values[0][0], types[0][0])

    names = None
    if headers:
        header_row = values[0]
        names = [v if v is not None else f"Unnamed: {j}" for j, v in enumerate(header_row)]
        values, types = values[1:], types[1:]

    ncols = len(values[0]) if values else (len(names) if names else 0)
    data = {
        j: _pygrid_xl_column([r[j] for r in values], [r[j] for r in types])
        for j in range(ncols)
    }
    df = pd.DataFrame(data)
    if df.empty and ncols:
        df = pd.DataFrame({j: [] for j in range(ncols)})
    df.columns = names if names is not None else pd.RangeIndex(ncols)
    return df


sheet = _pygrid_xl
xl = _pygrid_xl  # Python in Excel 호환 별칭
