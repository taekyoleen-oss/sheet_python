# 출력 변환 — 설계서 §3.3 출력 행 전체. 워커가 bootstrap.py·xl.py 다음에 로드한다.
# 마지막 표현식 값 → RunPayload 필드(JSON-safe). PyProxy는 밖으로 나가지 않는다.


def _pygrid_is_figure(v):
    return type(v).__name__ == "Figure" and type(v).__module__.startswith("matplotlib")


def _pygrid_cell(v):
    """스칼라 하나 → OutCell dict {"v","t","f"?}. NaN/NaT/None/inf → 빈 셀(v null)."""
    import datetime as _dt
    import math

    import numpy as np
    import pandas as pd

    try:
        if v is None or bool(pd.isna(v)):
            return {"v": None, "t": "s"}
    except (TypeError, ValueError):
        pass
    if isinstance(v, np.datetime64):
        v = pd.Timestamp(v)
    elif isinstance(v, np.generic):
        v = v.item()
    if isinstance(v, bool):
        return {"v": v, "t": "b"}
    if isinstance(v, (int, float)):
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            return {"v": None, "t": "s"}
        return {"v": v, "t": "n"}
    if isinstance(v, (pd.Timestamp, _dt.datetime, _dt.date)):
        ts = pd.Timestamp(v)
        if ts.hour == ts.minute == ts.second == 0 and ts.microsecond == 0 and ts.nanosecond == 0:
            return {"v": ts.strftime("%Y-%m-%d"), "t": "d", "f": "yyyy-mm-dd"}
        # 초 미만은 f 서식과 맞춰 잘라낸다 (경계 케이스 로그 #10)
        return {"v": ts.strftime("%Y-%m-%d %H:%M:%S"), "t": "d", "f": "yyyy-mm-dd hh:mm:ss"}
    if isinstance(v, str):
        return {"v": v, "t": "s"}
    return {"v": str(v), "t": "s"}  # 그 외 객체 → str() 1셀 (§3.3)


def _pygrid_df_include_index(df, include_index):
    import pandas as pd

    if include_index == "always":
        return True
    if include_index == "never":
        return False
    idx = df.index  # auto: 기본 RangeIndex(0,1,2…)만 제외 (경계 케이스 로그 #3)
    return not (isinstance(idx, pd.RangeIndex) and idx.start == 0 and idx.step == 1)


def _pygrid_df_cells(df, include_index):
    with_index = _pygrid_df_include_index(df, include_index)
    header = []
    if with_index:
        header.append({"v": str(df.index.name or ""), "t": "s"})
    header += [{"v": str(c), "t": "s"} for c in df.columns]
    rows = [header]
    for i in range(df.shape[0]):
        row = [_pygrid_cell(df.index[i])] if with_index else []
        row += [_pygrid_cell(df.iloc[i, j]) for j in range(df.shape[1])]
        rows.append(row)
    return rows


def _pygrid_to_cells(value, include_index):
    """값 모드: §3.3 출력 규칙으로 OutCell 2D를 만든다."""
    import numpy as np
    import pandas as pd

    if isinstance(value, pd.DataFrame):
        return _pygrid_df_cells(value, include_index)
    if isinstance(value, pd.Series):
        cells = [[_pygrid_cell(x)] for x in value]
        if value.name is not None:  # Series는 name을 헤더로 (§3.3)
            cells.insert(0, [{"v": str(value.name), "t": "s"}])
        return cells
    if isinstance(value, np.ndarray):
        if value.ndim == 1:
            return [[_pygrid_cell(x)] for x in value.tolist()]
        if value.ndim == 2:
            return [[_pygrid_cell(x) for x in row] for row in value.tolist()]
        return [[_pygrid_cell(value)]]  # 3D 이상 → str() 1셀
    if isinstance(value, dict):
        return [[_pygrid_cell(k), _pygrid_cell(v)] for k, v in value.items()] or [
            [{"v": None, "t": "s"}]
        ]
    if isinstance(value, (list, tuple)):
        if value and all(isinstance(r, (list, tuple)) for r in value):
            # 들쭉날쭉한 행은 최대 너비로 빈 셀 패딩 (경계 케이스 로그 #11)
            width = max(len(r) for r in value)
            return [
                [_pygrid_cell(x) for x in row] + [{"v": None, "t": "s"}] * (width - len(row))
                for row in value
            ]
        return [[_pygrid_cell(x)] for x in value] or [[{"v": None, "t": "s"}]]
    return [[_pygrid_cell(value)]]  # 스칼라·그 외 객체


def _pygrid_repr(value, limit=2000):
    r = repr(value)
    return r if len(r) <= limit else r[: limit] + "…"


def _pygrid_png_b64(fig):
    import base64
    import io

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150)
    return base64.b64encode(buf.getvalue()).decode()


def _pygrid_to_preview(value):
    """객체 모드: (kind, shape, preview, pngB64). 미리보기는 index 없이 열만 보여준다."""
    import datetime as _dt

    import numpy as np
    import pandas as pd

    if _pygrid_is_figure(value):
        return "image", None, {"kind": "image"}, _pygrid_png_b64(value)
    if isinstance(value, pd.Series):
        value = value.to_frame()
    if isinstance(value, pd.DataFrame):
        head = value.head(100)
        rows = [
            [_pygrid_cell(head.iloc[i, j])["v"] for j in range(head.shape[1])]
            for i in range(head.shape[0])
        ]
        preview = {
            "kind": "table",
            "columns": [str(c) for c in value.columns],
            "dtypes": [str(d) for d in value.dtypes],
            "rows": rows,
            "shape": list(value.shape),
        }
        return "table", list(value.shape), preview, None
    if value is None or isinstance(
        value, (bool, int, float, str, np.generic, _dt.date, _dt.datetime, pd.Timestamp)
    ):
        return "scalar", None, {"kind": "repr", "repr": _pygrid_repr(value)}, None
    shape = None
    if isinstance(value, np.ndarray) and value.ndim in (1, 2):
        shape = list(value.shape) if value.ndim == 2 else [value.shape[0], 1]
    return "object", shape, {"kind": "repr", "repr": _pygrid_repr(value)}, None


def _pygrid_select_output(value, output):
    """출력 선택의 열·행 필터를 §3.3 변환 **전에** 적용한다.

    변환 전에 값을 좁히므로 값 모드(spill)와 객체 모드(preview/shape)가 같은 선택을 반영한다.
    variable 선택은 _pygrid_exec_capture가 처리한다(경계 케이스 로그 #15~#19).
    """
    if not output:
        return value

    import numpy as np
    import pandas as pd

    columns = output.get("columns")
    if columns and isinstance(value, pd.DataFrame):
        # 열 이름은 str()로 맞춘다 — headers=False DataFrame의 정수 컬럼도 선택 가능 (#19)
        by_name = {}
        for c in value.columns:
            by_name.setdefault(str(c), c)
        keep = [by_name[c] for c in columns if c in by_name]
        if keep:  # 요청 열이 하나도 없으면 전체 열 유지 — 빈 spill 방지 (#16)
            value = value[keep]

    limit = output.get("rowLimit")
    if isinstance(limit, (int, float)) and limit > 0:  # 0 이하는 무제한 (#17)
        n = int(limit)
        if isinstance(value, (pd.DataFrame, pd.Series)):
            value = value.iloc[:n]
        elif isinstance(value, np.ndarray) and value.ndim in (1, 2):
            value = value[:n]
        elif isinstance(value, (list, tuple)):
            value = value[:n]
        # 그 외(스칼라·dict·Figure·객체)는 행 개념이 없어 무시한다
    return value


def _pygrid_convert(value, output_mode, include_index):
    """마지막 표현식 값 → 결과 dict. 워커가 JSON으로 받아 RunPayload로 변환한다."""
    type_name = type(value).__name__
    if output_mode == "values":
        if _pygrid_is_figure(value):  # 경계 케이스 로그 #4
            return {
                "ok": False,
                "etype": "PyGridImageError",
                "msg": "이미지는 값으로 펼칠 수 없습니다",
                "tb": "",
            }
        cells = _pygrid_to_cells(value, include_index)
        if not cells or not cells[0]:
            # 빈 DataFrame(0열)·빈 list/Series/ndarray → 항상 유효한 1×1 빈 셀 보장
            cells = [[{"v": None, "t": "s"}]]
        shape = [len(cells), len(cells[0])]
        kind = "scalar" if shape == [1, 1] else "table"
        return {
            "ok": True,
            "kind": kind,
            "typeName": type_name,
            "shape": shape,
            "cells": cells,
        }
    kind, shape, preview, png_b64 = _pygrid_to_preview(value)
    out = {"ok": True, "kind": kind, "typeName": type_name, "preview": preview}
    if shape is not None:
        out["shape"] = shape
    if png_b64 is not None:
        out["pngB64"] = png_b64
    return out


def _pygrid_run_convert(code, output_mode, include_index, output_json=None):
    """블록 실행 + 변환. 반환(JSON 문자열): _pygrid_convert 결과 또는 실패 dict.

    output_json: OutputSelection(JSON 문자열) 또는 None/"null" — 없으면 기존 동작(마지막 표현식).
    """
    import json

    output = json.loads(output_json) if output_json else None
    try:
        value = _pygrid_exec_capture(code, (output or {}).get("variable"))
        value = _pygrid_select_output(value, output)
    except BaseException as e:  # KeyboardInterrupt·출력 변수 NameError 포함
        return json.dumps(_pygrid_fail(e), ensure_ascii=False)
    try:
        return json.dumps(_pygrid_convert(value, output_mode, include_index), ensure_ascii=False)
    except Exception as e:  # 변환 자체의 예상 밖 실패
        return json.dumps(_pygrid_fail(e, "ConversionError"), ensure_ascii=False)


def _pygrid_fail(e, etype=None):
    """예외 → 실패 dict(JSON-safe). etype을 주면 합성 오류 이름으로 덮어쓴다."""
    return {
        "ok": False,
        "etype": etype or type(e).__name__,
        "msg": str(e),
        "tb": _pygrid_format_exc(e),
    }


def _pygrid_convert_one(last, req):
    """출력 요청 하나 → 결과 dict(+id). **이 출력만의** 실패는 여기서 잡는다.

    다른 출력은 계속 진행한다(경계 케이스 로그 #20). 중단(KeyboardInterrupt)만
    위로 올려 실행 전체 실패로 만든다.
    """
    sel = req.get("selection") or {}
    try:
        value = last
        variable = sel.get("variable")
        if variable:
            value = _pygrid_resolve(variable)  # 변수명 또는 출력 식 (부록 O.5)
        value = _pygrid_select_output(value, sel)
    except KeyboardInterrupt:
        raise
    except BaseException as e:
        return dict(_pygrid_fail(e), id=req.get("id"))
    try:
        out = _pygrid_convert(value, req.get("mode", "object"), req.get("includeIndex", "auto"))
    except Exception as e:  # 변환 자체의 예상 밖 실패 (KeyboardInterrupt는 위로)
        out = _pygrid_fail(e, "ConversionError")
    out["id"] = req.get("id")
    return out


def _pygrid_run_convert_multi(code, outputs_json):
    """다중 출력: 본문을 **1회만** 실행하고 요청마다 따로 변환한다.

    반환(JSON 문자열):
      성공 {"ok": true, "items": [{"id", ...변환 결과 또는 출력 단위 실패}]}
      실패 {"ok": false, "etype", "msg", "tb"}  — 본문 실행 실패(전체 run 실패)
    """
    import json

    requests = json.loads(outputs_json) if outputs_json else []
    try:
        last = _pygrid_exec_capture(code)  # variable 선택은 출력마다 따로 적용한다
        items = [_pygrid_convert_one(last, req) for req in requests or []]
    except BaseException as e:  # 본문 오류·중단 → 실행 전체 실패
        return json.dumps(_pygrid_fail(e), ensure_ascii=False)
    return json.dumps({"ok": True, "items": items}, ensure_ascii=False)
