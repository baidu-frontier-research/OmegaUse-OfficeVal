from core.csv_reporter import _completion_rate, _minimum_score, _safe_text
from core.result_store import build_system_result



def _verifier_error(error: str) -> dict:
    return {
        "id": "022",
        "file_name": None,
        "status": "error",
        "error": error,
        "dim1_pass": True,
        "dim1_reason": "",
        "dim2_items": [
            {
                "rule": "不应保留",
                "max_delta": 1,
                "delta": 1,
                "hit": True,
                "detail": "不应保留",
            }
        ],
        "total_score": 1,
        "max_score": 9,
    }



def test_safe_text_escapes_spreadsheet_formula_prefixes() -> None:
    assert _safe_text("=1+1") == "'=1+1"
    assert _safe_text("+cmd") == "'+cmd"
    assert _safe_text("plain") == "plain"


def test_dim1_failure_has_zero_completion_rate() -> None:
    result = {
        "status": "ok",
        "dim1_pass": False,
        "total_score": 0,
        "max_score": 0,
    }

    assert _completion_rate(result) == "0.0%"
    result["status"] = "skipped"
    assert _completion_rate(result) == ""



def test_missing_deliverable_error_becomes_dim1_failure() -> None:
    errors = (
        "目录内缺少必需文档: a.docx, b.docx",
        "文件不存在: target.pptx",
        "FileNotFoundError: 目录内未找到必需文件: target.xlsx",
        "未在目录 'officeval_073' 中找到 .pptx/.ppt 文件",
        "未在目录中发现 .xlsx/.xlsm 文档",
    )

    for error in errors:
        result = build_system_result(
            "022",
            _verifier_error(error),
            "error",
            error,
            0.34,
            "2026-07-22T00:00:00+08:00",
        )

        assert result["status"] == "ok"
        assert result["error"] is None
        assert result["dim1_pass"] is False
        assert result["dim1_reason"] == error
        assert result["dim2_items"] == []
        assert result["total_score"] == 0
        assert result["max_score"] == 9
        assert _completion_rate(result) == "0.0%"


def test_unrelated_errors_remain_errors() -> None:
    error = "ValueError: 文档内容解析失败"
    result = build_system_result(
        "022",
        _verifier_error(error),
        "error",
        error,
        0.34,
        "2026-07-22T00:00:00+08:00",
    )

    assert result["status"] == "error"
    assert result["error"] == error
    assert result["dim1_reason"] == ""

    content_error = "未找到文本为『目标层』的外侧流程框图"
    content_result = build_system_result(
        "073",
        _verifier_error(content_error),
        "error",
        content_error,
        0.34,
        "2026-07-22T00:00:00+08:00",
    )
    assert content_result["status"] == "error"
    assert content_result["error"] == content_error

    system_error = "FileNotFoundError: worker.py 不存在"
    system_result = build_system_result(
        "022",
        None,
        "error",
        system_error,
        0.34,
        "2026-07-22T00:00:00+08:00",
    )
    assert system_result["status"] == "error"
    assert system_result["error"] == system_error


def test_minimum_score_sums_only_negative_max_delta() -> None:

    result = {
        "dim2_items": [
            {"max_delta": 5},
            {"max_delta": -2},
            {"max_delta": -3.5},
            {"max_delta": True},
            {"max_delta": None},
        ]
    }

    assert _minimum_score(result) == -5.5
