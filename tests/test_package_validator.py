import zipfile

from core.document_validator import _check_one_dir
from core.package_validator import validate_package


def test_legacy_office_formats_are_not_valid_documents(tmp_path) -> None:
    task_dir = tmp_path / "officeval_001"
    task_dir.mkdir()
    for name in ("legacy.doc", "legacy.ppt", "legacy.xls"):
        (task_dir / name).write_bytes(b"legacy")

    issues = _check_one_dir("001", task_dir)

    unknown_paths = {
        issue.path for issue in issues if issue.code == "UNKNOWN_FILE_TYPE"
    }
    assert unknown_paths == {
        "officeval_001/legacy.doc",
        "officeval_001/legacy.ppt",
        "officeval_001/legacy.xls",
    }
    assert any(issue.code == "NO_VALID_DOC" for issue in issues)


def test_modern_office_formats_remain_valid_documents(tmp_path) -> None:
    task_dir = tmp_path / "officeval_001"
    task_dir.mkdir()
    for name in ("modern.docx", "modern.pptx", "modern.xlsx", "macro.xlsm"):
        (task_dir / name).write_bytes(b"modern")

    issues = _check_one_dir("001", task_dir)

    assert not any(issue.code in {"UNKNOWN_FILE_TYPE", "NO_VALID_DOC"} for issue in issues)


def test_rejects_parent_directory_traversal(tmp_path) -> None:
    archive = tmp_path / "traversal.zip"
    with zipfile.ZipFile(archive, "w") as zf:
        zf.writestr("../escape.txt", "unsafe")

    issue_codes = {issue.code for issue in validate_package(archive)}

    assert "ZIP_PATH_TRAVERSAL" in issue_codes


def test_rejects_top_level_file(tmp_path) -> None:
    archive = tmp_path / "top-level.zip"
    with zipfile.ZipFile(archive, "w") as zf:
        zf.writestr("unexpected.txt", "unsafe")

    issue_codes = {issue.code for issue in validate_package(archive)}

    assert "ZIP_TOP_LEVEL_FILES" in issue_codes
