from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_public_tree_excludes_runtime_and_internal_ci() -> None:
    assert not (ROOT / "ci.yml").exists()
    assert not (ROOT / "results").exists()
    assert not (ROOT / "submissions").exists()
    assert not (ROOT / "workspaces").exists()


def test_public_metadata_has_no_private_package_index() -> None:
    metadata = "\n".join(
        (ROOT / name).read_text(encoding="utf-8")
        for name in ("pyproject.toml", "requirements.txt")
    )
    assert "baidu-int.com" not in metadata
    assert "pip.baidu" not in metadata
    assert "Private :: Do Not Upload" not in metadata
    assert "Proprietary" not in metadata


def test_public_text_has_no_private_repository_markers() -> None:
    markers = (
        "guokaiqi",
        "icode.baidu.com",
        "ku.baidu-int.com",
        "pip.baidu-int.com",
        "C:\\Users\\guokaiqi",
        "All Rights Reserved",
    )
    checked_suffixes = {".py", ".md", ".toml", ".txt", ".yml", ".yaml"}
    for path in ROOT.rglob("*"):
        if (
            not path.is_file()
            or path == Path(__file__)
            or path.suffix.lower() not in checked_suffixes
            or any(part in {"dist", "build", ".pytest_cache"} for part in path.parts)
        ):
            continue
        text = path.read_text(encoding="utf-8")
        assert not any(marker in text for marker in markers), path
