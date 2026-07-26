import ipaddress
import re
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
CHECKED_SUFFIXES = {".py", ".md", ".toml", ".txt", ".yml", ".yaml"}
IGNORED_PARTS = {"dist", "build", ".pytest_cache"}
PUBLIC_DOMAINS = {
    "apache.org",
    "arxiv.org",
    "baidu.com",
    "contributor-covenant.org",
    "example.com",
    "example.net",
    "example.org",
    "github.com",
    "huggingface.co",

    "keepachangelog.com",
    "omegause-officeval.github.io",
    "semver.org",
}
DOMAIN_RE = re.compile(
    r"\b(?:[a-z0-9-]+\.)+"
    r"(?:co|com|corp|dev|internal|intranet|io|lan|local|net|org)\b",

    re.IGNORECASE,
)
IPV4_RE = re.compile(r"(?<![\d.])(?:\d{1,3}\.){3}\d{1,3}(?![\d.])")
WINDOWS_USER_PATH_RE = re.compile(
    r"\b[A-Z]:\\Users\\[^\\\s\"']+",
    re.IGNORECASE,
)
CREDENTIAL_ASSIGNMENT_RE = re.compile(
    r"\b(?:access[_-]?key|api[_-]?key|auth[_-]?token|bearer[_-]?token|"
    r"client[_-]?secret|password|passwd|secret(?:[_-]?key)?)"
    r"\s*[:=]\s*[\"'][^\"']{4,}[\"']",
    re.IGNORECASE,
)
CLOUD_CREDENTIAL_RE = re.compile(
    r"\b(?:AKIA[0-9A-Z]{16}|ASIA[0-9A-Z]{16}|"
    r"BCEAK[A-Za-z0-9]{12,40}|LTAI[A-Za-z0-9]{12,24})\b"
)
PRIVATE_KEY_HEADER = "-" * 5 + "BEGIN " + "PRIVATE KEY" + "-" * 5


def _iter_public_text_files():
    for path in ROOT.rglob("*"):
        if (
            path.is_file()
            and path.suffix.lower() in CHECKED_SUFFIXES
            and "verifiers" not in path.relative_to(ROOT).parts
            and not any(part in IGNORED_PARTS for part in path.parts)

        ):
            yield path


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
    assert "--index-url" not in metadata
    assert "--extra-index-url" not in metadata
    assert "--trusted-host" not in metadata
    assert "Private :: Do Not Upload" not in metadata
    assert "Proprietary" not in metadata


def test_internal_information_patterns_are_active() -> None:
    private_ip = "10" + ".0.0.1"
    local_path = "C:" + "\\Users\\sample\\project"
    internal_domain = "service" + ".internal"
    password_assignment = "password" + "='placeholder'"
    cloud_key = "AKIA" + "A" * 16

    assert ipaddress.ip_address(private_ip).is_private
    assert WINDOWS_USER_PATH_RE.search(local_path) is not None
    assert internal_domain in DOMAIN_RE.findall(internal_domain)
    assert CREDENTIAL_ASSIGNMENT_RE.search(password_assignment) is not None
    assert CLOUD_CREDENTIAL_RE.search(cloud_key) is not None


def test_public_text_has_no_internal_addresses_or_credentials() -> None:

    for path in _iter_public_text_files():
        text = path.read_text(encoding="utf-8")
        assert WINDOWS_USER_PATH_RE.search(text) is None, path
        assert CREDENTIAL_ASSIGNMENT_RE.search(text) is None, path
        assert CLOUD_CREDENTIAL_RE.search(text) is None, path
        assert PRIVATE_KEY_HEADER not in text, path

        for candidate in IPV4_RE.findall(text):
            try:
                address = ipaddress.ip_address(candidate)
            except ValueError:
                continue
            assert not address.is_private, (path, candidate)

        for domain in DOMAIN_RE.findall(text):
            normalized = domain.lower().removeprefix("www.")
            assert normalized in PUBLIC_DOMAINS, (path, domain)
