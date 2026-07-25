# Contributing

Thank you for contributing to Omegause Officeval.

## Before You Start

- Search existing issues before opening a new one.
- Use a public issue for bugs and feature proposals.
- Follow `SECURITY.md` for vulnerabilities or sensitive reports.
- Do not include Office documents, personal information, credentials, or
  proprietary datasets in issues, commits, or test fixtures.

## Development Setup

```bash
python -m venv .venv
python -m pip install -e ".[test]"
```

Run the project checks before submitting a pull request:

```bash
python -m compileall -q core verifiers omegause_officeval
python -m pytest
python -m build
```

## Verifier Contract

Each verifier must remain named `officeval_NNN_verifier.py` and expose:

```python
def evaluate(directory: str) -> dict:
    ...
```

Keep three-digit IDs, avoid absolute paths, and never delete submission
content. A verifier failure must be returned as structured data or isolated by
the worker; it must not terminate unrelated evaluations.

Only verifiers listed in `COM_REQUIRED_VERIFIER_IDS` may start Office COM.
Prefer static OOXML parsing whenever the required behavior can be evaluated
without launching Office.

## Pull Requests

- Keep changes focused and explain observable behavior changes.
- Add or update tests for behavior that can be covered without proprietary
  documents.
- Use synthetic, minimal fixtures that are safe to redistribute.
- Update documentation and `CHANGELOG.md` when the public interface changes.
- Confirm that generated files and runtime directories are not committed.

By submitting a contribution, you agree that it is licensed under the Apache
License 2.0 and that you have the right to submit it.
