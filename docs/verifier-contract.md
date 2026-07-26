# Verifier Contract

## File and Directory Names

Verifier source files and submission directories use matching three-digit IDs:

```text
verifiers/officeval_001_verifier.py
officeval_001/
```

IDs must remain strings so leading zeroes are preserved.

## Entry Point

Every verifier exposes:

```python
def evaluate(directory: str) -> dict:
    ...
```

The argument is the matching submission directory, not a file path. A verifier
may select one or more supported documents from that directory.

## Result

A successful verifier returns one dictionary containing:

```python
{
    "id": "001",
    "file_name": "document.docx",
    "status": "ok",
    "error": None,
    "dim1_pass": True,
    "dim1_reason": "",
    "dim2_items": [],
    "total_score": 0,
    "max_score": 0,
}
```

`dim2_items` contains rule-level descriptions, deltas, maximum deltas, and
optional evidence. System-level `error`, `timeout`, and `skipped` results are
created by the batch worker rather than by normal verifier scoring.

For the complete JSON and CSV field definitions, status meanings, completion
formula, and missing-deliverable behavior, see [Result Format](result-format.md).
In particular, `status="ok"` means that evaluation completed normally; it does
not imply that `dim1_pass` is true or that the score is positive.

## Rules


- Do not use absolute local paths.
- Do not delete or modify submitted files.
- Do not rely on process-wide mutable state.
- Close ZIP files, Office documents, images, and subprocesses deterministically.
- Prefer OOXML and other static parsing over Office automation.
- Do not start COM unless the verifier is listed as COM-required.
- Keep standard output diagnostic-only; return the authoritative result.
- Use only synthetic, redistributable files in public tests.
