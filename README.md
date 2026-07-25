# Omegause Officeval

[简体中文](README_zh-CN.md)

Omegause Officeval is a Python framework for securely validating, executing,
and aggregating 100 Office document evaluators. It accepts a ZIP submission,
checks its structure before extraction, runs each verifier in an isolated
subprocess, and writes machine-readable JSON and CSV reports.

> This repository contains the evaluation framework and verifier source code.
> Benchmark documents and submitted Office files are not distributed.

## Features

- Secure ZIP validation, including traversal, encryption, size, count, and
  compression-ratio checks.
- A uniform `evaluate(directory: str) -> dict` contract for verifiers
  `officeval_001` through `officeval_100`.
- Isolated subprocess execution with configurable concurrency and timeout.
- Visible progress, active verifier IDs, execution channel, and elapsed time.
- Atomic JSON and CSV reports with one result per verifier.
- Cross-platform normal mode for 96 verifiers.
- A serialized Office COM channel on Windows for verifiers `011`, `023`,
  `039`, and `081`.

## Requirements

- Python 3.10 or newer.
- Windows, macOS, or Linux for normal-mode verification.
- Microsoft Word and Excel on Windows only when the four COM-required
  verifiers must run.

## Installation

```bash
python -m venv .venv
```

On Linux or macOS:

```bash
source .venv/bin/activate
python -m pip install -e ".[test]"
```

On Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[test]"
```

`pywin32` is selected automatically on Windows and is not installed on other
platforms.

## Submission Layout

The input must be a ZIP archive whose root contains 100 three-digit task
directories:

```text
officeval_001/
officeval_002/
...
officeval_100/
```

Each directory contains the Office or PDF files required by its verifier.
Validation reports unexpected files but never deletes submission content.

## Usage

```bash
omegause-officeval --package /absolute/path/to/submission.zip
```

Equivalent module entry points are available:

```bash
python -m omegause_officeval --package /absolute/path/to/submission.zip
python -m core --package /absolute/path/to/submission.zip
```

Common options:

```text
--max-workers N
--timeout-seconds SECONDS
--com-mode auto|enabled|disabled
```

COM modes only affect verifiers `011`, `023`, `039`, and `081`:

- `auto`: enable COM for those verifiers on Windows and skip them elsewhere.
- `enabled`: require Windows and enable COM for those verifiers.
- `disabled`: skip those four verifiers on every platform.

Every verifier with a static parsing path remains in normal mode even when
`--com-mode enabled` is selected.

After validation, the CLI displays Fatal and Warning issues and asks for
explicit confirmation before evaluation starts.

## Output

Each submission creates a job-specific directory under `results/`:

```text
job.json
validation_report.json
summary.json
summary.csv
details.csv
001.json
...
100.json
```

The local `results/`, `submissions/`, and `workspaces/` directories are runtime
state and are ignored by Git.

## Workspace Cleanup

Extracted files remain under `workspaces/<job_id>/` after evaluation so they can
be inspected. Use the cleanup module to reclaim disk space:

```bash
# List eligible jobs, status, modification time, and estimated space only
python -m core.cleanup --list

# Delete one completed job workspace
python -m core.cleanup --job-id "<job_id>"

# Delete terminal-state workspaces older than 30 days
python -m core.cleanup --older-than-days 30
```

Deletion commands display their targets and require an explicit `y` or `yes`.
Only `workspaces/<job_id>/` is removed. The archived submission under
`submissions/<job_id>/` and JSON/CSV reports under `results/<job_id>/` remain.
Running or unknown-status jobs, symbolic links, and directory junctions are
skipped.

## Development

```bash
python -m compileall -q core verifiers omegause_officeval
python -m pytest
python -m build
```

See [Architecture](docs/architecture.md),
[Verifier Contract](docs/verifier-contract.md), and
[Security Model](docs/security.md) for implementation details.

## Contributing

Read [CONTRIBUTING.md](CONTRIBUTING.md) before opening an issue or pull
request. Please report security issues according to [SECURITY.md](SECURITY.md)
rather than through a public issue.

## License

Licensed under the [Apache License 2.0](LICENSE). See [NOTICE](NOTICE) for
attribution information and [Third-Party Licenses](THIRD_PARTY_LICENSES.md) for
dependency terms.
