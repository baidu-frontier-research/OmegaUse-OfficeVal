# Examples

Public examples use synthetic Office documents that are safe to redistribute.
Real submissions, benchmark answers, golden artifacts, and generated task
workspaces must not be added to this directory.

## Minimal Synthetic Submission

`create_minimal_submission.py` creates a small ZIP containing one generated
DOCX under `officeval_002/`. The other 99 task directories are intentionally
absent so the example exercises package validation, warning confirmation,
missing-deliverable zero scoring, verifier execution, and report generation.
It is a pipeline smoke test and does not produce a meaningful benchmark score.

Create the ZIP from the repository root:

```bash
python examples/create_minimal_submission.py \
  --output synthetic-submission.zip
```

Run it on Linux or macOS:

```bash
printf 'y\n' | python -m core \
  --package ./synthetic-submission.zip \
  --com-mode disabled
```

Run it in Windows PowerShell:

```powershell
"y" | python -m core `
  --package .\synthetic-submission.zip `
  --com-mode disabled
```

Expected behavior:

- validation reports 99 missing task-directory warnings;
- after confirmation, missing items receive `status="ok"`,
  `dim1_pass=false`, zero score, and `0.0%` completion;
- `officeval_002` enters its verifier using only the generated DOCX;
- JSON and CSV reports are written under `results/<job_id>/`.

Delete `synthetic-submission.zip` and the generated job's `results/`,
`submissions/`, and `workspaces/` directories after the smoke test.
