# Examples

Public examples must use synthetic Office documents that are safe to
redistribute. Real submissions, benchmark answers, and generated task
workspaces must not be added to this directory.

After creating a synthetic ZIP with the required `officeval_NNN/` layout, run:

```bash
omegause-officeval --package /absolute/path/to/synthetic-submission.zip \
  --com-mode disabled
```

The CLI validates the archive and asks for confirmation before execution.
