# Result Format

Each evaluation job writes one JSON result per task and aggregate JSON/CSV
reports under `results/<job_id>/`. The numbered JSON result is authoritative;
the CSV files are spreadsheet-friendly projections of the same data.

## Numbered Result

A normal result contains these fields:

```json
{
  "id": "001",
  "file_name": "document.docx",
  "status": "ok",
  "error": null,
  "dim1_pass": true,
  "dim1_reason": "",
  "dim2_items": [],
  "total_score": 0,
  "max_score": 0,
  "duration_seconds": 1.25,
  "started_at": "2026-07-26T10:00:00+08:00",
  "finished_at": "2026-07-26T10:00:01+08:00",
  "retried": false
}
```

| Field | Meaning |
| --- | --- |
| `id` | Three-digit task ID from `001` through `100`. |
| `file_name` | Evaluated file name or a verifier-specific list of file names. It may be empty when no deliverable is available. |
| `status` | Execution outcome: `ok`, `error`, `timeout`, or `skipped`. |
| `error` | System or execution error text. It is `null` for normal scoring results. |
| `dim1_pass` | Whether the deliverable satisfies the task's format and availability gate. |
| `dim1_reason` | Reason dimension one failed, including missing or invalid deliverables. |
| `dim2_items` | Rule-level scoring details. |
| `total_score` | Score awarded by the verifier after the dimension-one gate. |
| `max_score` | Maximum score reported by that task's verifier. |
| `duration_seconds` | Wall-clock verifier duration. |
| `started_at`, `finished_at` | ISO 8601 timestamps recorded by the worker. |
| `retried` | Whether the worker retried a transient Office COM failure. |

## Status Semantics

- `ok`: evaluation completed normally. This does not mean the deliverable passed
  dimension one or received a positive score.
- `error`: the verifier process failed, returned an invalid structure, or hit a
  system-level error that cannot be treated as a deliverable-quality issue.
- `timeout`: the task exceeded its configured timeout.
- `skipped`: the task was intentionally not run, for example because a required
  Windows Office COM channel was disabled or unavailable.

Business scoring failures therefore commonly use `status="ok"` together with
`dim1_pass=false`. Consumers must not interpret `status="ok"` as a passing
score.

## Dimension-Two Items

Each item in `dim2_items` contains:

| Field | Meaning |
| --- | --- |
| `rule` | Human-readable rule description. |
| `max_delta` | Score delta when the rule is hit. It may be negative for a penalty rule. |
| `delta` | Applied delta, normally either `0` or `max_delta`. |
| `hit` | Boolean indicating whether the rule was applied. |
| `detail` | Evidence or explanation produced by the verifier. |

Some dimension-one failures may retain explanatory dimension-two details, but
the authoritative `total_score` remains zero when the gate fails.

## Missing or Invalid Deliverables

A complete ZIP should contain `officeval_001/` through `officeval_100/`. The
following conditions are reported as validation warnings and can continue only
after explicit user confirmation:

- the numbered directory is missing;
- the directory is empty;
- the directory contains no supported document.

A verifier may also discover later that a required deliverable file does not
exist even though its directory contains another supported document. All of
these missing-deliverable cases are represented as normal zero-score business
outcomes:


```json
{
  "status": "ok",
  "error": null,
  "dim1_pass": false,
  "dim1_reason": "<missing or invalid deliverable reason>",
  "dim2_items": [],
  "total_score": 0
}
```

For directories that are missing, empty, or contain no supported document, the
batch runner constructs this result without entering the verifier. When a
verifier discovers a missing required file, the result store converts that
specific missing-deliverable error to the same semantics. Unrelated parsing or
system failures remain `status="error"`.

## CSV Reports and Completion Rate

`summary.csv` contains one row per task. In addition to the numbered-result
fields, it includes:

- `job_id`;
- `min_score`, calculated from all negative `max_delta` values;
- `completion_rate`;
- timing fields.

`details.csv` contains one row per dimension-two rule with `rule_index`, `rule`,
`max_delta`, `delta`, `hit`, and `detail`.

For an `ok` result with `dim1_pass=false`, completion is always `0.0%`.
Otherwise, when `max_score` is positive and numeric, the CSV calculation is:

```text
completion_rate = max(0, total_score) / max_score * 100%
```

The value is formatted with one decimal place. Completion is left blank when
the score or maximum is not numeric, `max_score <= 0`, or no completion can be
meaningfully calculated. Task score ranges differ, so raw scores and
`max_score` values should not be compared across task IDs without an explicit
normalization method.

## Aggregate Files

- `summary.json`: batch metadata plus all numbered results.
- `summary.csv`: one spreadsheet row per task.
- `details.csv`: one spreadsheet row per scoring rule.
- `validation_report.json`: package and document validation warnings/fatal
  issues that were observed before evaluation.

CSV text fields are escaped when they begin with spreadsheet formula prefixes,
and CSV files are written with a UTF-8 BOM for compatibility with spreadsheet
software.
