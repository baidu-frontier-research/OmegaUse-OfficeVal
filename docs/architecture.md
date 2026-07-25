# Architecture

Omegause Officeval separates submission handling, scheduling, verifier
execution, and reporting so that a single document or verifier failure cannot
terminate the batch.

## Components

- `omegause_officeval/`: installed command-line entry point.
- `core/`: package validation, secure extraction, scheduling, worker isolation,
  result persistence, CSV generation, and cleanup.
- `verifiers/`: 100 task-specific evaluators with a uniform interface.
- `tests/`: public tests that do not require benchmark documents or Office.

## Submission Lifecycle

1. Copy the submitted ZIP into immutable local submission storage.
2. Validate archive members before extraction.
3. Extract into a task-specific workspace.
4. Validate the expected `officeval_NNN/` directory layout.
5. Display Fatal and Warning issues and request confirmation.
6. Execute each verifier in an independent subprocess.
7. Persist numbered results and aggregate JSON/CSV reports atomically.
8. Retain task data until the caller explicitly runs cleanup.

Validation never deletes unexpected submission content.

## Scheduling

Normal verifiers run through a bounded parallel channel. Verifiers `011`,
`023`, `039`, and `081` require Office COM and share a serial channel on
Windows. All verifiers that can use static parsing are forced into normal mode,
even when batch COM mode is enabled.

The worker passes `OFFICEVAL_COM_ENABLED=1` only to a COM-required verifier on
a supported and enabled platform. Every other verifier receives `0`.

## Process Isolation

Each verifier is dynamically loaded inside a child process using its real
module name. The parent enforces a timeout, captures structured results, and
continues unrelated tasks after an error. COM tasks receive additional Office
process cleanup and controlled retries for transient automation failures.

## Result Model

Each numbered JSON result records status, dimension-one validation, scored
items, totals, timing, and system errors. `summary.json`, `summary.csv`, and
`details.csv` aggregate the same data for tools and spreadsheets.
