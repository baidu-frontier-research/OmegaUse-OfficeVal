# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- Public project links, package metadata, a platform compatibility matrix, and
  a complete result-field reference in `docs/result-format.md`.
- A runnable synthetic ZIP generator under `examples/`.
- Explicit private security reporting through `smart@baidu.com`.

### Changed

- Office COM is restricted to verifiers `001`, `021`, `023`, `024`, `025`,
  and `039`. Verifiers listed in `COM_FORCED_NORMAL_VERIFIER_IDS`, including
  `007`, always run in normal mode with `OFFICEVAL_COM_ENABLED=0`.
- Supported deliverables are `.docx`, `.xlsx`, `.xlsm`, `.pptx`, and `.pdf`.
  Legacy `.doc`, `.xls`, and `.ppt` files are no longer supported.
- Missing task directories, empty directories, unsupported-only directories,
  and missing required deliverables are documented as dimension-one failures
  with zero score after confirmation.
- Verifiers `001` through `005` and the PPTX coordinate handling in `051` and
  `052` use the updated public scoring behavior.
- Verifier `008` uses the existing `pdfplumber`/`pypdfium2` PDF backend instead
  of PyMuPDF. Verifier `062` no longer uses the optional `imageio` path.
- The unused `imageio` and `langdetect` runtime declarations were removed; no
  verifier imports either package.

### Fixed

- Public documentation now distinguishes normal evaluation completion from a
  passing score and documents JSON/CSV completion semantics.
- Verifiers `013`, `021`, and `042` report structurally damaged deliverables
  as dimension-one failures instead of system errors; `076` and `088` apply
  the same classification to malformed OPC relationships.
- Verifier `023` clamps section indices for malformed section properties,
  `037` reads raw `w:jc` values so `start`/`end` alignments are recognised,
  `017` tolerates a missing target page, and `088` reads footer XML directly
  from the part instead of relying on paragraph proxies.


## [0.1.0] - 2026-07-16

### Added

- Secure ZIP validation, archival, and extraction.
- Batch execution of 100 isolated Office document verifiers.
- Configurable concurrency, timeout, and COM mode.
- JSON and CSV result reporting.
- Cross-platform normal-mode execution and Windows Office COM isolation.
