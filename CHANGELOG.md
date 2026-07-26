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

- Office COM is restricted to verifiers `001`, `008`, `019`, `022`, `023`,
  `030`, `039`, `074`, and `081`; verifier `011` uses normal mode with its
  fallback disabled by the scheduler.
- Supported deliverables are `.docx`, `.xlsx`, `.xlsm`, `.pptx`, and `.pdf`.
  Legacy `.doc`, `.xls`, and `.ppt` files are no longer supported.
- Missing task directories, empty directories, unsupported-only directories,
  and missing required deliverables are documented as dimension-one failures
  with zero score after confirmation.
- Verifiers `001` through `005` and the PPTX coordinate handling in `051` and
  `052` use the updated public scoring behavior.
- Verifier `008` uses the existing `pdfplumber`/`pypdfium2` PDF backend instead
  of PyMuPDF. Verifier `062` no longer uses the optional `imageio` path.

### Fixed

- Public documentation now distinguishes normal evaluation completion from a
  passing score and documents JSON/CSV completion semantics.


## [0.1.0] - 2026-07-16

### Added

- Secure ZIP validation, archival, and extraction.
- Batch execution of 100 isolated Office document verifiers.
- Configurable concurrency, timeout, and COM mode.
- JSON and CSV result reporting.
- Cross-platform normal-mode execution and Windows Office COM isolation.
