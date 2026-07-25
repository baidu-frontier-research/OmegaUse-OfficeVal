# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed

- Only verifiers `011`, `023`, `039`, and `081` may use Office COM.
- Verifiers with a static parsing path always use normal mode on every platform.

## [0.1.0] - 2026-07-16

### Added

- Secure ZIP validation, archival, and extraction.
- Batch execution of 100 isolated Office document verifiers.
- Configurable concurrency, timeout, and COM mode.
- JSON and CSV result reporting.
- Cross-platform normal-mode execution and Windows Office COM isolation.
