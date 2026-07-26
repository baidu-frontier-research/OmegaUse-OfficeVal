# Public Release Checklist

Complete this checklist before publishing the repository.

## Legal and Ownership

- [x] Baidu Inc. approves publication under Apache License 2.0, including the
      Apache-2.0 patent grant.
- [x] Every verifier and copied code fragment is owned by Baidu Inc. or is
      covered by an approved compatible third-party license and attribution.
- [x] Product names, trademarks, and benchmark task descriptions are approved
      for public use.
- [x] `NOTICE` contains the confirmed copyright owner and required attributions.
- [x] The PyMuPDF AGPL/commercial dual-license requirement is resolved: PyMuPDF
      has been replaced with permissively licensed pdfplumber and pypdfium2.
- [x] Verifier literals were reviewed and confirmed de-identified; no additional
      personal-name, address, phone, or email approval is required for them.
- [x] The directly used and redistributed dependency licenses were reviewed for
      this release.

## Data and Security


- [x] Data and non-code assets in the public release were confirmed suitable
  for redistribution; user submissions, generated reports, and extracted
  workspaces are excluded.
- [x] The release boundary was confirmed: public source, documentation, and
  approved synthetic/example assets only; no private benchmark answers or
  user-delivered Office files.
- [x] Secret scanning reports no credentials, tokens, private URLs, or private

  keys; the public security scan returned zero findings on 2026-07-26.
- [x] All examples and tests use synthetic redistributable data.
- [x] Private security reports use `smart@baidu.com` as the official channel;
  the policy promises processing as soon as reasonably possible without a fixed
  response-time commitment.



## Repository

- [x] Replace any placeholder repository links after the public URL exists.
  The public repository is `https://github.com/baidu-frontier-research/OmegaUse-OfficeVal`.

- [ ] Enable branch protection and required CI checks.
- [ ] Enable dependency and secret scanning.
- [ ] Review issue and pull request templates.
- [ ] Create the first signed release and attach verified distributions.

## Verification

```bash
python -m pip install -e ".[test]"
python -m compileall -q core verifiers omegause_officeval
python -m pytest
python -m build
```
