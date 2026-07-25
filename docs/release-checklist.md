# Public Release Checklist

Complete this checklist before publishing the repository.

## Legal and Ownership

- [ ] The copyright owner approves publication under Apache License 2.0.
- [ ] Every verifier and copied code fragment is owned by the publisher or has
      a compatible third-party license and attribution.
- [ ] Product names, trademarks, and benchmark task descriptions are approved
      for public use.
- [ ] `NOTICE` contains all required attributions.
- [x] The PyMuPDF AGPL/commercial dual-license requirement is resolved: PyMuPDF
      has been replaced with permissively licensed pdfplumber and pypdfium2.
- [ ] Task literals that resemble personal names, email addresses, postal addresses, or phone numbers are approved for public release or replaced with synthetic data together with matching public fixtures.


## Data and Security

- [ ] No benchmark documents, user submissions, generated reports, or extracted
      workspaces are present.
- [ ] Secret scanning reports no credentials, tokens, private URLs, or private
      keys.
- [ ] All examples and tests use synthetic redistributable data.
- [ ] A private security contact or repository advisory channel is configured.

## Repository

- [ ] Replace any placeholder repository links after the public URL exists.
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
