## Summary

Describe the change and why it is needed.

## Verification

- [ ] `python -m compileall -q core verifiers omegause_officeval`
- [ ] `python -m pytest`
- [ ] `python -m build`
- [ ] No private documents, credentials, absolute local paths, or generated runtime files are included.
- [ ] Documentation and changelog are updated when public behavior changes.

## Office COM Impact

- [ ] This change does not add a new COM dependency.
- [ ] If COM is required, the static parsing alternative and platform behavior are documented.
