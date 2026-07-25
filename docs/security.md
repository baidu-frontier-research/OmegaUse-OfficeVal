# Security Model

Office documents and archives are untrusted input. The framework applies
controls before and during evaluation, but callers should still run it with the
least operating-system privileges available.

## Archive Controls

Before extraction, the package validator rejects or limits:

- malformed and encrypted ZIP archives;
- absolute paths and parent-directory traversal;
- excessive member count, individual size, or total uncompressed size;
- suspicious compression ratios;
- members outside the expected task layout.

Extraction writes only beneath a newly created workspace. Unexpected content
is reported and never automatically deleted.

## Verifier Isolation

Each verifier runs in a child process with a configured timeout. A crash,
timeout, or malformed return value becomes a numbered system result and does
not stop unrelated evaluations.

## Office Automation

Office COM is disabled outside Windows and restricted to four verifier IDs.
COM tasks run serially, create independent application instances, close opened
documents in `finally` paths, and trigger process cleanup after timeout or
failure. Run COM evaluation under a dedicated, non-administrator account.

## Output Controls

JSON and CSV files are written atomically. CSV values are escaped to reduce
spreadsheet formula injection. Runtime directories should not be published or
committed because they can contain submitted filenames, validation details,
and extracted documents.

## Deployment Guidance

- Run evaluation on an isolated host or disposable VM.
- Do not expose the CLI directly as an unauthenticated upload service.
- Apply filesystem quotas in addition to application limits.
- Keep Office and document-parsing dependencies patched.
- Delete retained workspaces according to your data-retention policy.
