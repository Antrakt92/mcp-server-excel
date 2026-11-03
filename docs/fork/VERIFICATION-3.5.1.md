# Fork 3.5.1 verification

This maintenance release updates ModelContextProtocol 2.3.0 to 2.3.1 and
video development tooling. No application behaviour or public schema changes.

## Source and dependencies

Build clean committed source using SDK 10.0.403 with zero warnings and errors.
Run the Commit Excel-free CLI/Core/MCP/Service groups and NuGet vulnerability
audit. Bind results to the tested source and verify that the changed runtime
inputs match the release commit. Run the fork packaging checks after updating
release metadata. Preserve all failed or interrupted receipts separately.

The video graph requires a clean lockfile install and npm audit. Check the
resolved HyperFrames and undici versions and exercise a local HTTP request.
Run the video build and check commands; absent original footage blocks full
render acceptance and must remain explicit in the release notes.

## Portable packages and publication

Build both portable packages with `scripts/Build-ForkRelease.ps1 -Version 3.5.1`
from the final clean commit. Verify source commit/tree, package SHA-256 hashes,
PE architecture and executable version metadata.

Run `scripts/tests/fork_protocol_smoke.py` against the extracted MCP executable
with a 15-second idle observation. Run `scripts/tests/fork_packaged_excel_smoke.py`
sequentially against both extracted executables, using only disposable synthetic
workbooks. Require all three protocol cases and both native profiles, correct
workbook state, and normal owned cleanup. Recheck executable hashes afterwards.

Publish both ZIPs, RELEASE-MANIFEST.json, VERIFICATION-RESULTS.json and SHA256SUMS
together. Download and compare every draft asset before publication and every
published asset afterwards. Hosted workflows remain disabled.

## Local installation and limits

Install the exact published packages, verify bytes, update the existing MCP
configuration and generation-specific CLI pipe, and check a fresh configured
MCP process plus the normal CLI wrapper. Inspect pending work before retiring
superseded owned hosts. Recycle old installations and redundant task builds;
preserve source, unique evidence and unsaved work. Existing chat reconnection
is a separate activation boundary.

Unsigned Windows x64 packages require desktop Excel and an interactive session.
Synthetic checks do not establish customer workbook calculations, print parity,
every Office configuration, long-duration stability or historical RPC causes.
