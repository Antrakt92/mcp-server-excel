# Fork 5.3.0 verification

Use the ZIPs, `SHA256SUMS`, `RELEASE-MANIFEST.json` and
`VERIFICATION-RESULTS.json` from the same GitHub release. The manifest binds
the packages to the committed source and embedded executable versions. The
verification receipt records checks against those package and executable hashes.

## Changed behavior

Run the Core `WorkbookPackageComparisonTests` and
`WorkbookPackageStructureTests`, CLI `InProcessCliCommandTests` and MCP
`ProgramTestTransportLifecycleTests`. The structural fixtures cover moved and
renumbered worksheet/chart parts, local filter scope, comments/VML ownership,
parent-style index shifts, changed fonts, duplicate candidates and Strict OOXML.
They also cover invalid targets, missing owners, XML/depth/size limits and
external-target redaction. The MCP boundary must classify malformed structural
XML as `InvalidInput` without creating a workbook session.

Run the generated contract checks after the Release build. Discover MCP
`file_read` and check its `include_structure` parameter. CLI help must expose
`session compare-packages --include-structure`. Batch JSON uses
`includeStructure` and accepts the same boolean setting.

## Packaged acceptance

Run the repository's `scripts/tests/fork_protocol_smoke.py` against the
extracted MCP executable and `fork_packaged_excel_smoke.py` against both
extracted entry points. Retain normal host and owned Excel shutdown evidence.
Use the VBA authentication smoke when the release verification route includes
that coverage. Keep its synthetic inputs and private diagnostics local.

Exercise the new option through the packaged MCP and CLI on a saved synthetic
workbook. Check exact sheet identities, relationships, name scopes and style
mapping results. A comparison of identical saved inputs must report no byte
changes. Hash input files before and after to verify read-only behavior.
Test a malformed structural XML part and an exceeded bound as explicit failures.

Working workbook acceptance may run locally with user authorization. Retain
only aggregate counts and outcomes in shared verification evidence. Workbook
parts, cell values, private names and paths must stay out of release assets.

The structural report describes saved metadata. Native Excel opening, VBA
behavior, calculation and rendering require their own checks.
