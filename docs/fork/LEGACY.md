# Historical 2.4.1 source recovery

The release also preserves `ExcelMcp-Antrakt-Legacy-2.4.1.patch`, the cumulative
local lifecycle and performance source patch developed before the 2.5.3 rebase.
Its SHA-256 is `c04150bde111d608828b73c7c0c37c6e9b04b610e22991594916594185299db0`.

This is archival source, not the recommended runtime. It applies only to upstream
v2.4.1, commit `0576493eebbd1eec8d968b6dab92606d8d5f2a69`. In an isolated checkout
of that revision, use `git apply --check <patch>` before `git apply <patch>` and
follow that revision's pinned SDK, build and verification instructions. It was
developed and checked locally with SDK 10.0.402 under its allowed SDK roll-forward.

Do not apply it to 2.5.x: upstream changed cancellation and workbook ownership.
The archival patch contains older cleanup and forced-close behavior. The current
fork instead preserves 2.5.3 unsaved-workbook handover and cancellation recovery.
Its old tests and timing results do not establish current-release acceptance.

The patch contains source and tests only, under the original MIT license. It
excludes workbooks, customer data, personal configuration and local evidence logs.
