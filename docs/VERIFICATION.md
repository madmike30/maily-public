# Verification scope

The release workflow runs the full local synthetic suite, fresh review and pinned-runtime/package inventory checks. Live lifecycle validation is version-specific: the install/uninstall/reinstall evidence below applies to 1.3.1 only. Version 1.3.2 leaves the active installation untouched.

Covered behavior includes draft identity/body safety, signature/history preservation, automatic language, font handling, summary handoff, comparison/refinements, Undo, privacy-safe diagnostics, runtime tool restrictions, package contents and mocked installation failure/recovery paths.

Local lifecycle evidence establishes behavior on one tested Windows/classic-Outlook installation. It does not prove fresh-account Microsoft consent/registration across tenants or every Outlook visual rendering variant. No real mailbox data is used in the synthetic suite; no test sends messages.

No unsigned artifact can be described as perfectly safe. Use the checksums, read the limitations, review output, and report problems without private data.

## Version 1.3.1.0 evidence

- Maintainer checkout: 56 synthetic checks passed, zero failed; merge gate passed. Independent release/privacy review found no remaining blocking findings.
- Independent public checkout: all 54 application checks passed. Its two Git checks were rerun separately after correcting the test environment's checkout-ownership restriction; whitespace and all 199 text files passed. A pre-existing project-file BOM was removed and the affected project rebuilt successfully.
- Packaged runtime: seven checks passed using synthetic English/Hebrew inputs, including current-draft meaning and language selection.
- Local uninstall/reinstall: application removal verified, settings/logs preserved, reinstall had zero failed steps, and post-restart activation verification returned 0. Installed files matched the audited package and current assets were served without caching.
- No existing mail was read and no messages were sent during this release validation.
- Package inventory: 952 file hashes verified. ZIP SHA256: `06D47251A94EA7915F29F6FE5DD66CBC8E71CB632E58D8AF71D546CEF4705A3A`.
- Public source and all reachable clean history were checked for private data; the original development history was not imported. Build logs, local profiles, authentication files, screenshots and private-key files are excluded.

The unsupported hosted-package uninstall path has a known cleanup defect that can prevent reinstalling over its leftover connector directories. This does not affect the local Windows package published here. Fresh-account registration and live compose visual checks remain open as described above.

## Version 1.3.2.0 evidence

- Packaged runtime: seven checks passed with isolated synthetic data, covering authentication/catalog, English/Hebrew generation, existing-draft commitments and automatic language. No Outlook access. An initial catalog check was unavailable during concurrent builds; the preserved follow-up after builds completed passed, with catalog response in 989ms. This does not guarantee latency under every workload.
- Independent source/package review: GO, no remaining findings. All 952 inventory hashes and sizes verified; the final ZIP contains 953 files and 26 directory entries. Executable bytes match the package tested above; only release-note links were corrected afterward.
- Initial full synthetic run: 55 passed, one test-helper file-sharing race failed. Atomic readiness publication and cleanup fixed that race; all five targeted lifecycle cases passed. No assertions were weakened.
- ZIP SHA256: `f72cc8cdf7b8cf49023379f20fa068f2a5a363489f8fbb1ad38d49206dbb3920`.
- No live install/uninstall, mailbox access or sends were performed for this version. The 1.3.1 lifecycle evidence remains historical, not a claim about a new live 1.3.2 installation.
- Final complete local synthetic rerun: **56 passed, 0 failed**. All 205 public text files passed strict UTF-8/no-BOM checks. Hosted CI remains disabled; no paid Actions run was used.

## Version 1.3.3.0 evidence

- ZIP: 275,580,679 bytes (262.8 MiB), down from 319,219,299 bytes: 43,638,620 bytes / 13.7% smaller. This is a download optimization; installed runtime copies are retained for compatibility.
- Final local suite: 58 passed, 0 failed; merge gate PASS; independent code/privacy review GO.
- Actual compact package reconstructed all 188 shared DLLs (80,167,952 bytes) with matching hashes. All 763 inventory hashes and 764 archived files verified. Synthetic map tests cover tampering, unsafe paths, collisions, legacy packages and rollback boundaries on both PowerShell versions.
- Connector executable/resource/license bytes are unchanged; only three development source maps were omitted. Isolated connector startup and integrity checks passed.
- Seven packaged-runtime checks passed with synthetic inputs and no Outlook access. Initial cold catalog discovery was unavailable during verification activity; the post-build follow-up was ready on its first attempt. Original failure evidence is retained; no universal latency guarantee is claimed.
- SHA256: `382ae730a41c69a600e3fd27a16a120db69a4cd1242d316934098917956463f4`.
- No live installation, uninstall, Outlook shutdown, mailbox action or send was performed. Prior preview limitations still apply.