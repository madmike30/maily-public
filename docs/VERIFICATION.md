# Verification scope

The release workflow runs the full local synthetic suite, fresh review, existing desktop-runtime checks and package inventory verification. Evidence below is version-specific; historical lifecycle checks do not establish that a later release is installed or verified.

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

Historical limitation: the hosted-package uninstaller left connector directories behind in these releases. The 1.3.4 candidate fixes that cleanup with a synthetic regression; hosted deployment remains outside the supported local Windows release. Fresh-account registration and live compose visual checks are subject to the version-specific evidence below.

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
## Version 1.3.4.0 candidate evidence (2026-10-29)

- Sanitized source build completed. Full local synthetic suite: **62 passed, 0 failed**. Fresh independent code review of implementation ce96353: **GO**, with no confirmed findings. The complete merge gate also passed. Live lifecycle checks must complete before final release promotion.
- Exact package: **8 synthetic runtime checks passed**, using the installed desktop runtime and existing ChatGPT sign-in. Checks cover readiness, model discovery, English/Hebrew generation, existing-draft commitments, reply language, incoming-email summary language, and a synthetic file-reading canary. This canary is a bounded regression, not proof against every prompt injection.
- Default package includes neither a ChatGPT installer nor a Codex executable. **757 inventory hashes and 758 archived files** verified. ZIP size: **166,133,072 bytes**. SHA256: `6b1a7a9902a7684e7ac196c892d8f4c57b3e6a7e0427e2e86ac9c225995a99c3`.
- Synthetic upgrade checks cover older compatible local manifests and reject changes to permissions, commands, requirements, bindings and endpoints, downgrade conflicts, hosted version mismatches and uncertain ownership. Installed-package comparison remains version-strict. Hosted removal regression preserves settings/logs and removes the owned connector payload.
- Synthetic subprocess tests verify that API-key environment variables are excluded from AI children without changing the parent environment or configured authentication home. No real API keys are used in these tests.
- Public export dependency check passed; **215 text files** passed UTF-8/no-BOM and targeted private-path/key-marker checks. The original repository and history remain private; hosted Actions are disabled.
- Live recovery installation is pending Windows certificate consent. Do not interpret synthetic results as completed installation, fresh-account tenant validation, or real-mail rendering certification. No existing mailbox messages have been read or sent during this release work.