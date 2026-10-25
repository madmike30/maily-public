# Verification scope

The release workflow runs the full local synthetic suite, fresh review, a pinned-runtime/package inventory check, and an authorized local install/uninstall/reinstall cycle. Release-specific counts, checksums and status are recorded with the published release after those gates complete.

Covered behavior includes draft identity/body safety, signature/history preservation, automatic language, font handling, summary handoff, comparison/refinements, Undo, privacy-safe diagnostics, runtime tool restrictions, package contents and mocked installation failure/recovery paths.

Local lifecycle evidence establishes behavior on one tested Windows/classic-Outlook installation. It does not prove fresh-account Microsoft consent/registration across tenants or every Outlook visual rendering variant. No real mailbox data is used in the synthetic suite; no test sends messages.

No unsigned artifact can be described as perfectly safe. Use the checksums, read the limitations, review output, and report problems without private data.

## Version 1.3.1.0 evidence

- Maintainer checkout: 56 synthetic checks passed, zero failed; merge gate passed. Independent release/privacy review found no remaining blocking findings.
- Packaged runtime: seven checks passed using synthetic English/Hebrew inputs, including current-draft meaning and language selection.
- Local uninstall/reinstall: application removal verified, settings/logs preserved, reinstall had zero failed steps, and post-restart activation verification returned 0. Installed files matched the audited package and current assets were served without caching.
- No existing mail was read and no messages were sent during this release validation.
- Package inventory: 952 file hashes verified. ZIP SHA256: `06D47251A94EA7915F29F6FE5DD66CBC8E71CB632E58D8AF71D546CEF4705A3A`.
- Public source and all reachable clean history were checked for private data; the original development history was not imported. Build logs, local profiles, authentication files, screenshots and private-key files are excluded.

The unsupported hosted-package uninstall path has a known cleanup defect that can prevent reinstalling over its leftover connector directories. This does not affect the local Windows package published here. Fresh-account registration and live compose visual checks remain open as described above.
