# Verification scope

The release workflow runs the full local synthetic suite, fresh review, a pinned-runtime/package inventory check, and an authorized local install/uninstall/reinstall cycle. Release-specific counts, checksums and status are recorded with the published release after those gates complete.

Covered behavior includes draft identity/body safety, signature/history preservation, automatic language, font handling, summary handoff, comparison/refinements, Undo, privacy-safe diagnostics, runtime tool restrictions, package contents and mocked installation failure/recovery paths.

Local lifecycle evidence establishes behavior on one tested Windows/classic-Outlook installation. It does not prove fresh-account Microsoft consent/registration across tenants or every Outlook visual rendering variant. No real mailbox data is used in the synthetic suite; no test sends messages.

No unsigned artifact can be described as perfectly safe. Use the checksums, read the limitations, review output, and report problems without private data.
