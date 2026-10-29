# Changelog

## 1.3.4.0 — release candidate (unreleased)

- Default package reuses the installed ChatGPT desktop app runtime and sign-in; no bundled Codex runtime or ChatGPT installer.
- Verify app/account readiness and rediscover updated runtime paths. Preserve shared app installation and authentication on Maily removal.
- Allow compatible version-only local upgrades while comparing complete manifest permissions, commands and endpoints; retain strict installed-version verification.
- Correct hosted connector cleanup and exclude inherited API-key variables from AI subprocesses.
- Keep summary language anchored to incoming email when an existing reply uses another language.


## 1.3.3.0 — public preview (2026-10-26)

- Reduce download size by storing identical backend/supervisor runtime files once and restoring verified copies during setup.
- Exclude connector development source maps while retaining executable code, resources and licenses.
- Preserve self-contained installation, existing application paths and settings/rollback behavior.

## 1.3.2.0 — public preview (2026-10-26)

- Refined compose, summary, settings, search and briefing hierarchy, including explicit Automatic-language recovery, clearer actions, keyboard focus and narrow-pane layouts.
- Corrected reply-author selection and preservation of first-name-only signatures. Custom prompt templates retain author/signature safeguards.
- Clarified installer completion/removal states and connection guidance.
- Reduced model-discovery overhead through startup warmup and bounded background refresh, with cancellation and cleanup at shutdown. The shipped model/runtime and paid speed tier are unchanged.
- Verified 56 synthetic gate checks, 10 narrow-pane scenarios and three synthetic real-model identity cases. Included in the 1.3.2 installer.

## 1.3.1.0 — public preview

Initial sanitized public snapshot. Includes review-before-apply, focused refinements, guarded Undo, summary-to-reply, language/typography preservation, optional onboarding/style learning, bundled installation runtimes, scoped cache/lifecycle handling and security/privacy hardening.

See docs/RELEASE_NOTES_1.3.1.md and the published release verification record. Private development history and screenshots are intentionally excluded.
