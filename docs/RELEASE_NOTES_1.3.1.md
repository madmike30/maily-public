# Maily 1.3.1.0 public preview

This first public source snapshot excludes private development history. The Windows x64 package is unsigned and experimental.

## Included

- Current-draft rewriting with a Before/After review and explicit Apply.
- Shorter, Clearer and Warmer refinements.
- Guarded one-step Undo when the same item and exact post-write body can be verified.
- Summary-to-reply previews with individually selectable action items.
- Automatic email/draft language selection and explicit overrides.
- Detectable Outlook font/size preservation and safe fallback.
- Discovered model choices, speed-oriented Automatic, and optional writing-style learning.
- A replayable Settings tour and one installer entry point with integrity checks, normal Outlook closure, scoped cache refresh, upgrade/rollback support and uninstall guidance.
- Restricted AI runtime tools and local browser-origin protection.

## Important limits

Classic Outlook on Windows only. Fresh-account Microsoft sign-in/registration has not been verified across organizations and may require an administrator. Personal Microsoft account compatibility is not assumed. Local cache cleanup cannot create a mailbox registration. Do not interpret a healthy backend as proof of visible Outlook commands.

A fresh reinstall needs normal Windows localhost-certificate trust approval. Cancellation and active inline replies stop setup safely. Shared account sign-in and settings/logs are retained by default. Local removal cannot remove an organization-managed mailbox add-in.

Complex source typography can fall back to an Office font. Undo is withheld for native/ambiguous/normalized writes that cannot be verified exactly. AI outputs remain fallible and account-dependent. Synthetic Office tests do not certify every live Outlook version.

See VERIFICATION.md for the tested release boundary and INSTALLER.md for setup/recovery.
