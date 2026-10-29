# Installation, upgrade and removal

## Requirements

Windows x64, classic Outlook desktop signed in, internet access and the ChatGPT desktop app with Codex features installed and signed in to an eligible account. Obtain organizational permission before processing work email. New Outlook, macOS and Outlook on the web are not supported by this preview.

Use only the ZIP from this repository's Releases page. Compare its SHA256 with the published checksum (`Get-FileHash <download.zip> -Algorithm SHA256`). The release is unsigned; checksum matching establishes artifact consistency, not publisher identity or a guarantee of safety.

## Install

1. Save drafts and finish or close inline replies.
2. Extract the ZIP. Run its single `MailyOutlookAddinSetup.exe` entry point.
3. Read and accept the responsibility notice. Setup verifies the package, existing desktop runtime and account readiness. It does not install ChatGPT or another Codex runtime.
4. Allow normal Outlook closure. A cancelled save prompt or an active inline reply stops setup safely; it never force-kills Outlook.
5. Approve the Windows localhost HTTPS certificate trust prompt when you initiated setup. This allows Outlook to display Maily's local page. Cancelling stops installation; retry setup when ready.
6. If requested, choose **Sign in and connect**. This requires a supported Microsoft 365 organizational account and may need administrator approval. Setup does not bypass tenant policy. Fresh-account registration is not yet verified across tenants; use the displayed supported fallback/administrator route if needed.
7. Open Outlook and check connection. Open a draft and find Maily. In Maily Settings, select **Show quick tour**. Style learning remains optional and must be started explicitly.

With classic Outlook and a compatible signed-in ChatGPT/Codex desktop app already installed, install only Maily. The default per-user directory is `%LOCALAPPDATA%\Maily\OutlookAddin`. The package includes .NET and Microsoft connection tooling, but no AI runtime or ChatGPT installer. Setup verifies the existing desktop runtime and sign-in; credentials are not included in downloads.

## Upgrade and recovery

Run the new extracted setup. It detects an existing copy, verifies files, backs up its managed application files and preserves settings. If replacement fails, it attempts rollback and reports the result. Keep the previous ZIP and all setup logs until recovery is complete. Only one setup/uninstall may run at a time.

A local version-only upgrade reuses the existing Outlook registration only when the complete manifest remains compatible, including permissions, commands and endpoints. Changed permissions or endpoints, a newer registered version, or uncertain ownership require repair or administrator help. Setup never clears a conflicting registration to hide the mismatch.

Frontend cache refresh is scoped to Maily and requires Outlook to close. It cannot install a mailbox add-in or override a conflicting organization-managed copy. Setup distinguishes local application readiness from Outlook registration. Restart Outlook and use Check connection when requested.

## Uninstall and reinstall

Use Windows Settings > Apps > Installed apps > Maily Outlook Add-in > Uninstall. Save drafts first. Settings and logs are retained by default; shared ChatGPT/Codex sign-in is not removed. Local uninstall cannot remove a mailbox-installed add-in. Follow the final Outlook removal guidance if you want that removed; administrator-managed copies require the administrator.

For reinstall, extract the release and run setup again. A fresh certificate may require Windows trust approval again. A silent reinstall is not guaranteed to be unattended. Uninstall does not undo sent mail or previous edits to drafts.

## Troubleshooting

- **Cannot close Outlook:** finish inline replies, answer save prompts, close other setup windows, then retry. Never force-terminate Outlook to bypass this guard.
- **Maily missing:** local installation is not mailbox registration. Use Check connection, then the supported Sign in and connect or administrator instructions. Cache cleanup alone cannot register the add-in.
- **Sign-in/model unavailable:** open/update the ChatGPT desktop app, sign in there, retry setup and refresh the model list in Settings. Availability depends on the account. Prefer Automatic unless you need a specific supported model.
- **Undo unavailable:** it is intentionally withheld if the target/body changed or the original write cannot be verified exactly. Do not overwrite newer edits.
- **Generation failed:** your compose text remains available. Retry after checking account readiness; inspect setup's collapsed technical details locally if necessary.

For maintainers, exit21 requests post-restart verification; exit22 means local removal completed but mailbox removal is unverified; exit1618 means another setup owns the operation. Report only redacted error codes and synthetic reproduction steps.
