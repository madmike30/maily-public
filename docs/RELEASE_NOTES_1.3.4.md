# Maily 1.3.4 release candidate

Maily now checks and reuses the runtime and ChatGPT sign-in already supplied by the installed ChatGPT desktop app with Codex features. The default download contains neither a ChatGPT installer nor another Codex runtime.

- Install and sign in to the desktop app first. Setup checks actual runtime/account readiness; merely finding an app folder does not count as success.
- If the app or sign-in is missing, setup explains how to open/update/sign in to the desktop app and retry. Maily does not install or sign out of that app.
- Runtime discovery is refreshed so supported desktop-app updates do not leave Maily pointing to a removed version directory.
- Maily still bundles its own application runtimes and Microsoft connection tools. No separate developer SDK, Node.js or command-line installation is required.

Extract the ZIP and run MailyOutlookAddinSetup.exe. Existing settings are preserved by the upgrade flow. Uninstall Maily through Windows Installed apps; the shared desktop app and its sign-in are not removed.

This unsigned experimental preview requires a compatible desktop app exposing its Codex runtime and an eligible ChatGPT account. Older ChatGPT-only clients may require updating. Microsoft mailbox registration is separate and may require an administrator. Review all generated content; use is your responsibility. See DISCLAIMER.md, the [privacy policy](https://github.com/madmike30/maily-public/blob/main/PRIVACY.md), and [verification](https://github.com/madmike30/maily-public/blob/main/docs/VERIFICATION.md). Fresh-account deployment and every Outlook rendering variant are not certified.
## Upgrade and reliability

- Reuse an older local Outlook registration for version-only upgrades when the complete manifest remains compatible. Changed commands, permissions, endpoints and unknown ownership still require repair; downgrades are not silently accepted.
- Keep the existing reply authoritative, preserve signatures and formatting, and require review before applying changes. Summary language follows the incoming email even when a reply uses another language.
- Hosted package removal now cleans up its connector files. The distributed and supported package remains local Windows/classic Outlook.

## Acceptance walkthrough

After setup completes and Outlook restarts, open a synthetic draft. Confirm Maily appears, check the available model list, generate a reply from existing draft text, review the comparison, apply it, and verify that your signature and font remain intact. Test Undo before making further edits. Test Summary with a synthetic message in another language and prepare a reply from a selected action item. Nothing should be sent automatically.

If setup reports a registration conflict, stop and follow the supported repair instructions. If generation or insertion fails, retain your draft and report the redacted error; do not repeatedly apply an uncertain write. Keep the previous package and setup backup until the upgrade is verified. Reinstall the previous version only if its registration remains compatible; a manifest downgrade may require deliberate registration repair.