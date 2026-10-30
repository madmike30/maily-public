# Maily

Release 1.3.4 reuses the installed ChatGPT desktop app runtime and sign-in instead of bundling another Codex runtime; see [CHANGELOG](CHANGELOG.md).

An open-source AI drafting assistant for **classic Outlook on Windows**. Generate a reply, compare it with your current draft, refine it, and apply it only when you are ready. Maily never sends mail automatically.

**Release 1.3.4.0.** This is an unsigned release, not a Microsoft or OpenAI product. Fresh-account Microsoft registration has not been verified across organizations. Organization policy may require administrator deployment. Read the [limitations](docs/RELEASE_NOTES_1.3.4.md) before installing.

## Install only Maily

Download the ZIP and checksum from [Releases](https://github.com/madmike30/maily-public/releases). Extract it and run `MailyOutlookAddinSetup.exe`. You need classic Outlook and a compatible ChatGPT desktop app with Codex features, installed and signed in to an eligible account. The package includes its application runtimes, so you do not need to install .NET, Node.js or a separate Codex CLI.

Setup checks and reuses the desktop app runtime and sign-in. If unavailable, it asks you to open or update that app, sign in and retry. Approve the Windows localhost-certificate prompt only when you initiated this installation. Microsoft connection is a separate explicit step. [Installation, upgrade and uninstall](docs/INSTALLER.md).

## Use it

- **Reply:** your existing draft is authoritative. Review the Before/After preview, choose Shorter/Clearer/Warmer if useful, then Apply. Undo is offered only when the same draft has not changed since the last verified write.
- **Summary:** follows the incoming email language. Select relevant action items and prepare a reply preview.
- **Formatting:** preserves detectable Outlook font and size. Replies follow your draft language, otherwise the email; explicit overrides remain available.
- **Settings:** choose tone, length and an available model, or keep Automatic. The optional quick tour explains the workflow and writing-style learning.
- **Style learning:** reads sent emails only when explicitly started. It is optional.

AI may be wrong. Review every result, never assume commitments or facts are correct, and use only data you are authorized to process. [Responsibility disclaimer](docs/DISCLAIMER.md) · [Privacy](PRIVACY.md) · [Security](SECURITY.md).

## Build and contribute

See [development and testing](docs/DEVELOPMENT.md), [contributing](CONTRIBUTING.md), [verification](docs/VERIFICATION.md), and [changelog](CHANGELOG.md). Test data must be synthetic. Do not post mailbox messages, credentials, screenshots of real mail, or unredacted logs in issues.

Maily is MIT licensed; bundled third-party components retain their own licenses and notices.
