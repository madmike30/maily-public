# Maily 1.3.2 public preview

This update makes composing and refining a reply clearer, protects the current sender identity and existing signature, and reduces model-discovery waiting.

- Clearer Apply, refinement, summary-to-reply and Settings flows, with narrow-pane and keyboard improvements.
- Incoming senders and stale saved names no longer override the current reply author. First-name-only signatures are preserved, including with custom prompts.
- Automatic language can be restored after a manual choice.
- Model discovery warms up in the background and refreshes without unnecessary waiting. The default model and bundled runtime remain unchanged; no paid speed tier is enabled.
- Clearer installer completion, removal and connection guidance.

Download and extract the ZIP, then run MailyOutlookAddinSetup.exe. Follow its connection guidance; Microsoft policy may require an administrator. Existing installations use the upgrade flow. Review generated text before applying or sending it. Uninstall through Windows Installed apps; mailbox deployment may require separate removal through Microsoft.

This unsigned experimental preview is not a Microsoft or OpenAI product. You are responsible for reviewing output and complying with your organization's policies. See the bundled DISCLAIMER.md and the public [privacy policy](https://github.com/madmike30/maily-public/blob/main/PRIVACY.md) and [verification record](https://github.com/madmike30/maily-public/blob/main/docs/VERIFICATION.md) for scope and limitations. Fresh-account registration and all live Outlook rendering variants are not certified. Version 1.3.2 was not installed over the maintainer's active Outlook installation during release validation. The previously documented unsupported hosted-package uninstall limitation remains; this release distributes the local Windows package.
