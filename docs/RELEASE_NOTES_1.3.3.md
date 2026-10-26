# Maily 1.3.3 public preview

A smaller, self-contained installer with the same application features and AI runtime as 1.3.2. No separate .NET, Node.js or Codex installation is needed.

- Identical backend/supervisor runtime files are carried once in the download. Setup verifies and restores the supervisor files automatically before starting the app.
- Development JavaScript source maps are excluded; executable connector code, resources and licenses remain included.
- Existing installation paths, settings preservation and rollback behavior remain unchanged. Download size is reduced; installed runtime copies are retained for compatibility.

Extract the ZIP and run MailyOutlookAddinSetup.exe. This is an unsigned experimental preview. Microsoft registration may require administrator deployment and remains separate from local installation. Review generated content; use is your responsibility. See the bundled DISCLAIMER.md, [privacy policy](https://github.com/madmike30/maily-public/blob/main/PRIVACY.md), and [verification record](https://github.com/madmike30/maily-public/blob/main/docs/VERIFICATION.md).

No live Outlook installation or mailbox operation was performed for this release. Fresh-account registration and every Outlook rendering variant are not certified. The previously documented unsupported hosted-package uninstall limitation remains; this release distributes the local Windows package.