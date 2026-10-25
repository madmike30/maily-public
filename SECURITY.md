# Security policy

This experimental public preview is supported on classic Outlook/Windows only. It has not been independently certified and does not promise freedom from vulnerabilities.

## Report privately

Use GitHub's private vulnerability reporting on this repository when available. Do not post exploit details containing private information, mailbox contents, credentials, tokens or logs in public issues. If private reporting is unavailable, open a minimal issue asking for a private reporting channel without sensitive details.

## Boundaries

The supported package serves the add-in over loopback HTTPS. Browser-origin/hostname checks protect the unauthenticated local surface; configured API-key deployments require a key for protected API routes. This is not protection against malware or another process running as the same Windows user.

AI input is untrusted. Generation uses the restricted runtime configuration: no shell/browser/computer/plugin tools, no inherited user instructions, ephemeral execution, and read-only sandboxing as defense in depth. Review generated content; prompt injection can still affect text and conclusions.

Maily checks the target item and body before writes. It never sends messages automatically. Undo refuses later edits and ambiguous writes. Source formatting is copied only through a limited font allowlist. Credentials are not shipped or stored in pane browser storage.

The unsigned installer checks package hashes and uses normal Windows certificate approval. Hashes do not replace code signing. Use the supported installer, not legacy/developer deployment helpers. Fresh Microsoft registration depends on account/tenant policy and remains a preview limitation.

## Development

Use synthetic messages and mocked Office.js/Graph/COM. Never commit secrets, private account files or real-mail screenshots. Run the full local gate and inspect its output before proposing a release. Hosted CI is opt-in and manual.
