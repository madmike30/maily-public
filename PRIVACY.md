# Privacy and data flow

Maily is a local Outlook add-in with a local backend. It is not an offline AI model. Text required for a requested reply, summary, search, briefing or style-learning operation is sent to the user's signed-in AI provider through the existing desktop app runtime. Provider account policies govern processing and retention; Maily does not promise zero retention by that provider.

The feature may use message body, subject, participants, selected search results and the current draft. Search and briefing can process more than one message when explicitly requested. Only use authorized data. Do not use a personal AI account for organization data without permission.

Settings and learned writing-style information are saved in the per-user application data directory. Undo and comparison text remain in the active pane's memory, are not saved to browser storage and are discarded with the pane. The quick tour stores only a non-sensitive completion flag. Normal diagnostics are designed to omit message bodies and credentials, but you must review and redact any log before sharing it.

Style learning is optional and reads sent messages only after an explicit start. You may leave personalization disabled. Removing the add-in preserves settings/logs by default, and does not remove shared AI sign-in. Review retained data locally if you need deletion; deletion must be a deliberate user action.

Release archives contain application files, templates, synthetic fixtures and third-party notices, not the maintainer's account state or mailbox data. The public source history starts from a sanitized snapshot. No private development screenshots or prior private history are published.
