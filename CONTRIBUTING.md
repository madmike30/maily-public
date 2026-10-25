# Contributing

Use a focused branch, a regression test for behavior changes, and a concise explanation. Run the local verification command in docs/DEVELOPMENT.md. Preserve failure evidence locally; never upload unredacted logs or real mailbox data.

Do not introduce automatic sends, silent overwrites of newer drafts, persistent browser copies of mail, broad filesystem tools for AI generation, or account credentials in packages. Respect existing signatures/history and language choices.

For UI changes, test the actual pane markup using the synthetic browser harness and refresh frontend asset revisions. For installer changes, use mocked lifecycle tests first, then an explicitly authorized local install cycle with settings/log preservation.

Report vulnerabilities using SECURITY.md. Keep public issues synthetic and reproducible. Contributions retain the project's existing MIT license; bundled dependencies retain their own terms.
