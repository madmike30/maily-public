# Pane harness fixtures

These scripts contain synthetic Office.js/backend responses and behavioral assertions. They contain no live mailbox credentials or messages. The generator combines them with the current product HTML, JavaScript and CSS on every run; product markup is not copied here.

Run `pwsh -NoProfile -File scripts/test-outlook-addin-harness.ps1` locally with Chrome/Edge. `-BrowserPath` or `MAILY_TEST_BROWSER` can select another Chromium executable. Every run retains generated pages, unique browser profiles, DOM, screenshots and logs. A restrictive page CSP blocks network connections and external resources even if a mock is missing.

The runner checks English/Hebrew compose behavior, summary compose/read surfaces, settings, search refinement and daily briefing. Generated pages, DOM, screenshots and results go under `output/playwright/addin-harness`. This mocks AI and Outlook; use the separate live tests for actual account/model and classic Outlook behavior.

The failed-regeneration scenario verifies recovery after a mocked HTTP 503 generation failure: the prior completed preview is restored, compose stays unchanged by the failed request, and Insert again inserts the prior complete draft. Run just this case with `-CaseName taskpane-failed-regenerate`; filtered runs write a separate case summary and preserve the full-suite summary.
