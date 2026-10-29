# Build and verification

The main components are the browser-based Office.js panes, the ASP.NET local backend, the generation library, Windows setup/supervisor, and bounded Outlook bridge scripts. Maily bundles its own application runtimes; the default package requires an installed, signed-in ChatGPT desktop app with its Codex runtime. Developers can opt into the legacy bundled-runtime build with -AiRuntimeMode Bundled; developers need the pinned .NET SDK in global.json, PowerShell7, Node.js, Git, and Chrome or Edge for browser harnesses.

From the repository root:

```powershell
pwsh -NoProfile -File scripts/test-fix-maily.ps1 -DotnetPath <path-to-dotnet.exe>
pwsh -NoProfile -File scripts/build-maily-outlook-installer.ps1 -OutputRoot <new-output-directory>
git diff --check
```

The build produces the ZIP and SHA256 sidecar plus a per-file inventory. It downloads hash-pinned third-party runtime components and includes their notices. Always choose a fresh output directory. Never include local settings, certificate private keys, authentication files, build logs or private screenshots in a release.

The default verification suite uses synthetic data and mocked Office.js/Graph/COM. It needs no mailbox or AI sign-in. Optional real-model verification must use synthetic inputs and an authorized account; live mailbox/installation tests need separate explicit authorization.

Automation is manual and gated by MAILY_HOSTED_CI_ENABLED. The release process does not require hosted CI or spend money automatically. Local full-suite output and public artifact checksums are the release evidence.

The hosted configuration is an advanced option requiring its own deployment/security review; the supported public download is the local Windows package. Public export scripts are maintainer tools and must not be used to expose private history.
