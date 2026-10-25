using System.Diagnostics;
using System.Drawing;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;

namespace Maily.OutlookAddin.Setup;

internal static class Program
{
	private const string DefaultInstallSubPath = @"Maily\OutlookAddin";
	private const string SetupFileName = "MailyOutlookAddinSetup.exe";

	[STAThread]
	private static int Main(string[] args)
	{
		var options = SetupOptions.Parse(args);
		if (!options.Uninstall && IsInstalledCopy())
		{
			options = options with
			{
				Uninstall = true,
				InstallRoot = Path.GetFullPath(AppContext.BaseDirectory).TrimEnd('\\')
			};
		}

		if (options.Uninstall && !options.FromTemp && IsInstalledCopy())
		{
			LaunchTempUninstaller(options);
			return 0;
		}

		WaitForParentExit(options);
		SetupInstanceLease? lease;
		try { lease = SetupInstanceLease.TryAcquire(SetupInstanceLease.CurrentUserName()); }
		catch
		{
			return ReportOwnershipFailure(options, "Maily setup could not establish exclusive access. Close other Maily setup windows and retry.", 1);
		}
		if (lease is null)
			return ReportOwnershipFailure(options, "Another Maily installation or uninstall is already running. Finish or close that setup window, then retry.", 1618);

		// One per-user owner covers every install root: startup, registry and ports
		// are shared. Acquire after the temporary uninstaller's parent has exited.
		using (lease)
		{
			if (options.Quiet)
				return ScriptRunner.Run(options, _ => { }).ExitCode;

			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);
			Application.Run(new SetupForm(options));
			return 0;
		}
	}

	private static int ReportOwnershipFailure(SetupOptions options, string message, int exitCode)
	{
		if (options.Quiet) Console.Error.WriteLine(message);
		else MessageBox.Show(message, "Maily setup is unavailable", MessageBoxButtons.OK, MessageBoxIcon.Information);
		return exitCode;
	}

	private static bool IsInstalledCopy() =>
		(File.Exists(Path.Combine(AppContext.BaseDirectory, "setup", "uninstall.ps1")) ||
			File.Exists(Path.Combine(AppContext.BaseDirectory, "uninstall.ps1"))) &&
		Path.GetFullPath(AppContext.BaseDirectory)
			.StartsWith(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), DefaultInstallSubPath), StringComparison.OrdinalIgnoreCase);

	private static void LaunchTempUninstaller(SetupOptions options)
	{
		var tempRoot = Path.Combine(Path.GetTempPath(), "MailyOutlookAddinSetup-" + Guid.NewGuid().ToString("N"));
		Directory.CreateDirectory(tempRoot);
		var tempExe = Path.Combine(tempRoot, SetupFileName);
		File.Copy(Application.ExecutablePath, tempExe, overwrite: true);

		var arguments = new StringBuilder("--uninstall --from-temp");
		arguments.Append(" --install-root ").Append(Quote(Path.GetFullPath(AppContext.BaseDirectory).TrimEnd('\\')));
		arguments.Append(" --package-mode ").Append(PackageModeResolver.Resolve(options));
		arguments.Append(" --parent-pid ").Append(Environment.ProcessId);
		if (options.Quiet)
		{
			arguments.Append(" --quiet");
		}

		if (options.RemoveTrustedCertificate)
		{
			arguments.Append(" --remove-trusted-certificate");
		}

		if (options.CloseOutlook)
		{
			arguments.Append(" --close-outlook");
		}

		Process.Start(new ProcessStartInfo
		{
			FileName = tempExe,
			Arguments = arguments.ToString(),
			UseShellExecute = true,
			WorkingDirectory = tempRoot
		});
	}

	private static void WaitForParentExit(SetupOptions options)
	{
		if (!options.FromTemp || options.ParentProcessId is not { } parentProcessId)
		{
			return;
		}

		try
		{
			using var parent = Process.GetProcessById(parentProcessId);
			parent.WaitForExit(30000);
		}
		catch
		{
			// If the parent process is already gone, uninstall can continue.
		}
	}

	private static string Quote(string value) => "\"" + value.Replace("\"", "\\\"") + "\"";
}

internal sealed record SetupOptions(
	bool Uninstall,
	bool FromTemp,
	bool Quiet,
	string? InstallRoot,
	string PackageMode,
	bool KeepUserData,
	bool KeepLogs,
	bool RemoveTrustedCertificate,
	bool TrustLocalhostRoot,
	bool CloseOutlook,
	int? ParentProcessId,
	bool VerifyOutlookActivation = false)
{
	public static SetupOptions Parse(string[] args)
	{
		var uninstall = Has(args, "--uninstall");
		var fromTemp = Has(args, "--from-temp");
		var quiet = Has(args, "--quiet");
		var installRoot = ValueAfter(args, "--install-root");
		var packageMode = NormalizePackageMode(ValueAfter(args, "--package-mode"));
		var keepUserData = true;
		var keepLogs = true;
		var removeTrustedCertificate = Has(args, "--remove-trusted-certificate");
		var trustLocalhostRoot = Has(args, "--trust-localhost-root");
		var closeOutlook = Has(args, "--close-outlook");
		var parentProcessIdText = ValueAfter(args, "--parent-pid");
		var parentProcessId = int.TryParse(parentProcessIdText, out var parsedParentProcessId)
			? parsedParentProcessId
			: (int?)null;
		return new SetupOptions(uninstall, fromTemp, quiet, installRoot, packageMode, keepUserData, keepLogs, removeTrustedCertificate, trustLocalhostRoot, closeOutlook, parentProcessId, Has(args, "--verify-outlook"));
	}

	private static bool Has(string[] args, string name) =>
		args.Any(arg => string.Equals(arg, name, StringComparison.OrdinalIgnoreCase));

	private static string? ValueAfter(string[] args, string name)
	{
		for (var index = 0; index < args.Length - 1; index++)
		{
			if (string.Equals(args[index], name, StringComparison.OrdinalIgnoreCase))
			{
				return args[index + 1];
			}
		}

		return null;
	}

	private static string NormalizePackageMode(string? value)
	{
		if (string.Equals(value, "Hosted", StringComparison.OrdinalIgnoreCase))
		{
			return "Hosted";
		}

		if (string.Equals(value, "Local", StringComparison.OrdinalIgnoreCase))
		{
			return "Local";
		}

		return "Auto";
	}
}

internal static class PackageModeResolver
{
	public static string Resolve(SetupOptions options)
	{
		if (string.Equals(options.PackageMode, "Hosted", StringComparison.OrdinalIgnoreCase) ||
			string.Equals(options.PackageMode, "Local", StringComparison.OrdinalIgnoreCase))
		{
			return options.PackageMode;
		}

		var root = options.Uninstall && !string.IsNullOrWhiteSpace(options.InstallRoot)
			? Path.GetFullPath(options.InstallRoot)
			: Path.GetFullPath(AppContext.BaseDirectory);
		var modePath = Path.Combine(root, "package-mode.txt");
		if (File.Exists(modePath))
		{
			var mode = File.ReadAllText(modePath).Trim();
			if (string.Equals(mode, "Hosted", StringComparison.OrdinalIgnoreCase))
			{
				return "Hosted";
			}

			if (string.Equals(mode, "Local", StringComparison.OrdinalIgnoreCase))
			{
				return "Local";
			}
		}

		var packagePayloadHost = Path.Combine(root, "payload", "app", "Maily.OutlookAddin.Web", "Maily.OutlookAddin.Web.exe");
		var installedHost = Path.Combine(root, "app", "Maily.OutlookAddin.Web", "Maily.OutlookAddin.Web.exe");
		return File.Exists(packagePayloadHost) || File.Exists(installedHost) ? "Local" : "Hosted";
	}
}

internal sealed class SetupForm : Form
{
	private readonly SetupOptions options;
	private readonly string packageMode;
	private readonly Label statusLabel = new();
	private readonly TextBox detailsBox = new();
	private readonly ProgressBar progressBar = new();
	private readonly Button primaryButton = new();
	private readonly Button closeButton = new();
	private readonly CheckBox keepDataCheckBox = new();
	private readonly CheckBox removeTrustedCertificateCheckBox = new();
	private readonly CheckBox responsibilityCheckBox = new();
	private readonly CancellationTokenSource setupCancellation = new();
	private readonly ListView checks = new();
	private readonly SetupFlow flow = new();
	private bool busy;
	private bool applyingChanges;
	private readonly string? installedVersion;

	public SetupForm(SetupOptions options) : this(options, PackageModeResolver.Resolve(options), DetectInstalledVersion(options))
	{
		if (!options.Uninstall) RestoreConnectionProgress();
	}

	// Explicit inputs let the isolated render harness exercise the real form without host discovery.
	internal SetupForm(SetupOptions options, string knownPackageMode, string? knownInstalledVersion)
	{
		this.options = options;
		packageMode = knownPackageMode;
		installedVersion = knownInstalledVersion;
		ConfigureForm();
		FormClosing += (_, e) => { if (applyingChanges) { e.Cancel = true; return; } setupCancellation.Cancel(); };
	}

	private void RestoreConnectionProgress()
	{
		try
		{
			var root = options.InstallRoot ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Maily", "OutlookAddin");
			var summary = Path.Combine(root, "data", "setup-connection-progress.json");
			if (!File.Exists(summary)) summary = Path.Combine(root, "data", "install-summary.json");
			if (!File.Exists(summary)) return;
			bool Matches(string source, string target) => File.Exists(source) && File.Exists(target) &&
				System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(source)).AsSpan().SequenceEqual(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(target)));
			var matches = Matches(Path.Combine(AppContext.BaseDirectory, "payload", "addin", "manifest.xml"), Path.Combine(root, "addin", "manifest.xml")) &&
				Matches(Application.ExecutablePath, Path.Combine(root, "MailyOutlookAddinSetup.exe"));
			if (!flow.TryResume(File.ReadAllText(summary), root, matches, PackageFingerprint())) return;
			responsibilityCheckBox.Checked = true; // Previous install already required acknowledgment; resumption only verifies.
			statusLabel.Text = flow.Message;
			primaryButton.Text = flow.PrimaryText;
			SetCheck("install", "Complete");
			SetCheck("connect", "Waiting for Outlook");
			progressBar.Value = 66;
		}
		catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException) { }
	}

	private bool IsHostedPackage => string.Equals(packageMode, "Hosted", StringComparison.OrdinalIgnoreCase);

	private void ConfigureForm()
	{
		Text = options.Uninstall ? "Remove Maily Outlook Add-in" : "Maily Outlook Add-in Setup";
		StartPosition = FormStartPosition.CenterScreen;
		FormBorderStyle = FormBorderStyle.FixedDialog;
		MaximizeBox = false;
		MinimizeBox = false;
		ClientSize = new Size(760, 570);
		AutoScaleMode = AutoScaleMode.Font;
		AutoScroll = true;
		Font = new Font("Segoe UI", 10F);
		BackColor = Color.White;

		var title = new Label
		{
			Text = options.Uninstall ? "Remove Maily" : installedVersion is null ? "Welcome to Maily" : "Update Maily",
			Font = new Font("Segoe UI Semibold", 17F, FontStyle.Bold),
			AutoSize = false,
			Location = new Point(28, 22),
			Size = new Size(700, 34)
		};
		Controls.Add(title);

		var subtitle = new Label
		{
			Text = options.Uninstall
				? (IsHostedPackage
					? "Remove Maily from this PC. Its Outlook account entry may need separate removal."
					: "Remove Maily from this PC. Its Outlook account entry may need separate removal.")
				: (IsHostedPackage
					? "A hosted, per-user Outlook desktop registration. No local backend, Codex setup, certificates, or background startup are installed."
					: "Email assistance in classic Outlook. Setup includes everything Maily needs."),
			AutoSize = false,
			Location = new Point(30, 64),
			Size = new Size(700, 42)
		};
		Controls.Add(subtitle);

		var prerequisites = new Label
		{
			Text = options.Uninstall
				? (IsHostedPackage
					? "Your settings, logs and shared sign-in will be kept."
					: "Your settings, logs and shared sign-in will be kept.")
				: (IsHostedPackage
					? "Before installing, make sure Outlook desktop is installed and signed in, and the hosted Maily URL is reachable from this PC."
					: "Your settings and sign-in will be kept. Save open drafts before continuing."),
			AutoSize = false,
			Location = new Point(30, 112),
			Size = new Size(700, 48)
		};
		Controls.Add(prerequisites);

		keepDataCheckBox.Text = "Settings and logs are retained";
		keepDataCheckBox.Location = new Point(30, 160);
		keepDataCheckBox.Size = new Size(240, 28);
		keepDataCheckBox.Visible = false;
        keepDataCheckBox.Checked = true;
		Controls.Add(keepDataCheckBox);

		removeTrustedCertificateCheckBox.Text = "Remove trusted localhost certificate (recommended)";
		removeTrustedCertificateCheckBox.Location = new Point(30, 190);
		removeTrustedCertificateCheckBox.Size = new Size(420, 28);
		removeTrustedCertificateCheckBox.Checked = !IsHostedPackage && (options.RemoveTrustedCertificate || options.Uninstall);
		removeTrustedCertificateCheckBox.Visible = false;
		Controls.Add(removeTrustedCertificateCheckBox);

		responsibilityCheckBox.Text = "I understand AI can make mistakes. I am responsible for reviewing its output and following my organization's data and AI policies.";
		responsibilityCheckBox.Location = new Point(30, 164);
		responsibilityCheckBox.Size = new Size(700, 62);
		responsibilityCheckBox.Visible = !options.Uninstall;
		responsibilityCheckBox.CheckedChanged += (_, _) => primaryButton.Enabled = !busy && (options.Uninstall || responsibilityCheckBox.Checked);
		Controls.Add(responsibilityCheckBox);

		checks.Location = new Point(30, 236);
		checks.Size = new Size(700, 110);
		checks.View = View.Details;
		checks.FullRowSelect = true;
		checks.HeaderStyle = ColumnHeaderStyle.Nonclickable;
		checks.Columns.Add("Setup checks", 485);
		checks.Columns.Add("Status", 190);
		foreach (var row in new[] { ("prepare", "1. Check this PC and your sign-in"), ("install", "2. Install Maily"), ("connect", "3. Connect Outlook") })
			checks.Items.Add(new ListViewItem(new[] { row.Item2, "Waiting" }) { Name = row.Item1 });
		checks.Visible = !options.Uninstall;
		Controls.Add(checks);

		statusLabel.Text = "Ready.";
		statusLabel.AutoSize = false;
		statusLabel.Location = new Point(30, 360);
		statusLabel.Size = new Size(700, 62);
		statusLabel.Font = new Font(Font, FontStyle.Bold);
		Controls.Add(statusLabel);

		progressBar.Location = new Point(30, 430);
		progressBar.Size = new Size(700, 12);
		progressBar.Style = ProgressBarStyle.Continuous;
		Controls.Add(progressBar);

		var detailsToggle = new LinkLabel { Text = "Show technical details", Location = new Point(30, 458), Size = new Size(240, 28) };
		detailsToggle.LinkClicked += (_, _) => { detailsBox.Visible = !detailsBox.Visible; detailsToggle.Text = detailsBox.Visible ? "Hide technical details" : "Show technical details"; ClientSize = new Size(760, detailsBox.Visible ? 770 : 570); };
		Controls.Add(detailsToggle);
		detailsBox.Location = new Point(30, 580);
		detailsBox.Size = new Size(700, 155);
		detailsBox.BackColor = Color.FromArgb(247, 249, 252);
		detailsBox.Multiline = true;
		detailsBox.ReadOnly = true;
		detailsBox.Visible = false;
		detailsBox.ScrollBars = ScrollBars.Vertical;
		detailsBox.Text = options.Uninstall
			? "Click Remove to uninstall Maily from this Windows user."
			: "Review the acknowledgement above, then continue. Setup shows each check here. If a check fails, follow the displayed recovery instructions and retry.";
		Controls.Add(detailsBox);

		primaryButton.Text = options.Uninstall ? "Remove" : installedVersion is null ? "Install" : "Update";
		primaryButton.Enabled = options.Uninstall;
		primaryButton.Location = new Point(454, 510);
		primaryButton.Size = new Size(164, 34);
		primaryButton.BackColor = Color.FromArgb(32, 86, 155);
		primaryButton.ForeColor = Color.White;
		primaryButton.FlatStyle = FlatStyle.Flat;
		primaryButton.Click += async (_, _) => await ContinueAsync();
		Controls.Add(primaryButton);

		closeButton.Text = "Close";
		closeButton.Location = new Point(638, 510);
		closeButton.Size = new Size(92, 34);
		closeButton.Click += (_, _) => Close();
		Controls.Add(closeButton);
	}

	private async Task ContinueAsync()
	{
		if (busy) return;
		if (options.Uninstall)
		{
			if (flow.LocalRemovalComplete) { ShowAccountRemovalGuidance(); return; }
			await RunAsync(); return;
		}
		switch (flow.NextAction)
		{
			case SetupAction.Connect:
				if (ShowConnectionGuide() != DialogResult.OK) return;
				if (flow.AfterConnectionGuide == SetupAction.CheckConnection) await VerifyOutlookAsync(); else await RunAsync();
				break;
			case SetupAction.CheckConnection:
				OutlookProcessHelper.StartOutlook("");
				await VerifyOutlookAsync(); break;
			case SetupAction.OpenOutlook:
				OutlookProcessHelper.StartOutlook(""); Close(); break;
			case SetupAction.Close: Close(); break;
			default: await RunAsync(); break;
		}
	}

	private DialogResult ShowConnectionGuide()
	{
		using var guide = CreateConnectionGuide();
		return guide.ShowDialog(this);
	}

	internal Form CreateConnectionGuide(Func<CancellationToken, Task<OutlookConnectionResult>>? connector = null)
	{
		var allowed = flow.CanConnectAutomatically;
		var guide = new Form { Text = "Connect Maily to Outlook", StartPosition = FormStartPosition.CenterParent, ClientSize = new Size(650, 360), Font = Font, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false };
		guide.Controls.Add(new Label { Text = flow.RepairRequired ? "Update your Outlook connection" : "Connect Maily to your Outlook account", Font = new Font(Font, FontStyle.Bold), Location = new Point(24, 20), Size = new Size(600, 40) });
		guide.Controls.Add(new Label { Text = allowed ? "Sign in with the Microsoft account you use in Outlook. Microsoft will ask you to approve access. If your account is unsupported, use the manual steps below." : "This connection requires administrator or Maily support. Setup will not change an organization-managed or unknown connection.", Location = new Point(24, 68), Size = new Size(600, 65) });
		var signIn = new Button { Name = "connectionSignIn", Text = "Sign in and connect", Location = new Point(24, 140), Size = new Size(210, 35), Enabled = allowed };
		var cancel = new Button { Name = "connectionCancel", Text = "Cancel connection", Location = new Point(250, 140), Size = new Size(180, 35), Visible = false };
		var connectionStatus = new Label { Name = "connectionStatus", Text = "Your Outlook connection is checked separately after Microsoft finishes.", Location = new Point(24, 190), Size = new Size(600, 55), AutoEllipsis = true };
		var fallback = new LinkLabel { Text = "Other ways to connect", Location = new Point(24, 250), Size = new Size(280, 25), Enabled = allowed };
		var manual = new Panel { Location = new Point(24, 280), Size = new Size(600, 220), Visible = false };
		manual.Controls.Add(new Label { Text = flow.ConnectionInstructions, Location = Point.Empty, Size = new Size(600, 160) });
		var open = new Button { Text = "Open Outlook add-ins", Location = new Point(0, 175), Size = new Size(210, 35) };
		open.Click += (_, _) => OutlookAddinsPage.Open();
		var copy = new Button { Text = "Copy setup file", Location = new Point(225, 175), Size = new Size(170, 35) };
		copy.Click += (_, _) => { var path = Path.Combine(AppContext.BaseDirectory, "payload", "addin", "manifest.xml"); if (!File.Exists(path)) { MessageBox.Show(guide, "The setup file is missing. Download a complete Maily package."); return; } try { Clipboard.SetText(path); copy.Text = "Copied"; } catch (System.Runtime.InteropServices.ExternalException) { MessageBox.Show(guide, "Could not copy the setup file location. Try again."); } };
		manual.Controls.AddRange(new Control[] { open, copy });
		var check = new Button { Text = flow.ApplicationInstalled ? "Check connection" : "Continue setup", DialogResult = DialogResult.OK, Location = new Point(415, 305), Size = new Size(205, 35), Enabled = allowed };
		fallback.LinkClicked += (_, _) => { manual.Visible = !manual.Visible; guide.ClientSize = new Size(650, manual.Visible ? 570 : 360); check.Top = manual.Visible ? 520 : 305; fallback.Text = manual.Visible ? "Hide manual steps" : "Other ways to connect"; };
		CancellationTokenSource? connectionCancellation = null;
		var connecting = false;
		var closeAfterCancel = false;
		cancel.Click += (_, _) => connectionCancellation?.Cancel();
		guide.FormClosing += (_, e) => { if (connecting) { e.Cancel = true; closeAfterCancel = true; connectionCancellation?.Cancel(); } };
		signIn.Click += async (_, _) =>
		{
			if (connecting || !flow.CanConnectAutomatically) return;
			connecting = true; signIn.Enabled = false; check.Enabled = false; fallback.Enabled = false; manual.Enabled = false; cancel.Visible = true;
			connectionStatus.Text = "Complete Microsoft's sign-in and approval in your browser. You can cancel here.";
			using var cancellation = CancellationTokenSource.CreateLinkedTokenSource(setupCancellation.Token);
			connectionCancellation = cancellation;
			try
			{
				var result = await (connector ?? ConnectOutlookAccountAsync)(cancellation.Token);
				connectionStatus.Text = result.Message;
				// Server registration is not proof Outlook loaded it. Never set flow verified here.
				AppendDetails("Microsoft connection attempt: " + result.Code + ". Outlook verification remains separate.");
			}
			catch { connectionStatus.Text = OutlookConnectionResult.For("connection-unavailable").Message; }
			finally
			{
				connecting = false; connectionCancellation = null; signIn.Enabled = allowed; check.Enabled = allowed; fallback.Enabled = allowed; manual.Enabled = true; cancel.Visible = false;
				if (closeAfterCancel) guide.Close();
			}
		};
		guide.Controls.AddRange(new Control[] { signIn, cancel, connectionStatus, fallback, manual, check });
		return guide;
	}

	private async Task<OutlookConnectionResult> ConnectOutlookAccountAsync(CancellationToken cancellationToken)
	{
		if (!flow.CanConnectAutomatically) return OutlookConnectionResult.For("connection-unavailable");
		try
		{
			await Task.Run(() => PackageIntegrity.Verify(AppContext.BaseDirectory, IsHostedPackage), cancellationToken);
			cancellationToken.ThrowIfCancellationRequested();
			var sessions = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Maily", "OutlookAddin", "data", "connection-sessions");
			var start = OutlookConnectionRunner.CreateStartInfo(AppContext.BaseDirectory, sessions);
			return await OutlookConnectionRunner.RunAsync(start, cancellationToken);
		}
		catch (OperationCanceledException) { return OutlookConnectionResult.For("cancelled"); }
		catch { return OutlookConnectionResult.For("invalid-package"); }
	}
	private void ShowAccountRemovalGuidance()
	{
		if (MessageBox.Show(this, "Maily is removed from this PC. Its Outlook account entry may still remain. Open My add-ins and remove your own Maily entry if it is listed. For an organization-managed entry, contact your administrator.\r\n\r\nOpen Outlook add-ins now?", "Finish removing Maily", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes) OutlookAddinsPage.Open();
		Close();
	}

	private async Task RunAsync()
	{
		if (busy) return;
		busy = true;
		try { await RunCoreAsync(); }
		catch (Exception ex)
		{
			if (IsDisposed) return;
			statusLabel.Text = "Setup could not continue. Review the details and retry.";
			AppendDetails($"{ex.GetType().Name}: {ex.Message}");
			AppendDetails("Keep this window's details for support. Your existing sign-in is not removed by setup.");
			progressBar.Style = ProgressBarStyle.Continuous;
			RestoreReadyControls();
		}
		finally { busy = false; applyingChanges = false; }
	}

	private async Task RunCoreAsync()
	{
		if (!options.Uninstall && !responsibilityCheckBox.Checked) return;
		foreach (ListViewItem item in checks.Items) { item.SubItems[1].Text = "Waiting"; item.ForeColor = Color.FromArgb(35, 48, 64); }
		if (!options.Uninstall)
		{
			detailsBox.Clear();
			SetCheck("package", "Checking");
			try
			{
				primaryButton.Enabled = false;
				responsibilityCheckBox.Enabled = false;
				await Task.Run(() => PackageIntegrity.Verify(AppContext.BaseDirectory, IsHostedPackage));
			}
			catch { SetCheck("package", "Needs attention"); throw; }
			if (IsDisposed) return;
			SetCheck("package", "Verified");
			SetCheck("account", "Checking");
			if (IsHostedPackage) { SetCheck("account", "Not required"); RestoreReadyControls(); }
		}
		if (!options.Uninstall && !IsHostedPackage)
		{
			primaryButton.Enabled = false;
			responsibilityCheckBox.Enabled = false;
			statusLabel.Text = "Checking existing ChatGPT sign-in...";
			var codexStatus = await Task.Run(CodexSetupPreflight.Check);
			if (IsDisposed) return;
			if (!codexStatus.Ok && File.Exists(codexStatus.CodexExecutablePath))
			{
				detailsBox.Text = CodexSetupPreflight.DescribeInstalledApps() + Environment.NewLine +
					"Existing authentication was not available. Sign in securely in your browser; Maily does not read or copy credentials.";
				if (MessageBox.Show(this, "Maily needs ChatGPT sign-in. Open secure browser sign-in now?", "Sign in to ChatGPT", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
				{
					statusLabel.Text = "Complete ChatGPT sign-in in your browser...";
					await CodexSetupPreflight.SignInAsync(codexStatus.CodexExecutablePath, setupCancellation.Token);
					if (IsDisposed) return;
					codexStatus = await Task.Run(CodexSetupPreflight.Check);
					if (IsDisposed) return;
				}
			}
			primaryButton.Enabled = responsibilityCheckBox.Checked;
			responsibilityCheckBox.Enabled = true;
			if (!codexStatus.Ok)
			{
				SetCheck("account", "Needs attention");
				statusLabel.Text = "ChatGPT sign-in is not ready. You can retry Install.";
				detailsBox.Text = codexStatus.ToDisplayText();
				MessageBox.Show(
					this,
					codexStatus.ToDisplayText(),
					"Maily sign-in",
					MessageBoxButtons.OK,
					MessageBoxIcon.Warning);
				return;
			}
			SetCheck("account", "Sign-in ready");

			if (CertificateTrustPrompt.Show(this) != DialogResult.OK)
			{
				return;
			}
		}

		var shouldCloseOutlook = false;
		var restartOutlookPath = "";
		if (OutlookProcessHelper.IsOutlookRunning())
		{
			if (OutlookClosePrompt.Show(this, options.Uninstall) != DialogResult.OK)
			{
				return;
			}

			shouldCloseOutlook = true;
			restartOutlookPath = options.Uninstall ? "" : OutlookProcessHelper.FindRunningOutlookPath();
		}

		primaryButton.Enabled = false;
		closeButton.Enabled = false;
		keepDataCheckBox.Enabled = false;
		removeTrustedCertificateCheckBox.Enabled = false;
		responsibilityCheckBox.Enabled = false;
		progressBar.Style = ProgressBarStyle.Marquee;
		detailsBox.Text = "";

		if (shouldCloseOutlook)
		{
			statusLabel.Text = "Closing Outlook...";
			AppendDetails("Closing Outlook so Maily registration and cache can be updated cleanly.");
			var closeResult = await Task.Run(() => OutlookProcessHelper.CloseOutlookGracefully(options.Uninstall && !string.IsNullOrWhiteSpace(options.InstallRoot) ? options.InstallRoot : AppContext.BaseDirectory));
			if (!closeResult.Success)
			{
				progressBar.Style = ProgressBarStyle.Continuous;
				progressBar.Value = 0;
				statusLabel.Text = "Outlook is still open.";
				AppendDetails("");
				AppendDetails("Outlook did not close. Finish/save and close inline replies, or finish any Outlook prompts, then retry here.");
				RestoreReadyControls();
				return;
			}

			AppendDetails("Outlook closed successfully.");
			AppendDetails("");
		}

		statusLabel.Text = options.Uninstall ? "Removing Maily..." : "Installing Maily...";
		applyingChanges = true;

		var effectiveOptions = options with
		{
			PackageMode = packageMode,
			KeepUserData = keepDataCheckBox.Checked,
			KeepLogs = true,
			RemoveTrustedCertificate = !IsHostedPackage && removeTrustedCertificateCheckBox.Checked,
			CloseOutlook = shouldCloseOutlook
		};

		var result = await Task.Run(() => ScriptRunner.Run(effectiveOptions, AppendDetails));
		applyingChanges = false;
		progressBar.Style = ProgressBarStyle.Continuous;
		progressBar.Value = result.ExitCode == 0 ? 100 : 0;
		closeButton.Enabled = true;


        if (options.Uninstall)
        {
            ApplyRemovalResult(result.ExitCode);
            RestoreReadyControls();
            return;
        }
        ApplyInstallationResult(result.ExitCode, result.ApplicationInstalled || result.ExitCode == 21, result.ActivationState);
        SaveConnectionProgress();
        if (shouldCloseOutlook) OutlookProcessHelper.StartOutlook(restartOutlookPath);
        if (result.ExitCode is 20 or 21) { SetCheck("connect", "Needs your action"); progressBar.Value = flow.ApplicationInstalled ? 66 : 33; }
        RestoreReadyControls();
    }

	internal void ApplyInstallationResult(int exitCode, bool applicationInstalled, string activationState = "unknown")
	{
		flow.InstallationFinished(exitCode, applicationInstalled, activationState);
		statusLabel.Text = flow.Message;
		primaryButton.Text = flow.PrimaryText;
		if (applicationInstalled) SetCheck("install", "Complete");
		if (exitCode is 20 or 21 or 23) SetCheck("connect", "Needs your action");
	}

	private void SaveConnectionProgress()
	{
		try
		{
			var root = options.InstallRoot ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Maily", "OutlookAddin");
			var progress = flow.SaveConnectionProgress(root, PackageFingerprint());
			if (progress is null) return;
			var target = Path.Combine(root, "data", "setup-connection-progress.json");
			var temporary = target + "." + Guid.NewGuid().ToString("N");
			File.WriteAllText(temporary, progress, new UTF8Encoding(false));
			File.Move(temporary, target, true);
		}
		catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException) { AppendDetails("Could not save connection progress. Keep this window open while connecting Outlook."); }
	}

	internal void ApplyRemovalResult(int exitCode)
	{
		flow.RemovalFinished(exitCode);
		statusLabel.Text = flow.Message;
		primaryButton.Text = flow.LocalRemovalComplete ? "Finish removal" : "Try again";
	}

	private static string PackageFingerprint() => Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "SHA256SUMS.json"))));

	private async Task VerifyOutlookAsync()
	{
		if (busy || options.Uninstall) return;
		busy = true;
		applyingChanges = true;
		primaryButton.Enabled = false;
		closeButton.Enabled = false;
		statusLabel.Text = "Checking Outlook registration and installed service...";
		try
		{
			var result = await Task.Run(() => ScriptRunner.Run(options with { VerifyOutlookActivation = true, Quiet = true, CloseOutlook = false }, AppendDetails));
			statusLabel.Text = result.ExitCode == 0 ? "Outlook registration verified. Open a new draft to check Maily." : "Outlook activation is not verified. Follow the steps below.";
			flow.ConnectionFinished(result.ExitCode, result.ActivationState);
            statusLabel.Text = flow.Message;
            primaryButton.Text = flow.PrimaryText;
            if (result.ExitCode != 0) SetCheck("connect", "Needs your action");
			else
			{
				checks.Items["connect"]!.SubItems[1].Text = "Complete";
				checks.Items["connect"]!.ForeColor = Color.DarkGreen;
				progressBar.Value = 100;
			}
		}
		catch (Exception ex) { statusLabel.Text = "Outlook verification failed."; AppendDetails(ex.Message); }
		finally { busy = false; applyingChanges = false; RestoreReadyControls(); }
	}

	private void RestoreReadyControls()
	{
		primaryButton.Enabled = true;
		responsibilityCheckBox.Enabled = !options.Uninstall;
		closeButton.Enabled = true;
		keepDataCheckBox.Enabled = false;
		removeTrustedCertificateCheckBox.Enabled = options.Uninstall && !IsHostedPackage;
	}

	private void AppendDetails(string line)
	{
		if (IsDisposed || Disposing) return;
		if (InvokeRequired)
		{
			try { BeginInvoke(new Action<string>(AppendDetails), line); } catch (InvalidOperationException) { }
			return;
		}
		if (line.StartsWith("MAILY_STEP|", StringComparison.Ordinal))
		{
			var parts = line.Split('|', 4);
			if (parts.Length >= 3)
			{
				var stage = parts[1];
				var state = parts[2];
				var group = stage.Contains("rollback") || stage.Contains("backup") || stage.Contains("replace") || stage.Contains("upgrade") || stage.Contains("copy-") || stage.Contains("stop-existing") ? "replace"
					: stage.Contains("package") || stage.Contains("integrity") ? "package"
					: stage.Contains("codex") ? "account"
					: stage.Contains("health") || stage.Contains("ready") || stage.Contains("verification") ? "verify" : "configure";
				SetCheck(group, state switch { "running" => "In progress", "passed" => "Complete", "failed" => "Needs attention", "warning" => "Review details", "skipped" => "Not required", _ => state });
                if (state == "running") statusLabel.Text = checks.Items[StageGroup(group)]?.Text + "...";
				line = $"[{state}] {stage}: {(parts.Length == 4 ? parts[3] : string.Empty)}";
			}
		}

		detailsBox.AppendText(line + Environment.NewLine);
	}

	private static string StageGroup(string key) => key switch { "package" or "account" => "prepare", "replace" => "install", "configure" or "verify" => "connect", _ => key };

    private void SetCheck(string key, string state)
    {
        key = StageGroup(key);
		var item = checks.Items[key];
		if (item is null) return;
		var previous = item.SubItems[1].Text;
		// Several script steps share one row. Never hide a failure or warning behind a later success.
		if (previous == "Needs attention" || (previous == "Review details" && state != "Needs attention")) return;
		if (state == "Not required" && previous is not "Waiting" and not "Checking") return;
		item.SubItems[1].Text = state;
		item.ForeColor = state == "Needs attention" ? Color.Firebrick : state == "Review details" ? Color.DarkGoldenrod : state == "Complete" ? Color.DarkGreen : Color.FromArgb(35, 48, 64);
	}

	private static string? DetectInstalledVersion(SetupOptions options)
	{
		try
		{
			var root = options.InstallRoot ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Maily", "OutlookAddin");
			var manifest = Path.Combine(root, "addin", "manifest.xml");
			if (File.Exists(manifest))
			{
				try
				{
					using var reader = System.Xml.XmlReader.Create(manifest, new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Prohibit, XmlResolver = null });
					var document = System.Xml.Linq.XDocument.Load(reader);
					var version = document.Root?.Elements().FirstOrDefault(element => element.Name.LocalName == "Version")?.Value.Trim();
					if (!string.IsNullOrWhiteSpace(version)) return version;
				}
				catch (System.Xml.XmlException) { }
			}
			var host = Path.Combine(root, "app", "Maily.OutlookAddin.Web", "Maily.OutlookAddin.Web.exe");
			if (File.Exists(host)) return FileVersionInfo.GetVersionInfo(host).FileVersion ?? "existing installation";
			if (File.Exists(manifest)) return "existing installation";
		}
		catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.ComponentModel.Win32Exception or ArgumentException) { }
		return null;
	}


}

internal static class InstallSummary
{
	public static bool HasMailboxConflictWarning()
	{
		try
		{
			var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
			var path = Path.Combine(localAppData, "Maily", "OutlookAddin", "data", "install-summary.json");
			if (!File.Exists(path))
			{
				return false;
			}

			using var document = JsonDocument.Parse(File.ReadAllText(path));
			if (!document.RootElement.TryGetProperty("steps", out var steps) || steps.ValueKind != JsonValueKind.Array)
			{
				return false;
			}

			foreach (var step in steps.EnumerateArray())
			{
				var name = step.TryGetProperty("name", out var nameElement) ? nameElement.GetString() : "";
				var status = step.TryGetProperty("status", out var statusElement) ? statusElement.GetString() : "";
				if (string.Equals(name, "mailbox-addin-conflict", StringComparison.OrdinalIgnoreCase) &&
					string.Equals(status, "warning", StringComparison.OrdinalIgnoreCase))
				{
					return true;
				}
			}
		}
		catch
		{
		}

		return false;
	}
}

internal static class OutlookAddinsPage
{
	public static void Open()
	{
		try
		{
			Process.Start(new ProcessStartInfo
			{
				FileName = "https://aka.ms/olksideload",
				UseShellExecute = true
			});
		}
		catch
		{
			MessageBox.Show(
				"Open Outlook add-ins and remove the older Maily entry, then run Maily setup again.",
				"Open Outlook add-ins",
				MessageBoxButtons.OK,
				MessageBoxIcon.Information);
		}
	}
}

internal static class CertificateTrustPrompt
{
	public static DialogResult Show(IWin32Window owner)
	{
		using var form = new Form
		{
			Text = "Trust Local Maily Connection",
			StartPosition = FormStartPosition.CenterParent,
			FormBorderStyle = FormBorderStyle.FixedDialog,
			MaximizeBox = false,
			MinimizeBox = false,
			ClientSize = new Size(560, 250),
			Font = new Font("Segoe UI", 9F)
		};

		var title = new Label
		{
			Text = "Maily needs a trusted local HTTPS connection",
			Font = new Font("Segoe UI Semibold", 14F, FontStyle.Bold),
			AutoSize = false,
			Location = new Point(22, 20),
			Size = new Size(510, 28)
		};
		form.Controls.Add(title);

		var body = new Label
		{
			Text = "Outlook requires the add-in pane to load over HTTPS. Setup will create a localhost-only certificate for this Windows user. If Windows asks whether to trust this certificate, choose Yes to let Outlook open Maily without browser certificate errors.",
			AutoSize = false,
			Location = new Point(24, 62),
			Size = new Size(510, 96)
		};
		form.Controls.Add(body);

		var continueButton = new Button
		{
			Text = "Continue setup",
			DialogResult = DialogResult.OK,
			Location = new Point(315, 195),
			Size = new Size(112, 32)
		};
		form.Controls.Add(continueButton);

		var cancelButton = new Button
		{
			Text = "Cancel",
			DialogResult = DialogResult.Cancel,
			Location = new Point(438, 195),
			Size = new Size(92, 32)
		};
		form.Controls.Add(cancelButton);

		form.AcceptButton = continueButton;
		form.CancelButton = cancelButton;
		return form.ShowDialog(owner);
	}
}

internal sealed record CodexSetupStatus(
	bool Ok,
	string CodexHome,
	string CodexExecutablePath,
	string LoginStatusOutput,
	string Message,
	string RecommendedAction,
	IReadOnlyList<string> MissingItems)
{
	public string ToDisplayText()
	{
		if (Ok)
		{
			return Message;
		}

		var builder = new StringBuilder();
		builder.AppendLine(Message);
		builder.AppendLine();
		builder.AppendLine(RecommendedAction);
		builder.AppendLine();
		builder.AppendLine("Checked Codex folder:");
		builder.AppendLine(CodexHome);
		builder.AppendLine();
		builder.AppendLine("Checked Codex CLI:");
		builder.AppendLine(CodexExecutablePath);
		if (!string.IsNullOrWhiteSpace(LoginStatusOutput))
		{
			builder.AppendLine();
			builder.AppendLine("Login status:");
			builder.AppendLine(LoginStatusOutput);
		}
		if (MissingItems.Count > 0)
		{
			builder.AppendLine();
			builder.AppendLine("Missing or incomplete:");
			foreach (var item in MissingItems)
			{
				builder.AppendLine("- " + item);
			}
		}

		return builder.ToString().TrimEnd();
	}
}

internal static class CodexSetupPreflight
{
	public static CodexSetupStatus Check()
	{
		var codexHome = ResolveCodexHome();
		var attempts = new List<string>();
		var selectedPath = string.Empty;
		var selectedOutput = string.Empty;
		var accessDenied = false;

		foreach (var candidate in GetCodexCandidates())
		{
			var loginStatus = RunCodex(candidate, codexHome, "login", "status");
			var output = Truncate(CombineOutput(loginStatus.Stdout, loginStatus.Stderr, loginStatus.Error), 1200);
			var loggedIn = loginStatus.Started &&
				loginStatus.ExitCode == 0 &&
				output.Contains("Logged in", StringComparison.OrdinalIgnoreCase) &&
				output.Contains("ChatGPT", StringComparison.OrdinalIgnoreCase);
			if (loginStatus.Started && loginStatus.ExitCode != -1 && string.IsNullOrWhiteSpace(selectedPath)) selectedPath = candidate;
			accessDenied |= output.Contains("Access is denied", StringComparison.OrdinalIgnoreCase);
			attempts.Add($"{candidate}: {(loggedIn ? "ready" : output.Split(Environment.NewLine).FirstOrDefault() ?? "not ready")}");
			if (loggedIn)
			{
				selectedPath = candidate;
				selectedOutput = "ChatGPT sign-in verified.";
				break;
			}
		}

		var missing = new List<string>();
		if (string.IsNullOrWhiteSpace(selectedOutput) && attempts.Count == 0)
		{
			missing.Add("Runnable Codex CLI");
		}
		else if (string.IsNullOrWhiteSpace(selectedOutput))
		{
			missing.Add(accessDenied
				? "Runnable Codex CLI for background add-in use"
				: "ChatGPT sign-in via codex login");
		}

		if (missing.Count == 0)
		{
			return new CodexSetupStatus(
				true,
				codexHome,
				selectedPath,
				selectedOutput,
				"Codex CLI is ready and signed in with ChatGPT.",
				string.Empty,
				missing);
		}

		return new CodexSetupStatus(
			false,
			codexHome,
			selectedPath,
			string.Empty,
			accessDenied
				? "Codex desktop is installed, but Windows denied direct execution of its packaged CLI from the Maily add-in backend."
				: "Maily needs a supported ChatGPT sign-in before generating replies, summaries, search answers, or briefings.",
			accessDenied
				? "The Maily runtime could not start. Obtain a complete Maily package and check your organization's application policy."
				: "Click Install to retry browser sign-in. If the runtime is missing, obtain a complete Maily package. No separate Codex installation is required.",
			missing);
	}

	private static IReadOnlyList<string> GetCodexCandidates()
	{
		var bundled = Path.Combine(AppContext.BaseDirectory, "payload", "runtime", "codex", "codex.exe");
		if (File.Exists(bundled)) return [bundled];
		var candidates = new List<string> { "codex", "codex.exe", "codex.cmd" };
		foreach (var directory in (Environment.GetEnvironmentVariable("PATH") ?? string.Empty)
			.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
		{
			foreach (var fileName in new[] { "codex.exe", "codex.cmd", "codex.bat", "codex" })
			{
				try
				{
					candidates.Add(Path.Combine(directory, fileName));
				}
				catch
				{
				}
			}
		}

		var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
		var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
		var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
		candidates.AddRange(new[]
		{
			Path.Combine(appData, "npm", "codex.cmd"),
			Path.Combine(appData, "npm", "codex.exe"),
			Path.Combine(localAppData, "Microsoft", "WindowsApps", "codex.exe"),
			Path.Combine(localAppData, "Microsoft", "WindowsApps", "codex"),
			Path.Combine(localAppData, "OpenAI", "Codex", "bin", "codex.exe"),
			Path.Combine(localAppData, "OpenAI", "Codex", "bin", "codex"),
			Path.Combine(localAppData, "Programs", "Codex", "codex.exe"),
			Path.Combine(localAppData, "Programs", "OpenAI Codex", "codex.exe"),
			Path.Combine(programFiles, "OpenAI", "Codex", "codex.exe")
		});

		var openAiCodexBin = Path.Combine(localAppData, "OpenAI", "Codex", "bin");
		if (Directory.Exists(openAiCodexBin))
		{
			try
			{
				candidates.AddRange(Directory
					.EnumerateFiles(openAiCodexBin, "codex*", SearchOption.AllDirectories)
					.Where(path =>
					{
						var fileName = Path.GetFileName(path);
						return string.Equals(fileName, "codex.exe", StringComparison.OrdinalIgnoreCase) ||
							string.Equals(fileName, "codex.cmd", StringComparison.OrdinalIgnoreCase) ||
							string.Equals(fileName, "codex.bat", StringComparison.OrdinalIgnoreCase) ||
							string.Equals(fileName, "codex", StringComparison.OrdinalIgnoreCase);
					})
					.OrderByDescending(path => File.GetLastWriteTimeUtc(path)));
			}
			catch
			{
			}
		}

		return candidates
			.Where(candidate => !string.IsNullOrWhiteSpace(candidate) && (!Path.IsPathRooted(candidate) || File.Exists(candidate)))
			.Distinct(StringComparer.OrdinalIgnoreCase)
			.OrderByDescending(Path.IsPathRooted)
			.ToList();
	}

	public static string DescribeInstalledApps()
	{
		// Installation is only a hint. Only the supported runtime login-status check establishes authentication.
		var packages = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Packages");
		try
		{
			if (Directory.Exists(packages) && Directory.EnumerateDirectories(packages)
				.Any(path => Path.GetFileName(path).Contains("ChatGPT", StringComparison.OrdinalIgnoreCase) || Path.GetFileName(path).Contains("Codex", StringComparison.OrdinalIgnoreCase)))
				return "ChatGPT or Codex app installation detected; its presence alone does not confirm sign-in.";
		}
		catch (IOException) { }
		catch (UnauthorizedAccessException) { }
		return "Maily includes the runtime needed for ChatGPT sign-in.";
	}

	public static async Task SignInAsync(string executablePath, CancellationToken cancellationToken)
	{
		var startInfo = CreateCodexStartInfo(executablePath, ["login"]);
		var codexHome = ResolveCodexHome();
		if (!string.IsNullOrWhiteSpace(codexHome)) startInfo.Environment["CODEX_HOME"] = codexHome;
		using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
		timeout.CancelAfter(TimeSpan.FromMinutes(5));
		Process? process = null;
		try
		{
			process = Process.Start(startInfo);
			if (process is null) return;
			// Drain output without retaining or displaying OAuth URLs or login diagnostics.
			var stdout = process.StandardOutput.BaseStream.CopyToAsync(Stream.Null, timeout.Token);
			var stderr = process.StandardError.BaseStream.CopyToAsync(Stream.Null, timeout.Token);
			await process.WaitForExitAsync(timeout.Token);
			await Task.WhenAll(stdout, stderr);
		}
		catch (Exception ex) when (ex is OperationCanceledException or System.ComponentModel.Win32Exception or IOException) { }
		finally
		{
			if (process is not null)
			{
				try { if (!process.HasExited) process.Kill(entireProcessTree: true); } catch (InvalidOperationException) { } catch (System.ComponentModel.Win32Exception) { }
				process.Dispose();
			}
		}
	}

	private static string ResolveCodexHome()
	{
		var envCodexHome = Environment.GetEnvironmentVariable("CODEX_HOME");
		if (!string.IsNullOrWhiteSpace(envCodexHome))
		{
			return Path.GetFullPath(envCodexHome);
		}

		return string.Empty;
	}

	private static CodexProcessResult RunCodex(string executablePath, string codexHome, params string[] arguments)
	{
		var startInfo = CreateCodexStartInfo(executablePath, arguments);
		if (!string.IsNullOrWhiteSpace(codexHome))
		{
			startInfo.Environment["CODEX_HOME"] = codexHome;
		}

		startInfo.Environment["NO_COLOR"] = "1";

		try
		{
			using var process = Process.Start(startInfo);
			if (process is null)
			{
				return new CodexProcessResult(false, -1, string.Empty, string.Empty, "Codex CLI did not start.");
			}

			var stdoutTask = process.StandardOutput.ReadToEndAsync();
			var stderrTask = process.StandardError.ReadToEndAsync();
			if (!process.WaitForExit((int)TimeSpan.FromSeconds(10).TotalMilliseconds))
			{
				try
				{
					process.Kill(entireProcessTree: true);
				}
				catch
				{
				}

				process.WaitForExit(1000);
				return new CodexProcessResult(
					true,
					-1,
					stdoutTask.IsCompletedSuccessfully ? stdoutTask.Result : string.Empty,
					stderrTask.IsCompletedSuccessfully ? stderrTask.Result : string.Empty,
					"Codex CLI timed out.");
			}

			return new CodexProcessResult(true, process.ExitCode, stdoutTask.Result, stderrTask.Result, string.Empty);
		}
		catch (Exception ex)
		{
			return new CodexProcessResult(false, -1, string.Empty, string.Empty, ex.Message);
		}
	}

	private static string CombineOutput(params string[] values) =>
		string.Join(Environment.NewLine, values.Where(value => !string.IsNullOrWhiteSpace(value))).Trim();

	private static ProcessStartInfo CreateCodexStartInfo(string executablePath, IReadOnlyList<string> arguments)
	{
		var startInfo = new ProcessStartInfo
		{
			FileName = executablePath,
			UseShellExecute = false,
			RedirectStandardOutput = true,
			RedirectStandardError = true,
			CreateNoWindow = true
		};
		if (executablePath.EndsWith(".cmd", StringComparison.OrdinalIgnoreCase) ||
			executablePath.EndsWith(".bat", StringComparison.OrdinalIgnoreCase))
		{
			startInfo.FileName = Environment.GetEnvironmentVariable("COMSPEC") ?? "cmd.exe";
			startInfo.ArgumentList.Add("/d");
			startInfo.ArgumentList.Add("/c");
			startInfo.ArgumentList.Add(string.Join(" ", new[] { executablePath }.Concat(arguments).Select(QuoteCommandArgument)));
			return startInfo;
		}

		foreach (var argument in arguments)
		{
			startInfo.ArgumentList.Add(argument);
		}

		return startInfo;
	}

	private static string QuoteCommandArgument(string value) =>
		"\"" + value.Replace("\"", "\\\"") + "\"";

	private static string Truncate(string value, int maxLength) =>
		string.IsNullOrWhiteSpace(value) || value.Length <= maxLength ? value : value[..maxLength];

	private sealed record CodexProcessResult(
		bool Started,
		int ExitCode,
		string Stdout,
		string Stderr,
		string Error);
}

internal static class OutlookClosePrompt
{
	public static DialogResult Show(IWin32Window owner, bool uninstall)
	{
		using var form = new Form
		{
			Text = "Close Outlook",
			StartPosition = FormStartPosition.CenterParent,
			FormBorderStyle = FormBorderStyle.FixedDialog,
			MaximizeBox = false,
			MinimizeBox = false,
			ClientSize = new Size(520, 220),
			Font = new Font("Segoe UI", 9F)
		};

		var title = new Label
		{
			Text = "Outlook needs to close first",
			Font = new Font("Segoe UI Semibold", 14F, FontStyle.Bold),
			AutoSize = false,
			Location = new Point(22, 20),
			Size = new Size(470, 28)
		};
		form.Controls.Add(title);

		var body = new Label
		{
			Text = uninstall
				? "Outlook is still running and can keep showing cached Maily ribbon buttons after removal. Save any drafts, then let setup close Outlook and finish the cleanup."
				: "Outlook is still running and can keep old add-in metadata cached. Save any drafts, then let setup close Outlook, install Maily, and reopen Outlook after setup succeeds.",
			AutoSize = false,
			Location = new Point(24, 62),
			Size = new Size(470, 74)
		};
		form.Controls.Add(body);

		var continueButton = new Button
		{
			Text = "Close Outlook and continue",
			DialogResult = DialogResult.OK,
			Location = new Point(235, 165),
			Size = new Size(155, 32)
		};
		form.Controls.Add(continueButton);

		var cancelButton = new Button
		{
			Text = "Cancel setup",
			DialogResult = DialogResult.Cancel,
			Location = new Point(400, 165),
			Size = new Size(92, 32)
		};
		form.Controls.Add(cancelButton);

		form.AcceptButton = continueButton;
		form.CancelButton = cancelButton;
		return form.ShowDialog(owner);
	}
}

internal static class OutlookProcessHelper
{
	public static bool IsOutlookRunning()
	{
		var processes = Process.GetProcessesByName("OUTLOOK");
		foreach (var process in processes)
		{
			process.Dispose();
		}

		return processes.Length > 0;
	}

	public static string FindRunningOutlookPath()
	{
		foreach (var process in Process.GetProcessesByName("OUTLOOK"))
		{
			using (process)
			{
				try
				{
					if (!string.IsNullOrWhiteSpace(process.MainModule?.FileName))
					{
						return process.MainModule.FileName;
					}
				}
				catch
				{
				}
			}
		}

		return "";
	}

	public static OutlookCloseResult CloseOutlookGracefully(string helperRoot)
	{
		return OutlookShutdown.Run(Path.Combine(helperRoot, "setup", "OutlookShutdown.ps1"));
	}
	public static void StartOutlook(string path)
	{
		try
		{
			if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
			{
				Process.Start(new ProcessStartInfo { FileName = path, UseShellExecute = true });
			}
			else
			{
				Process.Start(new ProcessStartInfo { FileName = "OUTLOOK.EXE", UseShellExecute = true });
			}
		}
		catch
		{
			// Reopening Outlook is a convenience, not a setup failure.
		}
	}
}

internal sealed record OutlookCloseResult(bool Success);

internal static class PackageIntegrity
{
	public static void Verify(string directory, bool hosted)
	{
		var root = Path.GetFullPath(directory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
		var inventory = Path.Combine(root, "SHA256SUMS.json");
		if (!File.Exists(inventory)) throw new InvalidOperationException("Package integrity information is missing. Extract a complete, current Maily installer ZIP and retry.");
		using var document = JsonDocument.Parse(File.ReadAllText(inventory));
		if (document.RootElement.ValueKind != JsonValueKind.Array) throw new InvalidOperationException("Invalid package integrity inventory.");
		var covered = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		foreach (var item in document.RootElement.EnumerateArray())
		{
			var relative = item.GetProperty("path").GetString() ?? string.Empty;
			var parts = relative.Replace('\\', '/').Split('/');
			if (parts.Any(part => part is "" or "." or ".." || part.Contains(':')) || Path.IsPathRooted(relative)) throw new InvalidOperationException("Unsafe path in package inventory.");
			var full = Path.GetFullPath(Path.Combine(root, Path.Combine(parts)));
			if (!full.StartsWith(root, StringComparison.OrdinalIgnoreCase) || !covered.Add(full)) throw new InvalidOperationException("Duplicate or unsafe package inventory entry.");
			for (var current = full; current.Length >= root.Length; current = Path.GetDirectoryName(current) ?? string.Empty)
				if ((File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0) throw new InvalidOperationException("Package files must not be symbolic links or directory junctions.");
			var expected = item.GetProperty("sha256").GetString();
			if (expected is null || expected.Length != 64 || !expected.All(Uri.IsHexDigit)) throw new InvalidOperationException("Invalid package checksum.");
			using var stream = File.OpenRead(full);
			if (stream.Length != item.GetProperty("bytes").GetInt64() || !Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(stream)).Equals(expected, StringComparison.OrdinalIgnoreCase))
				throw new InvalidOperationException($"Package verification failed for {relative}. Extract a fresh copy of the complete Maily ZIP and retry.");
		}
		var required = new List<string> { "setup/install.ps1", "setup/uninstall.ps1", "setup/OutlookShutdown.ps1", "payload/addin/manifest.xml", "payload/runtime/atk/node.exe", "payload/runtime/atk/atk/cli.js", "payload/scripts/connect-outlook.cjs", "payload/scripts/outlook-credential-preflight.cjs" };
		if (!hosted) required.AddRange(["payload/app/Maily.OutlookAddin.Web/Maily.OutlookAddin.Web.exe", "payload/runtime/codex/codex.exe"]);
		foreach (var relative in required)
			if (!covered.Contains(Path.GetFullPath(Path.Combine(root, relative.Replace('/', Path.DirectorySeparatorChar))))) throw new InvalidOperationException($"Required file is not covered by package checksums: {relative}.");
		foreach (var subtree in new[] { "payload", "setup" })
			VerifyCoverage(Path.Combine(root, subtree), covered);
	}

	private static void VerifyCoverage(string directory, HashSet<string> covered)
	{
		if (!Directory.Exists(directory)) throw new InvalidOperationException("Required package folder is missing.");
		foreach (var entry in Directory.EnumerateFileSystemEntries(directory))
		{
			var attributes = File.GetAttributes(entry);
			if ((attributes & FileAttributes.ReparsePoint) != 0) throw new InvalidOperationException("Package links are not supported.");
			if ((attributes & FileAttributes.Directory) != 0) VerifyCoverage(entry, covered);
			else if (!covered.Contains(Path.GetFullPath(entry))) throw new InvalidOperationException("Unverified file found in package: " + Path.GetFileName(entry));
		}
	}
}

internal static class WindowsPowerShellBootstrap
{
	public static ProcessStartInfo Create(string arguments, string workingDirectory)
	{
		var windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
		var psHome = Path.Combine(windows, "System32", "WindowsPowerShell", "v1.0");
		var executable = Path.Combine(psHome, "powershell.exe");
		if (!File.Exists(executable)) throw new FileNotFoundException("Windows PowerShell is unavailable. Repair the Windows PowerShell installation before running Maily setup.", executable);
		var startInfo = new ProcessStartInfo
		{
			FileName = executable,
			Arguments = arguments,
			UseShellExecute = false,
			CreateNoWindow = true,
			RedirectStandardOutput = true,
			RedirectStandardError = true,
			WorkingDirectory = workingDirectory
		};
		// A .NET bootstrap can inherit PowerShell 7's module path. Windows PowerShell
		// then misses its own utility/CIM modules. Set only this child's environment.
		startInfo.Environment["PSModulePath"] = string.Join(Path.PathSeparator, new[]
		{
			Path.Combine(psHome, "Modules"),
			Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "WindowsPowerShell", "Modules"),
			Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "WindowsPowerShell", "Modules")
		});
		return startInfo;
	}
}

internal static class ScriptRunner
{
	public static ScriptResult Run(SetupOptions options, Action<string> log)
	{
		var root = ResolveRoot(options);
		var scriptPath = ResolveScriptPath(root, options.Uninstall ? "uninstall.ps1" : "install.ps1");
		if (!File.Exists(scriptPath))
		{
			log($"Setup script was not found: {scriptPath}");
			return new ScriptResult(1);
		}

		var arguments = BuildPowerShellArguments(scriptPath, options);
		log(options.Uninstall ? "Starting uninstall..." : "Starting install...");
		var progress = new SetupScriptProgress();

		using var process = new Process
		{
			StartInfo = WindowsPowerShellBootstrap.Create(arguments, options.Uninstall ? Path.GetTempPath() : root),
			EnableRaisingEvents = true
		};

		process.OutputDataReceived += (_, eventArgs) =>
		{
			if (!string.IsNullOrWhiteSpace(eventArgs.Data))
			{
				progress.Observe(eventArgs.Data);
				log(eventArgs.Data);
			}
		};
		process.ErrorDataReceived += (_, eventArgs) =>
		{
			if (!string.IsNullOrWhiteSpace(eventArgs.Data))
			{
				log(eventArgs.Data);
			}
		};

		process.Start();
		process.BeginOutputReadLine();
		process.BeginErrorReadLine();
		process.WaitForExit();

		LogSummary(options, log);
		return new ScriptResult(process.ExitCode, progress.ApplicationInstalled, progress.ActivationState);
	}

	private static string ResolveRoot(SetupOptions options)
	{
		if (options.Uninstall && !string.IsNullOrWhiteSpace(options.InstallRoot))
		{
			return Path.GetFullPath(options.InstallRoot);
		}

		return Path.GetFullPath(AppContext.BaseDirectory);
	}

	private static string ResolveScriptPath(string root, string scriptName)
	{
		var internalPath = Path.Combine(root, "setup", scriptName);
		if (File.Exists(internalPath))
		{
			return internalPath;
		}

		return Path.Combine(root, scriptName);
	}

	private static string BuildPowerShellArguments(string scriptPath, SetupOptions options)
	{
		var builder = new StringBuilder();
		var packageMode = PackageModeResolver.Resolve(options);
		builder.Append("-NoProfile -ExecutionPolicy Bypass -File ");
		builder.Append(Quote(scriptPath));
		builder.Append(" -Quiet");
		builder.Append(" -PackageMode ").Append(packageMode);
		if (options.VerifyOutlookActivation && !options.Uninstall) builder.Append(" -VerifyOutlookActivation");
		if (!string.IsNullOrWhiteSpace(options.InstallRoot)) builder.Append(" -InstallRoot ").Append(Quote(options.InstallRoot));
		if (options.Uninstall)
		{
			if (options.KeepUserData)
			{
				builder.Append(" -KeepUserData");
			}

			if (options.KeepLogs)
			{
				builder.Append(" -KeepLogs");
			}

			if (options.RemoveTrustedCertificate)
			{
				builder.Append(" -RemoveTrustedCertificate");
			}

			if (options.CloseOutlook)
			{
				builder.Append(" -CloseOutlook");
			}
		}
		else if (options.CloseOutlook)
		{
			builder.Append(" -CloseOutlook");
		}

		if (!options.Uninstall && (options.TrustLocalhostRoot || !options.Quiet) && string.Equals(packageMode, "Local", StringComparison.OrdinalIgnoreCase))
		{
			builder.Append(" -TrustLocalhostRoot");
		}

		return builder.ToString();
	}

	private static void LogSummary(SetupOptions options, Action<string> log)
	{
		try
		{
			var summaryPath = options.Uninstall ? FindLatestUninstallSummary() : FindInstallSummary();
			if (string.IsNullOrWhiteSpace(summaryPath) || !File.Exists(summaryPath))
			{
				return;
			}

			using var document = JsonDocument.Parse(File.ReadAllText(summaryPath));
			if (!document.RootElement.TryGetProperty("steps", out var steps) || steps.ValueKind != JsonValueKind.Array)
			{
				return;
			}

			log("");
			log("Setup checks:");
			foreach (var step in steps.EnumerateArray())
			{
				var name = step.TryGetProperty("name", out var nameElement) ? nameElement.GetString() : "step";
				var status = step.TryGetProperty("status", out var statusElement) ? statusElement.GetString() : "unknown";
				var detail = step.TryGetProperty("detail", out var detailElement) ? detailElement.GetString() : "";
				log($"- {name}: {status}. {detail}");
			}
		}
		catch
		{
			// Summary rendering is helpful, not critical.
		}
	}

	private static string? FindInstallSummary()
	{
		var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
		var path = Path.Combine(localAppData, "Maily", "OutlookAddin", "data", "install-summary.json");
		return File.Exists(path) ? path : null;
	}

	private static string? FindLatestUninstallSummary()
	{
		var temp = Path.GetTempPath();
		return Directory.EnumerateFiles(temp, "MailyOutlookAddin-uninstall-*.json")
			.OrderByDescending(File.GetLastWriteTimeUtc)
			.FirstOrDefault();
	}

	private static string Quote(string value) => "\"" + value.Replace("\"", "\\\"") + "\"";
}

internal sealed record ScriptResult(int ExitCode, bool ApplicationInstalled = false, string ActivationState = "unknown");
