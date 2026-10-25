using System.Collections.ObjectModel;
using System.Diagnostics;
using Microsoft.Extensions.Logging;
using Maily.Generation.Models;

namespace Maily.Generation.Services;

public sealed class OracleResponsesGateway
{
	private static readonly TimeSpan ExecutionTimeout = TimeSpan.FromSeconds(210);

	private readonly ILogger<OracleResponsesGateway> logger;
	private readonly CodexModelCatalog? catalog;

	public OracleResponsesGateway(ILogger<OracleResponsesGateway> logger, CodexModelCatalog? catalog = null)
	{
		this.logger = logger;
		this.catalog = catalog;
	}

	public async Task<string> GenerateReplyAsync(
		CodexRuntimeConfig config,
		string systemPrompt,
		string userPrompt,
		string? requestId = null,
		CancellationToken cancellationToken = default)
	{
		if (catalog is not null)
			CodexModelCatalog.ApplySelection(config, await catalog.GetAsync(config, cancellationToken));
		var elapsed = Stopwatch.StartNew();
		logger.LogInformation(
			"codex_exec_started requestId={RequestId} model={Model} systemChars={SystemChars} userChars={UserChars}",
			requestId ?? string.Empty,
			string.IsNullOrWhiteSpace(config.Model) ? "(default)" : config.Model,
			systemPrompt?.Length ?? 0,
			userPrompt?.Length ?? 0);

		var prompt = BuildCodexPrompt(systemPrompt ?? string.Empty, userPrompt ?? string.Empty);
		using var process = new Process
		{
			StartInfo = BuildStartInfo(config),
			EnableRaisingEvents = true
		};

		using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
		timeout.CancelAfter(ExecutionTimeout);
		Task<string>? stdoutTask = null;
		Task<string>? stderrTask = null;

		try
		{
			if (!process.Start())
			{
				throw new InvalidOperationException("Codex CLI did not start.");
			}

			stdoutTask = process.StandardOutput.ReadToEndAsync(timeout.Token);
			stderrTask = process.StandardError.ReadToEndAsync(timeout.Token);

			await process.StandardInput.WriteAsync(prompt.AsMemory(), timeout.Token);
			await process.StandardInput.FlushAsync(timeout.Token);
			process.StandardInput.Close();

			await process.WaitForExitAsync(timeout.Token);
			var stdout = await stdoutTask.WaitAsync(timeout.Token);
			var stderr = await stderrTask.WaitAsync(timeout.Token);

			if (process.ExitCode != 0)
			{
				logger.LogWarning(
					"codex_exec_failed requestId={RequestId} model={Model} exitCode={ExitCode} recovery={Recovery}",
					requestId ?? string.Empty,
					string.IsNullOrWhiteSpace(config.Model) ? "(default)" : config.Model,
					process.ExitCode,
					ClassifyRecovery(stderr));
				throw new InvalidOperationException(
					$"Codex CLI generation failed (exit {process.ExitCode}). {RecoveryMessage(stderr)}");
			}

			var output = NormalizeCodexOutput(stdout);
			if (string.IsNullOrWhiteSpace(output))
			{
				logger.LogWarning(
					"codex_exec_empty_output requestId={RequestId} model={Model} recovery={Recovery}",
					requestId ?? string.Empty,
					string.IsNullOrWhiteSpace(config.Model) ? "(default)" : config.Model,
					ClassifyRecovery(stderr));
				throw new InvalidOperationException(
					$"Codex CLI completed without reply text. {RecoveryMessage(stderr)}");
			}

			logger.LogInformation(
				"codex_exec_completed requestId={RequestId} model={Model} outputChars={OutputChars} recovery={Recovery} elapsedMs={ElapsedMs}",
				requestId ?? string.Empty,
				string.IsNullOrWhiteSpace(config.Model) ? "(default)" : config.Model,
				output.Length,
				ClassifyRecovery(stderr), elapsed.ElapsedMilliseconds);
			return output;
		}
		catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
		{
			throw new TimeoutException($"Codex CLI generation timed out after {(int)ExecutionTimeout.TotalSeconds} seconds.");
		}
		catch (OperationCanceledException)
		{
			throw;
		}
		catch (Exception ex) when (ex is System.ComponentModel.Win32Exception or FileNotFoundException)
		{
			if (ex.Message.Contains("Access is denied", StringComparison.OrdinalIgnoreCase))
			{
				throw new InvalidOperationException(
					"Windows denied execution of Codex CLI. Configure a supported standalone Codex CLI executable.");
			}

			throw new InvalidOperationException(
				"Codex CLI was not found. Make sure Codex is installed or configure MailyGeneration:CodexExecutablePath.");
		}
		finally
		{
			// A failed write must not leave a child or its redirected readers behind.
			timeout.Cancel();
			await StopProcessTreeAsync(process);
			foreach (var reader in new[] { stdoutTask, stderrTask })
			{
				if (reader is null) continue;
				try { await reader.WaitAsync(TimeSpan.FromSeconds(2)); }
				catch { /* Preserve the original generation error; readers are cancellation-aware. */ }
			}
		}
	}

	private static ProcessStartInfo BuildStartInfo(CodexRuntimeConfig config)
	{
		var executablePath = string.IsNullOrWhiteSpace(config.CodexExecutablePath) ? "codex" : config.CodexExecutablePath;
		var startInfo = new ProcessStartInfo
		{
			FileName = executablePath,
			WorkingDirectory = Path.GetTempPath(),
			RedirectStandardInput = true,
			RedirectStandardOutput = true,
			RedirectStandardError = true,
			StandardInputEncoding = new System.Text.UTF8Encoding(false),
			StandardOutputEncoding = new System.Text.UTF8Encoding(false),
			StandardErrorEncoding = new System.Text.UTF8Encoding(false),
			UseShellExecute = false,
			CreateNoWindow = true
		};
		if (executablePath.EndsWith(".cmd", StringComparison.OrdinalIgnoreCase) ||
			executablePath.EndsWith(".bat", StringComparison.OrdinalIgnoreCase))
		{
			startInfo.FileName = Environment.GetEnvironmentVariable("COMSPEC") ?? "cmd.exe";
			startInfo.Arguments = BuildCodexCommandLine(executablePath, config);
		}
		else
		{
			AddCodexExecArguments(startInfo, config);
		}

		if (!string.IsNullOrWhiteSpace(config.CodexHome))
		{
			startInfo.Environment["CODEX_HOME"] = config.CodexHome;
		}

		startInfo.Environment["NO_COLOR"] = "1";
		return startInfo;
	}

	private static void AddCodexExecArguments(ProcessStartInfo startInfo, CodexRuntimeConfig config)
	{
		startInfo.ArgumentList.Add("exec");
		startInfo.ArgumentList.Add("--ephemeral");
		startInfo.ArgumentList.Add("--skip-git-repo-check");
		startInfo.ArgumentList.Add("--sandbox");
		startInfo.ArgumentList.Add("read-only");
		startInfo.ArgumentList.Add("--ignore-user-config");
		startInfo.ArgumentList.Add("--ignore-rules");
		AddDisabledCodexFeature(startInfo.ArgumentList, "plugins");
		AddDisabledCodexFeature(startInfo.ArgumentList, "apps");
		AddDisabledCodexFeature(startInfo.ArgumentList, "browser_use");
		AddDisabledCodexFeature(startInfo.ArgumentList, "browser_use_external");
		AddDisabledCodexFeature(startInfo.ArgumentList, "computer_use");
		AddDisabledCodexFeature(startInfo.ArgumentList, "tool_search");
		AddDisabledCodexFeature(startInfo.ArgumentList, "shell_snapshot");
		// Email is untrusted input. A read-only sandbox still permits reading local
		// files, so remove tool capabilities instead of relying on prompt wording.
		foreach (var feature in DisabledGenerationCapabilities) AddDisabledCodexFeature(startInfo.ArgumentList, feature);
		AddGenerationIsolationArguments(startInfo.ArgumentList);
		startInfo.ArgumentList.Add("--color");
		startInfo.ArgumentList.Add("never");
		if (!string.IsNullOrWhiteSpace(config.Model))
		{
			startInfo.ArgumentList.Add("-m");
			startInfo.ArgumentList.Add(config.Model);
		}
		if (!string.IsNullOrWhiteSpace(config.ReasoningEffort))
		{
			startInfo.ArgumentList.Add("-c");
			startInfo.ArgumentList.Add($"model_reasoning_effort=\"{config.ReasoningEffort}\"");
		}
		startInfo.ArgumentList.Add("-");
	}

	private static string BuildCodexCommandLine(string executablePath, CodexRuntimeConfig config)
	{
		var arguments = new List<string>
		{
			executablePath,
			"exec",
			"--ephemeral",
			"--skip-git-repo-check",
			"--sandbox",
			"read-only",
			"--ignore-user-config",
			"--ignore-rules",
			"--disable",
			"plugins",
			"--disable",
			"apps",
			"--disable",
			"browser_use",
			"--disable",
			"browser_use_external",
			"--disable",
			"computer_use",
			"--disable",
			"tool_search",
			"--disable",
			"shell_snapshot",
		};
		foreach (var feature in DisabledGenerationCapabilities) { arguments.Add("--disable"); arguments.Add(feature); }
		AddGenerationIsolationArguments(arguments);
		arguments.Add("--color");
		arguments.Add("never");
		if (!string.IsNullOrWhiteSpace(config.Model))
		{
			arguments.Add("-m");
			arguments.Add(config.Model);
		}
		if (!string.IsNullOrWhiteSpace(config.ReasoningEffort))
		{
			arguments.Add("-c");
			arguments.Add($"model_reasoning_effort=\"{config.ReasoningEffort}\"");
		}
		arguments.Add("-");
		return WindowsBatchCommand.BuildArguments(arguments);
	}

	private static void AddDisabledCodexFeature(Collection<string> arguments, string feature)
	{
		arguments.Add("--disable");
		arguments.Add(feature);
	}

	private static readonly string[] DisabledGenerationCapabilities =
	[
		"shell_tool", "view_image", "multi_agent", "image_generation", "hooks",
		"workspace_dependencies", "skill_search", "memories", "code_mode", "code_mode_host", "tool_suggest"
	];

	private static void AddGenerationIsolationArguments(ICollection<string> arguments)
	{
		arguments.Add("--enable");
		arguments.Add("skip_host_skill_discovery");
		foreach (var setting in new[] { "web_search=\"disabled\"", "tools.view_image=false", "project_doc_max_bytes=0", "approval_policy=\"never\"" })
		{
			arguments.Add("-c");
			arguments.Add(setting);
		}
	}


	private static string BuildCodexPrompt(string systemPrompt, string userPrompt) =>
		$"""
		Return only the final answer. Do not use tools, markdown fences, labels, diagnostics, or commentary.
		If JSON is requested, return one valid JSON object only.

		System:
		{systemPrompt}

		Task:
		{userPrompt}
		""";

	private static string NormalizeCodexOutput(string stdout)
	{
		var output = stdout.Trim();
		if (!output.StartsWith("```", StringComparison.Ordinal))
		{
			return output;
		}

		var lines = output.Split(["\r\n", "\n"], StringSplitOptions.None);
		if (lines.Length >= 2 && lines[^1].Trim() == "```")
		{
			var contentLines = lines.Skip(1).Take(lines.Length - 2);
			return string.Join(Environment.NewLine, contentLines).Trim();
		}

		return output;
	}

	private static async Task StopProcessTreeAsync(Process process)
	{
		try
		{
			if (!process.HasExited)
			{
				process.Kill(entireProcessTree: true);
			}
			await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(3));
		}
		catch
		{
		}
	}

	private static string ClassifyRecovery(string stderr)
	{
		if (stderr.Contains("not supported", StringComparison.OrdinalIgnoreCase)) return "model-unavailable";
		if (stderr.Contains("401", StringComparison.OrdinalIgnoreCase) || stderr.Contains("sign in", StringComparison.OrdinalIgnoreCase) || stderr.Contains("login", StringComparison.OrdinalIgnoreCase)) return "sign-in-required";
		if (stderr.Contains("429", StringComparison.OrdinalIgnoreCase) || stderr.Contains("rate limit", StringComparison.OrdinalIgnoreCase)) return "usage-limit";
		return "retry";
	}

	private static string RecoveryMessage(string stderr) => ClassifyRecovery(stderr) switch
	{
		"model-unavailable" => "Choose a model supported by your account.",
		"sign-in-required" => "Complete ChatGPT sign-in in Codex and retry.",
		"usage-limit" => "Your account usage limit was reached. Retry later.",
		_ => "Please retry. Provider diagnostics are withheld to protect message content."
	};
}
