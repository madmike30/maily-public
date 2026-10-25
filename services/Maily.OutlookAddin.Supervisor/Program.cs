using System.Diagnostics;
using System.Net.Http;
using System.ServiceProcess;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Hosting.WindowsServices;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

var builder = Host.CreateApplicationBuilder(new HostApplicationBuilderSettings
{
	Args = args,
	ContentRootPath = AppContext.BaseDirectory
});
builder.Logging.ClearProviders();
builder.Services.AddWindowsService(options =>
{
	options.ServiceName = "MailyOutlookAddinSupervisor";
});
builder.Services.Configure<MailySupervisorOptions>(builder.Configuration.GetSection("MailySupervisor"));
builder.Services.AddHostedService<OutlookAddinHostSupervisor>();

await builder.Build().RunAsync();

internal sealed class MailySupervisorOptions
{
	public string ServiceName { get; set; } = "MailyOutlookAddinSupervisor";
	public string HostServiceName { get; set; } = "MailyOutlookAddinHost";
	public string HealthUrl { get; set; } = "https://localhost:7097/api/health";
	public int PollSeconds { get; set; } = 5;
	public int ActivePollSeconds { get; set; } = 5;
	public int IdlePollSeconds { get; set; } = 3;
	public int HealthCheckSeconds { get; set; } = 10;
	public int StopAfterOutlookExitSeconds { get; set; } = 10;
	public int StartupTimeoutSeconds { get; set; } = 20;
	public int LogRetentionDays { get; set; } = 31;
	public bool PreferWindowsService { get; set; }
	public bool StopHostWhenOutlookExits { get; set; } = true;
	public bool ExitWhenOutlookExits { get; set; }
	public int ExitWhenOutlookNotRunningOnStartSeconds { get; set; }
	public string? RepoRoot { get; set; }
	public string? HostExecutablePath { get; set; }
	public string? HostWorkingDirectory { get; set; }
	public string? LogDirectory { get; set; }
}

internal sealed class OutlookAddinHostSupervisor : BackgroundService
{
	private const string MutexName = "Local\\MailyOutlookAddinSupervisor";
	private readonly IOptionsMonitor<MailySupervisorOptions> optionsMonitor;
	private readonly ILogger<OutlookAddinHostSupervisor> logger;
	private readonly IHostApplicationLifetime lifetime;
	private readonly HttpClient httpClient;
	private readonly SemaphoreSlim logLock = new(1, 1);
	private readonly DateTimeOffset startedAtUtc = DateTimeOffset.UtcNow;
	private DateTimeOffset lastOutlookSeenUtc = DateTimeOffset.MinValue;
	private DateTimeOffset lastLogPruneUtc = DateTimeOffset.MinValue;
	private DateTimeOffset lastHostHealthCheckUtc = DateTimeOffset.MinValue;
	private bool lastHostReady;

	public OutlookAddinHostSupervisor(
		IOptionsMonitor<MailySupervisorOptions> optionsMonitor,
		ILogger<OutlookAddinHostSupervisor> logger,
		IHostApplicationLifetime lifetime)
	{
		this.optionsMonitor = optionsMonitor;
		this.logger = logger;
		this.lifetime = lifetime;
		var handler = new HttpClientHandler
		{
			ServerCertificateCustomValidationCallback = HttpClientHandler.DangerousAcceptAnyServerCertificateValidator
		};
		httpClient = new HttpClient(handler)
		{
			Timeout = TimeSpan.FromSeconds(5)
		};
	}

	protected override async Task ExecuteAsync(CancellationToken stoppingToken)
	{
		var createdNew = false;
		using var mutex = new Mutex(true, MutexName, out createdNew);
		if (!createdNew)
		{
			await WriteLogAsync("Another supervisor instance is already running. Exiting.", stoppingToken);
			lifetime.StopApplication();
			return;
		}

		await WriteLogAsync("Supervisor started.", stoppingToken);
		try
		{
			while (!stoppingToken.IsCancellationRequested)
			{
				var options = NormalizeOptions(optionsMonitor.CurrentValue);
				await PruneLogsIfNeededAsync(options, stoppingToken);
				var outlookRunning = IsOutlookRunning();
				if (outlookRunning)
				{
					lastOutlookSeenUtc = DateTimeOffset.UtcNow;
					if (!await IsHostReadyCachedAsync(options, stoppingToken))
					{
						await WriteLogAsync("Outlook is running and host health is not ready. Starting host.", stoppingToken);
						await StartHostAsync(options, stoppingToken);
					}
				}
				else if (options.StopHostWhenOutlookExits && ShouldStopHost(options))
				{
					await WriteLogAsync("Outlook is not running. Stopping host after idle grace period.", stoppingToken);
					await StopHostAsync(options, stoppingToken);
					if (options.ExitWhenOutlookExits)
					{
						await WriteLogAsync("Outlook has exited. Stopping supervisor.", stoppingToken);
						lifetime.StopApplication();
						return;
					}

					lastOutlookSeenUtc = DateTimeOffset.MinValue;
				}
				else if (ShouldExitBecauseOutlookNeverStarted(options))
				{
					await WriteLogAsync("Outlook did not start during the supervisor startup grace period. Stopping supervisor.", stoppingToken);
					lifetime.StopApplication();
					return;
				}

				var delaySeconds = outlookRunning
					? Math.Max(1, options.ActivePollSeconds)
					: Math.Max(1, options.IdlePollSeconds);
				await Task.Delay(TimeSpan.FromSeconds(delaySeconds), stoppingToken);
			}
		}
		catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
		{
		}
		catch (Exception ex)
		{
			logger.LogError(ex, "Maily Outlook add-in supervisor stopped unexpectedly.");
			await WriteLogAsync($"Supervisor fatal error: {ex.Message}", CancellationToken.None);
			throw;
		}
		finally
		{
			await WriteLogAsync("Supervisor stopped.", CancellationToken.None);
		}
	}

	public override void Dispose()
	{
		httpClient.Dispose();
		logLock.Dispose();
		base.Dispose();
	}

	private bool ShouldStopHost(MailySupervisorOptions options) =>
		lastOutlookSeenUtc != DateTimeOffset.MinValue &&
		DateTimeOffset.UtcNow - lastOutlookSeenUtc >= TimeSpan.FromSeconds(Math.Max(5, options.StopAfterOutlookExitSeconds));

	private bool ShouldExitBecauseOutlookNeverStarted(MailySupervisorOptions options) =>
		options.ExitWhenOutlookExits &&
		lastOutlookSeenUtc == DateTimeOffset.MinValue &&
		DateTimeOffset.UtcNow - startedAtUtc >= TimeSpan.FromSeconds(Math.Max(5, options.ExitWhenOutlookNotRunningOnStartSeconds));

	private static bool IsOutlookRunning()
	{
		var processes = Process.GetProcessesByName("OUTLOOK");
		try
		{
			return processes.Length > 0;
		}
		finally
		{
			foreach (var process in processes)
			{
				process.Dispose();
			}
		}
	}

	private async Task<bool> IsHostReadyAsync(MailySupervisorOptions options, CancellationToken cancellationToken)
	{
		try
		{
			using var response = await httpClient.GetAsync(options.HealthUrl, cancellationToken);
			return response.IsSuccessStatusCode;
		}
		catch
		{
			return false;
		}
	}

	private async Task<bool> IsHostReadyCachedAsync(MailySupervisorOptions options, CancellationToken cancellationToken)
	{
		var now = DateTimeOffset.UtcNow;
		if (now - lastHostHealthCheckUtc < TimeSpan.FromSeconds(Math.Max(1, options.HealthCheckSeconds)))
		{
			return lastHostReady;
		}

		lastHostReady = await IsHostReadyAsync(options, cancellationToken);
		lastHostHealthCheckUtc = now;
		return lastHostReady;
	}

	private async Task StartHostAsync(MailySupervisorOptions options, CancellationToken cancellationToken)
	{
		if (options.PreferWindowsService && TryGetService(options.HostServiceName, out var service))
		{
			using (service)
			{
				await StartHostServiceAsync(service, options, cancellationToken);
			}

			return;
		}

		if (!IsHostProcessRunning(options))
		{
			var startInfo = new ProcessStartInfo
			{
				FileName = options.HostExecutablePath!,
				WorkingDirectory = options.HostWorkingDirectory!,
				UseShellExecute = false,
				CreateNoWindow = true,
				WindowStyle = ProcessWindowStyle.Hidden
			};
			startInfo.Environment.TryAdd("ASPNETCORE_ENVIRONMENT", "Production");
			Process.Start(startInfo);
			await WriteLogAsync($"Started hidden host process: {options.HostExecutablePath}", cancellationToken);
		}

		await WaitForHostReadyAsync(options, cancellationToken);
	}

	private async Task StartHostServiceAsync(ServiceController service, MailySupervisorOptions options, CancellationToken cancellationToken)
	{
		if (service.Status is ServiceControllerStatus.Running or ServiceControllerStatus.StartPending)
		{
			await WriteLogAsync($"Host service {options.HostServiceName} is already {service.Status}.", cancellationToken);
			await WaitForHostReadyAsync(options, cancellationToken);
			return;
		}

		service.Start();
		await WriteLogAsync($"Started host Windows service: {options.HostServiceName}", cancellationToken);
		await WaitForHostReadyAsync(options, cancellationToken);
	}

	private async Task WaitForHostReadyAsync(MailySupervisorOptions options, CancellationToken cancellationToken)
	{
		var deadline = DateTimeOffset.UtcNow.AddSeconds(Math.Max(5, options.StartupTimeoutSeconds));
		while (DateTimeOffset.UtcNow < deadline && !cancellationToken.IsCancellationRequested)
		{
			if (await IsHostReadyAsync(options, cancellationToken))
			{
				lastHostReady = true;
				lastHostHealthCheckUtc = DateTimeOffset.UtcNow;
				await WriteLogAsync("Host health check is ready.", cancellationToken);
				return;
			}

			await Task.Delay(TimeSpan.FromSeconds(1), cancellationToken);
		}

		await WriteLogAsync("Host did not become ready before startup timeout.", cancellationToken);
	}

	private async Task StopHostAsync(MailySupervisorOptions options, CancellationToken cancellationToken)
	{
		if (options.PreferWindowsService && TryGetService(options.HostServiceName, out var service))
		{
			using (service)
			{
				if (service.Status is ServiceControllerStatus.Running or ServiceControllerStatus.StartPending)
				{
					service.Stop();
					await WriteLogAsync($"Stopped host Windows service: {options.HostServiceName}", cancellationToken);
				}
			}

			return;
		}

		foreach (var process in GetHostProcesses(options))
		{
			try
			{
				process.Kill(entireProcessTree: true);
				await WriteLogAsync($"Stopped hidden host process id={process.Id}.", cancellationToken);
			}
			catch (Exception ex)
			{
				await WriteLogAsync($"Could not stop host process id={process.Id}: {ex.Message}", cancellationToken);
			}
			finally
			{
				process.Dispose();
			}
		}
	}

	private bool IsHostProcessRunning(MailySupervisorOptions options)
	{
		var processes = GetHostProcesses(options);
		try
		{
			return processes.Count > 0;
		}
		finally
		{
			foreach (var process in processes)
			{
				process.Dispose();
			}
		}
	}

	private static List<Process> GetHostProcesses(MailySupervisorOptions options)
	{
		var managedPaths = GetManagedHostExecutablePaths(options);
		var result = new List<Process>();
		foreach (var process in Process.GetProcessesByName("Maily.OutlookAddin.Web"))
		{
			try
			{
				var path = process.MainModule?.FileName;
				if (string.IsNullOrWhiteSpace(path) ||
					managedPaths.Contains(Path.GetFullPath(path)))
				{
					result.Add(process);
				}
				else
				{
					process.Dispose();
				}
			}
			catch
			{
				result.Add(process);
			}
		}

		return result;
	}

	private static HashSet<string> GetManagedHostExecutablePaths(MailySupervisorOptions options)
	{
		var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		AddPath(options.HostExecutablePath);
		if (!string.IsNullOrWhiteSpace(options.RepoRoot))
		{
			AddPath(Path.Combine(options.RepoRoot, "services", "Maily.OutlookAddin.Web", "bin", "Release", "net10.0", "Maily.OutlookAddin.Web.exe"));
			AddPath(Path.Combine(options.RepoRoot, "services", "Maily.OutlookAddin.Web", "bin", "Debug", "net10.0", "Maily.OutlookAddin.Web.exe"));
			AddPath(Path.Combine(options.RepoRoot, "output", "startup", "Maily.OutlookAddin.Web", "Maily.OutlookAddin.Web.exe"));
			AddPath(Path.Combine(options.RepoRoot, "output", "service", "Maily.OutlookAddin.Web", "Maily.OutlookAddin.Web.exe"));
		}

		return paths;

		void AddPath(string? path)
		{
			if (!string.IsNullOrWhiteSpace(path))
			{
				paths.Add(Path.GetFullPath(path));
			}
		}
	}

	private static bool TryGetService(string serviceName, out ServiceController service)
	{
		service = new ServiceController(serviceName);
		try
		{
			_ = service.Status;
			return true;
		}
		catch
		{
			service.Dispose();
			service = null!;
			return false;
		}
	}

	private MailySupervisorOptions NormalizeOptions(MailySupervisorOptions source)
	{
		var repoRoot = ResolveRepoRoot(source.RepoRoot);
		var hostExecutablePath = ResolveHostExecutablePath(source.HostExecutablePath, repoRoot);
		return new MailySupervisorOptions
		{
			ServiceName = string.IsNullOrWhiteSpace(source.ServiceName) ? "MailyOutlookAddinSupervisor" : source.ServiceName,
			HostServiceName = string.IsNullOrWhiteSpace(source.HostServiceName) ? "MailyOutlookAddinHost" : source.HostServiceName,
			HealthUrl = string.IsNullOrWhiteSpace(source.HealthUrl) ? "https://localhost:7097/api/health" : source.HealthUrl,
			PollSeconds = source.PollSeconds <= 0 ? 5 : source.PollSeconds,
			ActivePollSeconds = source.ActivePollSeconds <= 0 ? Math.Max(1, source.PollSeconds <= 0 ? 5 : source.PollSeconds) : source.ActivePollSeconds,
			IdlePollSeconds = source.IdlePollSeconds <= 0 ? Math.Max(3, source.PollSeconds <= 0 ? 3 : source.PollSeconds) : source.IdlePollSeconds,
			HealthCheckSeconds = source.HealthCheckSeconds <= 0 ? 10 : source.HealthCheckSeconds,
			StopAfterOutlookExitSeconds = source.StopAfterOutlookExitSeconds <= 0 ? 10 : source.StopAfterOutlookExitSeconds,
			StartupTimeoutSeconds = source.StartupTimeoutSeconds <= 0 ? 20 : source.StartupTimeoutSeconds,
			LogRetentionDays = source.LogRetentionDays <= 0 ? 31 : source.LogRetentionDays,
			PreferWindowsService = source.PreferWindowsService,
			StopHostWhenOutlookExits = source.StopHostWhenOutlookExits,
			RepoRoot = repoRoot,
			HostExecutablePath = hostExecutablePath,
			HostWorkingDirectory = string.IsNullOrWhiteSpace(source.HostWorkingDirectory)
				? Path.GetDirectoryName(hostExecutablePath)
				: Path.GetFullPath(source.HostWorkingDirectory),
			LogDirectory = string.IsNullOrWhiteSpace(source.LogDirectory)
				? Path.Combine(repoRoot, "output", "logs")
				: Path.GetFullPath(source.LogDirectory)
		};
	}

	private static string ResolveRepoRoot(string? configuredRepoRoot)
	{
		if (!string.IsNullOrWhiteSpace(configuredRepoRoot) && Directory.Exists(configuredRepoRoot))
		{
			return Path.GetFullPath(configuredRepoRoot);
		}

		var candidates = new[]
		{
			Directory.GetCurrentDirectory(),
			AppContext.BaseDirectory
		};

		foreach (var start in candidates.Where(static value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase))
		{
			var current = new DirectoryInfo(start);
			while (current is not null)
			{
				if (File.Exists(Path.Combine(current.FullName, "maily.sln")) &&
					Directory.Exists(Path.Combine(current.FullName, "services", "Maily.OutlookAddin.Web")))
				{
					return current.FullName;
				}

				current = current.Parent;
			}
		}

		return Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
	}

	private static string ResolveHostExecutablePath(string? configuredPath, string repoRoot)
	{
		if (!string.IsNullOrWhiteSpace(configuredPath) && File.Exists(configuredPath))
		{
			return Path.GetFullPath(configuredPath);
		}

		var candidates = new[]
		{
			Path.Combine(repoRoot, "output", "startup", "Maily.OutlookAddin.Web", "Maily.OutlookAddin.Web.exe"),
			Path.Combine(repoRoot, "output", "service", "Maily.OutlookAddin.Web", "Maily.OutlookAddin.Web.exe"),
			Path.Combine(repoRoot, "services", "Maily.OutlookAddin.Web", "bin", "Release", "net10.0", "Maily.OutlookAddin.Web.exe"),
			Path.Combine(repoRoot, "services", "Maily.OutlookAddin.Web", "bin", "Debug", "net10.0", "Maily.OutlookAddin.Web.exe")
		};

		var resolved = candidates.FirstOrDefault(File.Exists);
		if (string.IsNullOrWhiteSpace(resolved))
		{
			throw new FileNotFoundException("Could not find the Maily Outlook add-in host executable. Build services\\Maily.OutlookAddin.Web first.");
		}

		return Path.GetFullPath(resolved);
	}

	private async Task PruneLogsIfNeededAsync(MailySupervisorOptions options, CancellationToken cancellationToken)
	{
		if (DateTimeOffset.UtcNow - lastLogPruneUtc < TimeSpan.FromHours(12))
		{
			return;
		}

		lastLogPruneUtc = DateTimeOffset.UtcNow;
		if (!Directory.Exists(options.LogDirectory))
		{
			return;
		}

		var cutoffUtc = DateTime.UtcNow.AddDays(-Math.Max(1, options.LogRetentionDays));
		var deleted = 0;
		foreach (var path in Directory.EnumerateFiles(options.LogDirectory, "*.log").Concat(Directory.EnumerateFiles(options.LogDirectory, "*.jsonl")))
		{
			try
			{
				if (File.GetLastWriteTimeUtc(path) < cutoffUtc)
				{
					File.Delete(path);
					deleted++;
				}
			}
			catch
			{
			}
		}

		if (deleted > 0)
		{
			await WriteLogAsync($"Pruned {deleted} expired log files.", cancellationToken);
		}
	}

	private async Task WriteLogAsync(string message, CancellationToken cancellationToken)
	{
		MailySupervisorOptions options;
		try
		{
			options = NormalizeOptions(optionsMonitor.CurrentValue);
		}
		catch
		{
			options = new MailySupervisorOptions
			{
				LogDirectory = Path.Combine(AppContext.BaseDirectory, "logs")
			};
		}

		var logDirectory = options.LogDirectory ?? Path.Combine(AppContext.BaseDirectory, "logs");
		Directory.CreateDirectory(logDirectory);
		var logPath = Path.Combine(logDirectory, $"addin-supervisor-{DateTime.UtcNow:yyyyMMdd}.log");
		var line = $"{DateTimeOffset.UtcNow:o}\t{message}";

		await logLock.WaitAsync(cancellationToken);
		try
		{
			await File.AppendAllTextAsync(logPath, line + Environment.NewLine, cancellationToken);
		}
		catch
		{
			// Logging must never break add-in startup.
		}
		finally
		{
			logLock.Release();
		}
	}
}
