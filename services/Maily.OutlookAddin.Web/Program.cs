using Maily.OutlookAddin.Web.Services;
using Maily.Generation.Contracts;
using Maily.Generation.Models;
using Maily.Generation.Services;
using Maily.OutlookAddin.Web.Logging;
using Maily.OutlookAddin.Web.Models;
using System.Diagnostics;
using System.Text.Json;
using Microsoft.AspNetCore.HttpLogging;
using Microsoft.AspNetCore.Server.Kestrel.Core;
using Microsoft.Extensions.Hosting.WindowsServices;

var hostRoot = ResolveHostRoot();
var builder = WebApplication.CreateBuilder(new WebApplicationOptions
{
	Args = args,
	ContentRootPath = hostRoot,
	WebRootPath = Path.Combine(hostRoot, "wwwroot")
});
var isWindowsService = WindowsServiceHelpers.IsWindowsService();

var hostingOptions = new HostingOptions();
builder.Configuration.GetSection("MailyHosting").Bind(hostingOptions);

var generationOptions = new GenerationOptions();
builder.Configuration.GetSection("MailyGeneration").Bind(generationOptions);
var repoRoot = ResolveRepoRoot(hostRoot, generationOptions.RepoRoot);
var logsRoot = string.IsNullOrWhiteSpace(hostingOptions.LogDirectory)
	? Path.GetFullPath(Path.Combine(repoRoot, "output", "logs"))
	: Path.GetFullPath(hostingOptions.LogDirectory);

builder.Host.UseWindowsService(options =>
{
	options.ServiceName = string.IsNullOrWhiteSpace(hostingOptions.ServiceName)
		? "MailyOutlookAddinHost"
		: hostingOptions.ServiceName;
});

builder.Logging.ClearProviders();
var addinLogPath = Path.GetFullPath(Path.Combine(logsRoot, $"addin-web-{DateTime.UtcNow:yyyyMMdd}.jsonl"));
Directory.CreateDirectory(logsRoot);
var deletedExpiredLogs = DeleteExpiredLogFiles(logsRoot, TimeSpan.FromDays(31));
builder.Logging.AddProvider(new JsonlFileLoggerProvider(() =>
	Path.GetFullPath(Path.Combine(logsRoot, $"addin-web-{DateTime.UtcNow:yyyyMMdd}.jsonl"))));
if (OperatingSystem.IsWindows() && isWindowsService && hostingOptions.EnableEventLog)
{
#pragma warning disable CA1416
	builder.Logging.AddEventLog(settings =>
	{
		settings.SourceName = string.IsNullOrWhiteSpace(hostingOptions.ServiceName)
			? "MailyOutlookAddinHost"
			: hostingOptions.ServiceName;
	});
#pragma warning restore CA1416
}

if (!isWindowsService)
{
	builder.Logging.AddSimpleConsole(options =>
	{
		options.SingleLine = true;
		options.TimestampFormat = "HH:mm:ss ";
	});
}

builder.Services.AddMailyGenerationCore(options =>
{
	options.RepoRoot = repoRoot;
	options.PromptsRoot = generationOptions.PromptsRoot;
	options.SettingsFilePath = generationOptions.SettingsFilePath;
	options.CodexHome = generationOptions.CodexHome;
	options.CodexExecutablePath = generationOptions.CodexExecutablePath;
	options.Model = generationOptions.Model;
	options.PromptVersion = string.IsNullOrWhiteSpace(generationOptions.PromptVersion) ? "v1" : generationOptions.PromptVersion;
});
builder.Services.AddSingleton<OutlookSettingsCatalog>();
builder.Services.AddHostedService<CodexCatalogWarmup>();
builder.Services.AddSingleton<OutlookSearchSessionStore>();
builder.Services.AddSingleton<OutlookSearchFetchCache>();
var enableHttpLogging = builder.Environment.IsDevelopment() && !isWindowsService;
if (enableHttpLogging)
{
	builder.Services.AddHttpLogging(options =>
	{
		options.LoggingFields = HttpLoggingFields.RequestMethod |
			HttpLoggingFields.RequestPath |
			HttpLoggingFields.ResponseStatusCode |
			HttpLoggingFields.Duration;
	});
}
builder.WebHost.ConfigureKestrel((context, options) =>
{
	options.ConfigureEndpointDefaults(endpointOptions =>
	{
		endpointOptions.Protocols = HttpProtocols.Http1AndHttp2;
	});

	var configuredHosting = new HostingOptions();
	context.Configuration.GetSection("MailyHosting").Bind(configuredHosting);

	if (configuredHosting.EnableHttp && Uri.TryCreate(configuredHosting.HttpUrl, UriKind.Absolute, out var httpUri))
	{
		options.ListenLocalhost(httpUri.Port);
	}

	if (Uri.TryCreate(configuredHosting.HttpsUrl, UriKind.Absolute, out var httpsUri))
	{
		if (!string.IsNullOrWhiteSpace(configuredHosting.CertificatePath) &&
			File.Exists(configuredHosting.CertificatePath))
		{
			options.ListenLocalhost(httpsUri.Port, listenOptions =>
			{
				listenOptions.UseHttps(configuredHosting.CertificatePath, configuredHosting.CertificatePassword);
			});
		}
		else
		{
			options.ListenLocalhost(httpsUri.Port, listenOptions => listenOptions.UseHttps());
		}
	}
});

var app = builder.Build();

var backendApiKey = builder.Configuration["MailyGeneration:BackendApiKey"];
var enableRequestLogging = hostingOptions.EnableRequestLog || builder.Environment.IsDevelopment();
var startupRequestLogPath = enableRequestLogging ? GetCurrentRequestLogPath(logsRoot) : string.Empty;
if (enableRequestLogging)
{
	Directory.CreateDirectory(Path.GetDirectoryName(startupRequestLogPath)!);
}
var requestLogLock = new SemaphoreSlim(1, 1);

static async Task AppendRequestLogAsync(
	string logDirectory,
	SemaphoreSlim logLock,
	HttpContext context,
	int statusCode,
	Exception? exception = null)
{
	var path = GetCurrentRequestLogPath(logDirectory);
	var line = string.Join("\t", new[]
	{
		DateTimeOffset.UtcNow.ToString("o"),
		context.Request.Method,
		context.Request.Path.ToString(),
		"", // Query strings may contain mailbox identities and message searches.
		statusCode.ToString(),
		exception?.GetType().Name ?? "",
		"" // Exception messages may include provider or mailbox content.
	});

	await logLock.WaitAsync();
	try
	{
		Directory.CreateDirectory(Path.GetDirectoryName(path)!);
		await using var stream = new FileStream(
			path,
			FileMode.Append,
			FileAccess.Write,
			FileShare.ReadWrite,
			4096,
			true);
		await using var writer = new StreamWriter(stream);
		await writer.WriteLineAsync(line);
	}
	catch
	{
		// Request logging must never break pane delivery.
	}
	finally
	{
		logLock.Release();
	}
}

static async Task<string> InvokeOutlookBridgeAsync(
	string repoRootPath,
	string action,
	IReadOnlyDictionary<string, string>? arguments,
	CancellationToken cancellationToken)
{
	var scriptPath = Path.GetFullPath(Path.Combine(repoRootPath, "scripts", "outlook-bridge.ps1"));
	var startInfo = new ProcessStartInfo
	{
		FileName = "powershell",
		WorkingDirectory = repoRootPath,
		UseShellExecute = false,
		CreateNoWindow = true,
		RedirectStandardOutput = true,
		RedirectStandardError = true
	};

	startInfo.ArgumentList.Add("-NoProfile");
	startInfo.ArgumentList.Add("-ExecutionPolicy");
	startInfo.ArgumentList.Add("Bypass");
	startInfo.ArgumentList.Add("-File");
	startInfo.ArgumentList.Add(scriptPath);
	startInfo.ArgumentList.Add("-Action");
	startInfo.ArgumentList.Add(action);

	if (arguments is not null)
	{
		foreach (var pair in arguments)
		{
			startInfo.ArgumentList.Add("-" + pair.Key);
			startInfo.ArgumentList.Add(pair.Value);
		}
	}

	return await Maily.OutlookAddin.Web.Services.BridgeProcessRunner.RunAsync(
		startInfo, Program.OutlookBridgeLock, TimeSpan.FromSeconds(60), cancellationToken);
}

static string ReadJsonString(JsonElement element, string propertyName)
{
	return element.TryGetProperty(propertyName, out var property) && property.ValueKind == JsonValueKind.String
		? property.GetString() ?? string.Empty
		: string.Empty;
}

static bool ReadJsonBoolean(JsonElement element, string propertyName, bool defaultValue = false)
{
	return element.TryGetProperty(propertyName, out var property) && property.ValueKind is JsonValueKind.True or JsonValueKind.False
		? property.GetBoolean()
		: defaultValue;
}

static OutlookSearchCorpusEmail MapSearchCorpusEmail(JsonElement item, bool includeBodyText)
{
	var from = item.TryGetProperty("from", out var fromElement) && fromElement.TryGetProperty("emailAddress", out var emailAddress)
		? emailAddress
		: default;
	var body = item.TryGetProperty("body", out var bodyElement) ? bodyElement : default;
	var folderName = ReadJsonString(item, "folderName");
	var sentDateTime = ReadJsonString(item, "sentDateTime");
	var receivedDateTime = ReadJsonString(item, "receivedDateTime");
	var effectiveDateTime = folderName.Contains("sent", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(sentDateTime)
		? sentDateTime
		: receivedDateTime;

	return new OutlookSearchCorpusEmail
	{
		Id = ReadJsonString(item, "id"),
		Subject = ReadJsonString(item, "subject"),
		FromDisplayName = from.ValueKind == JsonValueKind.Object ? ReadJsonString(from, "name") : string.Empty,
		FromAddress = from.ValueKind == JsonValueKind.Object ? ReadJsonString(from, "address") : string.Empty,
		FolderName = folderName,
		ReceivedDateTime = effectiveDateTime,
		BodyPreview = ReadJsonString(item, "bodyPreview"),
		BodyText = includeBodyText && body.ValueKind == JsonValueKind.Object ? ReadJsonString(body, "text") : string.Empty,
		IsRead = ReadJsonBoolean(item, "isRead", defaultValue: true),
		Importance = ReadJsonString(item, "importance")
	};
}

static async Task<List<OutlookSearchCorpusEmail>> HydrateTopSearchCandidatesAsync(
	string hostRootPath,
	string mailboxAddress,
	IReadOnlyList<OutlookSearchCorpusEmail> previewEmails,
	OutlookSearchPreparationService searchPreparationService,
	string query,
	CancellationToken cancellationToken)
{
	var ranked = searchPreparationService.RankCandidates(query, previewEmails, Math.Min(24, previewEmails.Count));
	var hydratedIds = ranked
		.Take(8)
		.Select(email => email.Id)
		.Where(id => !string.IsNullOrWhiteSpace(id))
		.ToHashSet(StringComparer.OrdinalIgnoreCase);

	var enriched = new List<OutlookSearchCorpusEmail>(previewEmails.Count);
	foreach (var email in previewEmails)
	{
		if (!hydratedIds.Contains(email.Id))
		{
			enriched.Add(email);
			continue;
		}

		try
		{
			var payload = await InvokeOutlookBridgeAsync(
				hostRootPath,
				"get-email",
				new Dictionary<string, string>
				{
					["MailId"] = email.Id,
					["MailboxAddress"] = mailboxAddress
				},
				cancellationToken);

			using var json = JsonDocument.Parse(string.IsNullOrWhiteSpace(payload) ? "{}" : payload);
			var hydrated = MapSearchCorpusEmail(json.RootElement, includeBodyText: true);
			hydrated.SearchScore = email.SearchScore;
			hydrated.SearchSignals = email.SearchSignals;
			enriched.Add(hydrated);
		}
		catch
		{
			enriched.Add(email);
		}
	}

	return searchPreparationService.RankCandidates(query, enriched, 24).ToList();
}



static DateTimeOffset ParseSearchDate(string? value)
{
	return DateTimeOffset.TryParse(value, out var parsed)
		? parsed
		: DateTimeOffset.MinValue;
}

static OutlookBriefingMeetingItem MapBriefingMeeting(JsonElement item)
{
	var subject = ReadJsonString(item, "subject");
	var startDateTime = ReadJsonString(item, "startDateTime");
	var endDateTime = ReadJsonString(item, "endDateTime");
	var organizer = ReadJsonString(item, "organizer");
	var location = ReadJsonString(item, "location");
	var preview = ReadJsonString(item, "bodyPreview");

	return new OutlookBriefingMeetingItem
	{
		Subject = string.IsNullOrWhiteSpace(subject) ? "(No subject)" : subject,
		StartDateTime = startDateTime,
		EndDateTime = endDateTime,
		Organizer = organizer,
		Location = location,
		Preview = preview
	};
}

static OutlookBriefingMailItem MapBriefingMailItem(OutlookSearchCorpusEmail email, string whyItMatters)
{
	var fromLine = string.IsNullOrWhiteSpace(email.FromDisplayName)
		? email.FromAddress
		: $"{email.FromDisplayName} <{email.FromAddress}>";

	return new OutlookBriefingMailItem
	{
		Id = email.Id,
		Subject = string.IsNullOrWhiteSpace(email.Subject) ? "(No subject)" : email.Subject,
		FromLine = fromLine,
		ReceivedDateTime = email.ReceivedDateTime,
		FolderName = email.FolderName,
		BodyPreview = string.IsNullOrWhiteSpace(email.BodyText) ? email.BodyPreview : email.BodyText,
		WhyItMatters = whyItMatters
	};
}

static DateTime? ResolveSearchStartDateBound(string range) =>
	range.Trim().ToLowerInvariant() switch
	{
		"30d" => DateTime.UtcNow.AddDays(-30).Date,
		"90d" => DateTime.UtcNow.AddDays(-90).Date,
		"180d" => DateTime.UtcNow.AddDays(-180).Date,
		"365d" => DateTime.UtcNow.AddDays(-365).Date,
		"all" => DateTime.UtcNow.AddDays(-730).Date,
		_ => null
	};

static (int WindowDays, int WindowLimit, int MaxTotal) ResolveSearchFetchPlan(string scope, string dateRange)
{
	var normalizedScope = NormalizeSearchScope(scope);
	var normalizedRange = (dateRange ?? string.Empty).Trim().ToLowerInvariant();

	return (normalizedScope, normalizedRange) switch
	{
		("inbox", "30d") => (30, 250, 250),
		("inbox", "90d") => (30, 300, 900),
		("inbox", "180d") => (45, 350, 1600),
		("inbox", "365d") => (60, 400, 2600),
		("inbox", "all") => (60, 400, 3200),

		("sent-items", "30d") => (30, 250, 250),
		("sent-items", "90d") => (30, 300, 900),
		("sent-items", "180d") => (45, 350, 1600),
		("sent-items", "365d") => (60, 400, 2600),
		("sent-items", "all") => (60, 400, 3200),

		("all-mail", "30d") => (30, 250, 250),
		("all-mail", "90d") => (45, 300, 700),
		("all-mail", "180d") => (60, 300, 1000),
		("all-mail", "365d") => (60, 300, 1800),
		("all-mail", "all") => (90, 300, 2400),

		("all-mail", _) => (60, 300, 1000),
		_ => (45, 350, 1600)
	};
}

static async Task<SearchFetchWindowResult> FetchSearchCandidateWindowAsync(
	string hostRootPath,
	string mailboxAddress,
	string scope,
	DateTime? startDateBound,
	DateTime endDateBound,
	int limit,
	CancellationToken cancellationToken)
{
	const int sentItemsSafeWindowLimit = 31;
	const int sentItemsSubwindowDays = 21;

	if (scope == "sent-items" && startDateBound is not null && limit > sentItemsSafeWindowLimit)
	{
		var collected = new Dictionary<string, OutlookSearchCorpusEmail>(StringComparer.OrdinalIgnoreCase);
		var totalRawCount = 0;
		var isPartial = false;
		var subwindowEnd = endDateBound;

		while (subwindowEnd >= startDateBound.Value && collected.Count < limit)
		{
			var subwindowStart = subwindowEnd.AddDays(-(sentItemsSubwindowDays - 1));
			if (subwindowStart < startDateBound.Value)
			{
				subwindowStart = startDateBound.Value;
			}

			var subwindowResult = await FetchSearchCandidateWindowCoreAsync(
				hostRootPath,
				mailboxAddress,
				scope,
				subwindowStart,
				subwindowEnd,
				Math.Min(sentItemsSafeWindowLimit, limit - collected.Count),
				cancellationToken);

			totalRawCount += subwindowResult.Candidates.Count;
			isPartial |= subwindowResult.Diagnostic.IsPartial;
			foreach (var email in subwindowResult.Candidates)
			{
				if (!string.IsNullOrWhiteSpace(email.Id))
				{
					collected[email.Id] = email;
				}
			}

			subwindowEnd = subwindowStart.AddDays(-1);
		}

		var candidates = collected.Values
			.OrderByDescending(email => ParseSearchDate(email.ReceivedDateTime))
			.Take(limit)
			.ToList();

		return new SearchFetchWindowResult
		{
			Candidates = candidates,
			Diagnostic = new SearchFetchWindowDiagnostic
			{
				StartDateUtc = startDateBound.Value.ToString("yyyy-MM-dd"),
				EndDateUtc = endDateBound.ToString("yyyy-MM-dd"),
				IsPartial = isPartial || SearchCoverage.IsLimited(collected.Count, limit),
				RawCount = totalRawCount
			}
		};
	}

	return await FetchSearchCandidateWindowCoreAsync(
		hostRootPath,
		mailboxAddress,
		scope,
		startDateBound,
		endDateBound,
		limit,
		cancellationToken);
}

static async Task<SearchFetchWindowResult> FetchSearchCandidateWindowCoreAsync(
	string hostRootPath,
	string mailboxAddress,
	string scope,
	DateTime? startDateBound,
	DateTime endDateBound,
	int limit,
	CancellationToken cancellationToken)
{
	var arguments = new Dictionary<string, string>
	{
		["FolderName"] = scope,
		["MailboxAddress"] = mailboxAddress,
		["EndDate"] = endDateBound.ToString("yyyy-MM-dd"),
		["Limit"] = limit.ToString(System.Globalization.CultureInfo.InvariantCulture)
	};

	if (startDateBound is not null)
	{
		arguments["StartDate"] = startDateBound.Value.ToString("yyyy-MM-dd");
	}

	var payload = await InvokeOutlookBridgeAsync(
		hostRootPath,
		"list-search-emails",
		arguments,
		cancellationToken);

	using var json = JsonDocument.Parse(string.IsNullOrWhiteSpace(payload) ? "[]" : payload);
	var emails = new List<OutlookSearchCorpusEmail>();
	foreach (var item in json.RootElement.EnumerateArray())
	{
		emails.Add(MapSearchCorpusEmail(item, includeBodyText: false));
	}

	return new SearchFetchWindowResult
	{
		Candidates = emails,
		Diagnostic = new SearchFetchWindowDiagnostic
		{
			StartDateUtc = startDateBound?.ToString("yyyy-MM-dd") ?? string.Empty,
			EndDateUtc = endDateBound.ToString("yyyy-MM-dd"),
			IsPartial = SearchCoverage.IsLimited(emails.Count, limit),
			RawCount = emails.Count
		}
	};
}

static string NormalizeSearchScope(string? scope)
{
	var normalized = (scope ?? string.Empty).Trim().ToLowerInvariant();
	return normalized switch
	{
		"inbox" => "inbox",
		"sent-items" => "sent-items",
		_ => "all-mail"
	};
}

static string GetScopeLabel(string scope) =>
	scope switch
	{
		"inbox" => "Inbox",
		"sent-items" => "Sent Items",
		_ => "All mail"
	};

static async Task<SearchFetchResult> FetchSearchCandidatesAsync(
	string hostRootPath,
	string mailboxAddress,
	OutlookSearchFilters filters,
	OutlookSearchFetchCache fetchCache,
	CancellationToken cancellationToken)
{
	var scope = NormalizeSearchScope(filters.Scope);
	var normalizedMailboxAddress = string.IsNullOrWhiteSpace(mailboxAddress) ? "local-user" : mailboxAddress.Trim();
	var cached = fetchCache.TryGet(normalizedMailboxAddress, scope, filters.DateRange);
	if (cached is not null)
	{
		return new SearchFetchResult
		{
			Candidates = cached.Candidates,
			Diagnostics = new SearchFetchDiagnostics
			{
				Scope = scope,
				DateRange = filters.DateRange,
				StartDateUtc = ResolveSearchStartDateBound(filters.DateRange)?.ToString("yyyy-MM-dd") ?? string.Empty,
				EndDateUtc = DateTime.UtcNow.Date.ToString("yyyy-MM-dd"),
				WindowCount = 0,
				RawCandidateCount = cached.Candidates.Count,
				DistinctCandidateCount = cached.Candidates.Count,
				OldestCandidateUtc = cached.Candidates.OrderBy(email => ParseSearchDate(email.ReceivedDateTime)).FirstOrDefault()?.ReceivedDateTime ?? string.Empty,
				NewestCandidateUtc = cached.Candidates.OrderByDescending(email => ParseSearchDate(email.ReceivedDateTime)).FirstOrDefault()?.ReceivedDateTime ?? string.Empty,
				CacheHit = true,
				IsPartial = cached.IsPartial,
				CacheAgeSeconds = (int)Math.Round(cached.Age.TotalSeconds),
				CacheKey = cached.Key
			}
		};
	}

	var startDateBound = ResolveSearchStartDateBound(filters.DateRange);
	var (windowDays, windowLimit, maxTotal) = ResolveSearchFetchPlan(scope, filters.DateRange);
	var diagnostics = new SearchFetchDiagnostics
	{
		Scope = scope,
		DateRange = filters.DateRange,
		WindowDays = windowDays,
		WindowLimit = windowLimit,
		MaxTotal = maxTotal,
		StartDateUtc = startDateBound?.ToString("yyyy-MM-dd") ?? string.Empty,
		EndDateUtc = DateTime.UtcNow.Date.ToString("yyyy-MM-dd")
	};

	if (startDateBound is null)
	{
		var windowResult = await FetchSearchCandidateWindowAsync(
			hostRootPath,
			mailboxAddress,
			scope,
			null,
			DateTime.UtcNow.Date,
			maxTotal,
			cancellationToken);

		diagnostics.WindowCount = 1;
		diagnostics.IsPartial = windowResult.Diagnostic.IsPartial;
		diagnostics.RawCandidateCount = windowResult.Candidates.Count;
		diagnostics.DistinctCandidateCount = windowResult.Candidates.Count;
		diagnostics.WindowSummaries.Add(windowResult.Diagnostic);
		diagnostics.OldestCandidateUtc = windowResult.Candidates
			.OrderBy(email => ParseSearchDate(email.ReceivedDateTime))
			.FirstOrDefault()?.ReceivedDateTime ?? string.Empty;
		diagnostics.NewestCandidateUtc = windowResult.Candidates
			.OrderByDescending(email => ParseSearchDate(email.ReceivedDateTime))
			.FirstOrDefault()?.ReceivedDateTime ?? string.Empty;

		return new SearchFetchResult
		{
			Candidates = windowResult.Candidates,
			Diagnostics = diagnostics
		};
	}

	var collected = new Dictionary<string, OutlookSearchCorpusEmail>(StringComparer.OrdinalIgnoreCase);
	var windowEnd = DateTime.UtcNow.Date;

	while (windowEnd >= startDateBound.Value && collected.Count < maxTotal)
	{
		var windowStart = windowEnd.AddDays(-(windowDays - 1));
		if (windowStart < startDateBound.Value)
		{
			windowStart = startDateBound.Value;
		}

		var windowResult = await FetchSearchCandidateWindowAsync(
			hostRootPath,
			mailboxAddress,
			scope,
			windowStart,
			windowEnd,
			windowLimit,
			cancellationToken);
		var distinctBefore = collected.Count;

		foreach (var email in windowResult.Candidates)
		{
			if (!string.IsNullOrWhiteSpace(email.Id))
			{
				collected[email.Id] = email;
			}
		}

		windowResult.Diagnostic.DistinctAddedCount = collected.Count - distinctBefore;
		diagnostics.WindowSummaries.Add(windowResult.Diagnostic);
		diagnostics.WindowCount += 1;
		diagnostics.RawCandidateCount += windowResult.Candidates.Count;
		diagnostics.IsPartial |= windowResult.Diagnostic.IsPartial;

		windowEnd = windowStart.AddDays(-1);
	}

	var candidates = collected.Values
		.OrderByDescending(email =>
		{
			return ParseSearchDate(email.ReceivedDateTime);
		})
		.Take(maxTotal)
		.ToList();

	diagnostics.IsPartial |= SearchCoverage.IsLimited(collected.Count, maxTotal);
	diagnostics.DistinctCandidateCount = candidates.Count;
	diagnostics.OldestCandidateUtc = candidates
		.OrderBy(email => ParseSearchDate(email.ReceivedDateTime))
		.FirstOrDefault()?.ReceivedDateTime ?? string.Empty;
	diagnostics.NewestCandidateUtc = candidates
		.OrderByDescending(email => ParseSearchDate(email.ReceivedDateTime))
		.FirstOrDefault()?.ReceivedDateTime ?? string.Empty;

	fetchCache.Set(normalizedMailboxAddress, scope, filters.DateRange, candidates, diagnostics.IsPartial);

	return new SearchFetchResult
	{
		Candidates = candidates,
		Diagnostics = diagnostics
	};
}

static OutlookSearchSessionResponse BuildSearchSessionResponse(
	string requestId,
	string queryStage,
	OutlookSearchFilters filters,
	OutlookSearchResponse response,
	IReadOnlyList<OutlookSearchCorpusEmail> filteredCandidates,
	IReadOnlyList<OutlookSearchCorpusEmail> hydratedCandidates,
	OutlookSearchSessionState session,
	OutlookSearchPreparationService preparationService)
{
	return new OutlookSearchSessionResponse
	{
		Query = session.Query,
		Filters = new OutlookSearchFilters
		{
			Scope = filters.Scope,
			DateRange = filters.DateRange,
			Sender = filters.Sender,
			Focus = filters.Focus
		},
		AvailableRefinements = new OutlookSearchRefinements
		{
			Scope = OutlookSearchPreparationService.BuildStaticScopeOptions(),
			DateRange = OutlookSearchPreparationService.BuildStaticDateRangeOptions(),
			Focus = OutlookSearchPreparationService.BuildStaticFocusOptions(),
			Sender = preparationService.BuildSenderRefinements(filteredCandidates)
		},
		AnswerSummary = response.AnswerSummary,
		Commitments = response.Commitments,
		Results = response.Results,
		Meta = new OutlookSearchSessionMeta
		{
			IsPartial = session.IsPartial,
			Warning = SearchCoverage.MergeWarning(null, session.IsPartial),
			RequestId = requestId,
			SessionId = session.SessionId,
			Model = response.Model,
			QueryStage = queryStage,
			ScopeLabel = GetScopeLabel(filters.Scope),
			EvidenceMode = queryStage == "start" ? "preview-ranked" : "hydrated-top-results",
			CandidateCount = filteredCandidates.Count,
			ResultCount = response.Results.Count,
			HydratedCount = hydratedCandidates.Count(email => !string.IsNullOrWhiteSpace(email.BodyText)),
			LastUpdatedUtc = DateTimeOffset.UtcNow.ToString("o")
		}
	};
}

app.UseMailyBrowserBoundary(allowExternalHosts: !string.IsNullOrWhiteSpace(backendApiKey));
app.UseMailyApiKey(backendApiKey);

if (enableRequestLogging)
{
	app.Use(async (context, next) =>
	{
		try
		{
			await next();
			await AppendRequestLogAsync(logsRoot, requestLogLock, context, context.Response.StatusCode);
		}
		catch (Exception ex)
		{
			await AppendRequestLogAsync(logsRoot, requestLogLock, context, StatusCodes.Status500InternalServerError, ex);
			throw;
		}
	});
}

if (enableHttpLogging)
{
	app.UseHttpLogging();
}
app.UseDefaultFiles();
app.UseStaticFiles(new StaticFileOptions
{
	OnPrepareResponse = context =>
	{
		var requestPath = context.Context.Request.Path;
		if (requestPath.StartsWithSegments("/addin", StringComparison.OrdinalIgnoreCase) &&
			!requestPath.StartsWithSegments("/addin/assets", StringComparison.OrdinalIgnoreCase))
		{
			context.Context.Response.Headers.CacheControl = "no-store, no-cache, must-revalidate, max-age=0";
			context.Context.Response.Headers.Pragma = "no-cache";
			context.Context.Response.Headers.Expires = "0";
		}
	}
});

app.MapGet("/api/health", () => Results.Ok(new
{
	status = "ok",
	service = "Maily.OutlookAddin.Web",
	mode = isWindowsService ? "windows-service" : "interactive",
	timestampUtc = DateTimeOffset.UtcNow
}));

app.MapGet("/api/outlook/readiness", (
	PathResolver paths,
	CodexRuntimeConfigLoader codexConfigLoader) =>
{
	var checks = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
	var ready = true;

	void AddCheck(string name, bool ok, object details)
	{
		checks[name] = new
		{
			ok,
			details
		};
		ready &= ok;
	}

	AddCheck("repoRoot", Directory.Exists(repoRoot), new { path = repoRoot });
	AddCheck("manifest", File.Exists(Path.Combine(repoRoot, "addin", "manifest.xml")), new { path = Path.Combine(repoRoot, "addin", "manifest.xml") });
	AddCheck("prompts", Directory.Exists(paths.PromptsRoot), new { path = paths.PromptsRoot });
	var settingsDirectory = Path.GetDirectoryName(paths.SettingsFilePath) ?? string.Empty;
	if (!string.IsNullOrWhiteSpace(settingsDirectory))
	{
		Directory.CreateDirectory(settingsDirectory);
	}
	AddCheck("settingsStore", Directory.Exists(settingsDirectory), new { path = paths.SettingsFilePath });
	var outlookBridgePath = Path.Combine(repoRoot, "scripts", "outlook-bridge.ps1");
	var outlookBridgeExists = File.Exists(outlookBridgePath);
	if (hostingOptions.RequireOutlookBridge)
	{
		AddCheck("outlookBridge", outlookBridgeExists, new { path = outlookBridgePath, required = true });
	}
	else
	{
		checks["outlookBridge"] = new
		{
			ok = outlookBridgeExists,
			details = new { path = outlookBridgePath, required = false }
		};
	}

	var codexReadiness = CachedCodexReadiness.GetOrRefresh(
		paths,
		codexConfigLoader,
		TimeSpan.FromSeconds(Math.Clamp(hostingOptions.ReadinessCacheSeconds, 5, 300)),
		() => BuildCodexReadiness(paths, codexConfigLoader));
	AddCheck("codex", codexReadiness.Ok, codexReadiness.Details);

	return Results.Ok(new
	{
		status = ready ? "ready" : "degraded",
		ready,
		codexReady = codexReadiness.Ok,
		userMessage = ready
			? "Maily is ready."
			: codexReadiness.UserMessage,
		recommendedAction = ready
			? string.Empty
			: codexReadiness.RecommendedAction,
		service = "Maily.OutlookAddin.Web",
		mode = isWindowsService ? "windows-service" : "interactive",
		contentRoot = hostRoot,
		repoRoot,
		logsRoot,
		timestampUtc = DateTimeOffset.UtcNow,
		checks
	});
});

app.MapPost("/api/outlook/generate-reply", async (
	GenerateReplyRequest request,
	OutlookReplyGenerationService generationService,
	ILoggerFactory loggerFactory,
	CancellationToken cancellationToken) =>
{
	var logger = loggerFactory.CreateLogger("OutlookGenerateReply");
	var requestId = Guid.NewGuid().ToString("N");
	var timer = System.Diagnostics.Stopwatch.StartNew();

	try
	{
		logger.LogInformation(
			"generate_reply_started requestId={RequestId} mailbox={Mailbox} profile={ProfileId} mode={Mode}",
			requestId,
			"[redacted]",
			"[redacted]",
			request.Operation.Mode);

		var response = await generationService.GenerateReplyAsync(request, requestId, cancellationToken);
		timer.Stop();

		logger.LogInformation(
			"generate_reply_completed requestId={RequestId} model={Model} elapsedMs={ElapsedMs}",
			requestId,
			response.Meta.Model,
			timer.ElapsedMilliseconds);

		return Results.Ok(response);
	}
	catch (Exception)
	{
		timer.Stop();
		logger.LogError(
			"generate_reply_failed requestId={RequestId} mailbox={Mailbox} elapsedMs={ElapsedMs}",
			requestId,
			"[redacted]",
			timer.ElapsedMilliseconds);

		return Results.BadRequest(new
		{
			error = "The operation failed. Please retry or review your account settings.",
			requestId
		});
	}
});

app.MapPost("/api/outlook/rewrite-active-compose", async (
	OutlookComposeRewriteRequest request,
	ILoggerFactory loggerFactory,
	CancellationToken cancellationToken) =>
{
	var logger = loggerFactory.CreateLogger("OutlookComposeRewrite");
	var requestId = Guid.NewGuid().ToString("N");
	var timer = Stopwatch.StartNew();
	var tempHtmlPath = ComposeTempFile.CreatePath();

	try
	{
		var html = request.Html ?? string.Empty;
		if (!html.Contains("<!-- MAILY_REPLY_START -->", StringComparison.Ordinal))
		{
			html = $"<!-- MAILY_REPLY_START -->{html}<!-- MAILY_REPLY_END -->";
		}

		await File.WriteAllTextAsync(tempHtmlPath, html, cancellationToken);
		logger.LogInformation(
			"compose_rewrite_started requestId={RequestId} mailbox={Mailbox} subject={Subject} seedChars={SeedChars} composeType={ComposeType}",
			requestId,
			"[redacted]",
			"[redacted]",
			request.SeedText?.Length ?? 0,
			request.ComposeType);

		var bridgePayload = await InvokeOutlookBridgeAsync(
			repoRoot,
			"rewrite-active-compose",
			new Dictionary<string, string>
			{
				["Subject"] = request.Subject ?? string.Empty,
				["HtmlPath"] = tempHtmlPath,
				["DraftText"] = request.Text ?? string.Empty,
				["SeedText"] = request.SeedText ?? string.Empty,
				["UserName"] = request.UserName ?? string.Empty,
				["MailboxAddress"] = request.MailboxAddress ?? string.Empty,
				["ComposeType"] = request.ComposeType ?? "unknown"
			},
			cancellationToken);

		string bridgeStrategy = string.Empty;
		bool? bridgeSeedMatched = null;
		string bridgeDetectedDraft = string.Empty;
		int? bridgeReplaceEnd = null;

		try
		{
			using var bridgeJson = JsonDocument.Parse(string.IsNullOrWhiteSpace(bridgePayload) ? "{}" : bridgePayload);
			if (bridgeJson.RootElement.TryGetProperty("strategy", out var strategyElement))
			{
				bridgeStrategy = strategyElement.GetString() ?? string.Empty;
			}

			if (bridgeJson.RootElement.TryGetProperty("seedMatched", out var seedMatchedElement) &&
				(seedMatchedElement.ValueKind is JsonValueKind.True or JsonValueKind.False))
			{
				bridgeSeedMatched = seedMatchedElement.GetBoolean();
			}

			if (bridgeJson.RootElement.TryGetProperty("detectedDraftText", out var detectedDraftElement))
			{
				bridgeDetectedDraft = detectedDraftElement.GetString() ?? string.Empty;
			}

			if (bridgeJson.RootElement.TryGetProperty("replaceEnd", out var replaceEndElement) &&
				replaceEndElement.TryGetInt32(out var replaceEnd))
			{
				bridgeReplaceEnd = replaceEnd;
			}
		}
		catch
		{
		}

		timer.Stop();
		logger.LogInformation(
			"compose_rewrite_completed requestId={RequestId} elapsedMs={ElapsedMs} strategy={Strategy} seedMatched={SeedMatched} replaceEnd={ReplaceEnd} detectedDraftChars={DetectedDraftChars}",
			requestId,
			timer.ElapsedMilliseconds,
			bridgeStrategy,
			bridgeSeedMatched,
			bridgeReplaceEnd,
			bridgeDetectedDraft.Length);

		return Results.Content(string.IsNullOrWhiteSpace(bridgePayload) ? "{}" : bridgePayload, "application/json");
	}
	catch (Exception)
	{
		timer.Stop();
		logger.LogError(
			"compose_rewrite_failed requestId={RequestId} mailbox={Mailbox} subject={Subject} elapsedMs={ElapsedMs}",
			requestId,
			"[redacted]",
			"[redacted]",
			timer.ElapsedMilliseconds);

		return Results.BadRequest(new
		{
			error = "The operation failed. Please retry or review your account settings.",
			requestId
		});
	}
	finally
	{
		if (File.Exists(tempHtmlPath))
		{
			File.Delete(tempHtmlPath);
		}
	}
});

app.MapGet("/api/outlook/models", async (CodexModelCatalog catalog, CodexRuntimeConfigLoader configLoader, HttpContext context) =>
	Results.Ok(await catalog.GetAsync(configLoader.Load(), context.RequestAborted)));

app.MapGet("/api/outlook/settings", (
	string? mailboxAddress,
	string? profileId,
	ProfileSettingsStore store,
	OutlookSettingsCatalog catalog,
	ILoggerFactory loggerFactory) =>
{
	var logger = loggerFactory.CreateLogger("OutlookSettings");
	var resolvedMailbox = string.IsNullOrWhiteSpace(mailboxAddress) ? "local-user" : mailboxAddress.Trim();
	var resolvedProfileId = string.IsNullOrWhiteSpace(profileId) ? "default" : profileId.Trim();
	var response = store.Get(resolvedMailbox, resolvedProfileId);
	var catalogSnapshot = catalog.BuildSnapshot(response.Settings.Model);
	logger.LogInformation(
		"settings_loaded mailbox={Mailbox} profile={ProfileId}",
		"[redacted]",
		"[redacted]");
	return Results.Ok(OutlookSettingsSchema.BuildDto(response, catalogSnapshot));
});

app.MapPut("/api/outlook/settings", (
	OutlookSettingsSaveRequestDto request,
	string? mailboxAddress,
	ProfileSettingsStore store,
	OutlookSettingsCatalog catalog,
	ILoggerFactory loggerFactory) =>
{
	var logger = loggerFactory.CreateLogger("OutlookSettings");
	var resolvedMailbox = string.IsNullOrWhiteSpace(mailboxAddress) ? "local-user" : mailboxAddress.Trim();
	var resolvedProfileId = string.IsNullOrWhiteSpace(request.ProfileId) ? "default" : request.ProfileId.Trim();
	var existing = store.Get(resolvedMailbox, resolvedProfileId).Settings;
	var validation = OutlookSettingsSchema.ValidateAndMerge(request, existing);
	if (validation.Settings is null)
	{
		logger.LogWarning(
			"settings_validation_failed mailbox={Mailbox} profile={ProfileId} fields={FieldCount}",
			"[redacted]",
			"[redacted]",
			validation.Errors.Count);
		return Results.BadRequest(new
		{
			error = "Please review the highlighted settings.",
			fieldErrors = validation.Errors
		});
	}

	var response = store.Upsert(resolvedMailbox, resolvedProfileId, validation.Settings);
	var catalogSnapshot = catalog.BuildSnapshot(response.Settings.Model);
	logger.LogInformation(
		"settings_saved mailbox={Mailbox} profile={ProfileId}",
		"[redacted]",
		"[redacted]");
	return Results.Ok(OutlookSettingsSchema.BuildDto(response, catalogSnapshot));
});

app.MapPost("/api/outlook/style-profile/refresh", async (
	string? mailboxAddress,
	string? profileId,
	OutlookInsightsService insightsService,
	ProfileSettingsStore store,
	ILoggerFactory loggerFactory,
	CancellationToken cancellationToken) =>
{
	var logger = loggerFactory.CreateLogger("OutlookStyleProfile");
	var requestId = Guid.NewGuid().ToString("N");
	var resolvedMailbox = string.IsNullOrWhiteSpace(mailboxAddress) ? "local-user" : mailboxAddress.Trim();
	var resolvedProfileId = string.IsNullOrWhiteSpace(profileId) ? "default" : profileId.Trim();
	var timer = Stopwatch.StartNew();

	try
	{
		logger.LogInformation("style_refresh_started requestId={RequestId} mailbox={Mailbox} profile={ProfileId}", requestId, "[redacted]", "[redacted]");
		var payload = await InvokeOutlookBridgeAsync(
			repoRoot,
			"list-sent-emails",
			new Dictionary<string, string> { ["Limit"] = "12", ["BodyMode"] = "full", ["MailboxAddress"] = resolvedMailbox },
			cancellationToken);

		using var json = JsonDocument.Parse(string.IsNullOrWhiteSpace(payload) ? "[]" : payload);
		var samples = new List<StyleRefreshEmailSample>();
		foreach (var item in json.RootElement.EnumerateArray())
		{
			var body = item.TryGetProperty("body", out var bodyElement) ? ReadJsonString(bodyElement, "text") : string.Empty;
			if (string.IsNullOrWhiteSpace(body))
			{
				continue;
			}

			samples.Add(new StyleRefreshEmailSample
			{
				Subject = ReadJsonString(item, "subject"),
				BodyText = body,
				SentDateTime = ReadJsonString(item, "sentDateTime")
			});
		}

		var response = await insightsService.RefreshStyleProfileAsync(new OutlookStyleRefreshRequest
		{
			MailboxAddress = resolvedMailbox,
			ProfileId = resolvedProfileId,
			Samples = samples
		}, requestId, cancellationToken);

		timer.Stop();
		logger.LogInformation("style_refresh_completed requestId={RequestId} sampleCount={SampleCount} elapsedMs={ElapsedMs}", requestId, response.SampleCount, timer.ElapsedMilliseconds);
		return Results.Ok(response);
	}
	catch (Exception)
	{
		timer.Stop();
		var existing = store.Get(resolvedMailbox, resolvedProfileId).Settings;
		existing.StyleLastRefreshStatus = "error";
		existing.StyleLastRefreshError = "Style refresh failed. Please retry.";
		store.Upsert(resolvedMailbox, resolvedProfileId, existing);
		logger.LogError("style_refresh_failed requestId={RequestId} mailbox={Mailbox} elapsedMs={ElapsedMs}", requestId, "[redacted]", timer.ElapsedMilliseconds);
		return Results.BadRequest(new { error = "The operation failed. Please retry or review your account settings.", requestId });
	}
});

app.MapPost("/api/outlook/summary", async (
	OutlookThreadSummaryRequest request,
	OutlookInsightsService insightsService,
	ILoggerFactory loggerFactory,
	CancellationToken cancellationToken) =>
{
	var logger = loggerFactory.CreateLogger("OutlookSummary");
	var requestId = Guid.NewGuid().ToString("N");
	var timer = Stopwatch.StartNew();
	try
	{
		logger.LogInformation("summary_started requestId={RequestId} mailbox={Mailbox} surface={Surface}", requestId, "[redacted]", request.Surface);
		var response = await insightsService.SummarizeAsync(request, requestId, cancellationToken);
		timer.Stop();
		logger.LogInformation("summary_completed requestId={RequestId} model={Model} elapsedMs={ElapsedMs}", requestId, response.Meta.Model, timer.ElapsedMilliseconds);
		return Results.Ok(response);
	}
	catch (Exception)
	{
		timer.Stop();
		logger.LogError("summary_failed requestId={RequestId} mailbox={Mailbox} elapsedMs={ElapsedMs}", requestId, "[redacted]", timer.ElapsedMilliseconds);
		return Results.BadRequest(new { error = "The operation failed. Please retry or review your account settings.", requestId });
	}
});

app.MapPost("/api/outlook/daily-briefing", async (
	OutlookDailyBriefingRequest request,
	OutlookInsightsService insightsService,
	OutlookSearchPreparationService searchPreparationService,
	OutlookSearchFetchCache fetchCache,
	ILoggerFactory loggerFactory,
	CancellationToken cancellationToken) =>
{
	var logger = loggerFactory.CreateLogger("OutlookDailyBriefing");
	var requestId = Guid.NewGuid().ToString("N");
	var timer = Stopwatch.StartNew();

	try
	{
		if (string.IsNullOrWhiteSpace(request.MailboxAddress))
		{
			return Results.BadRequest(new { error = "mailboxAddress is required.", requestId });
		}

		logger.LogInformation("daily_briefing_started requestId={RequestId} mailbox={Mailbox}", requestId, "[redacted]");

		var meetingsPayload = await InvokeOutlookBridgeAsync(
			repoRoot,
			"list-calendar-events",
			new Dictionary<string, string>
			{
				["MailboxAddress"] = request.MailboxAddress,
				["StartDate"] = DateTime.Now.Date.ToString("yyyy-MM-dd"),
				["EndDate"] = DateTime.Now.Date.ToString("yyyy-MM-dd"),
				["Limit"] = "10"
			},
			cancellationToken);
		using var meetingsJson = JsonDocument.Parse(string.IsNullOrWhiteSpace(meetingsPayload) ? "[]" : meetingsPayload);
		var meetings = meetingsJson.RootElement
			.EnumerateArray()
			.Select(MapBriefingMeeting)
			.Where(meeting =>
			{
				var end = ParseSearchDate(meeting.EndDateTime);
				return end == DateTimeOffset.MinValue || end >= DateTimeOffset.Now.AddMinutes(-30);
			})
			.OrderBy(meeting => ParseSearchDate(meeting.StartDateTime))
			.Take(6)
			.ToList();

		var unreadFetch = await FetchSearchCandidatesAsync(
			repoRoot,
			request.MailboxAddress,
			new OutlookSearchFilters
			{
				Scope = "inbox",
				DateRange = "30d",
				Focus = "all",
				Sender = string.Empty
			},
			fetchCache,
			cancellationToken);
		var importantUnread = unreadFetch.Candidates
			.Where(email => !email.IsRead)
			.OrderByDescending(email => string.Equals(email.Importance, "high", StringComparison.OrdinalIgnoreCase))
			.ThenByDescending(email => ParseSearchDate(email.ReceivedDateTime))
			.Take(6)
			.Select(email => MapBriefingMailItem(
				email,
				string.Equals(email.Importance, "high", StringComparison.OrdinalIgnoreCase)
					? "Unread high-importance inbox email."
					: "Unread inbox email that may need attention."))
			.ToList();

		var actionFetch = await FetchSearchCandidatesAsync(
			repoRoot,
			request.MailboxAddress,
			new OutlookSearchFilters
			{
				Scope = "all-mail",
				DateRange = "30d",
				Focus = "all",
				Sender = string.Empty
			},
			fetchCache,
			cancellationToken);
		var rankedActionEmails = searchPreparationService.RankCandidates(
			"open action items follow up commitments next steps owners",
			actionFetch.Candidates,
			12);
		var actionCluster = searchPreparationService.SelectStrongResultCluster(rankedActionEmails, 6);
		var openActionItems = actionCluster
			.Select(email => MapBriefingMailItem(
				email,
				email.SearchSignals.Count == 0
					? "Likely follow-up or commitment evidence."
					: string.Join("; ", email.SearchSignals)))
			.ToList();

		var enrichedRequest = new OutlookDailyBriefingRequest
		{
			MailboxAddress = request.MailboxAddress,
			ProfileId = request.ProfileId,
			OutputLanguageOverride = request.OutputLanguageOverride,
			Meetings = meetings,
			OpenActionItems = openActionItems,
			ImportantUnreadEmails = importantUnread
		};
		var response = await insightsService.BuildDailyBriefingAsync(enrichedRequest, requestId, cancellationToken);

		timer.Stop();
		logger.LogInformation(
			"daily_briefing_completed requestId={RequestId} meetings={MeetingCount} openActions={OpenActionCount} unread={UnreadCount} model={Model} elapsedMs={ElapsedMs}",
			requestId,
			response.Meta.MeetingCount,
			response.Meta.OpenActionItemCount,
			response.Meta.ImportantUnreadCount,
			response.Meta.Model,
			timer.ElapsedMilliseconds);
		return Results.Ok(response);
	}
	catch (Exception)
	{
		timer.Stop();
		logger.LogError("daily_briefing_failed requestId={RequestId} mailbox={Mailbox} elapsedMs={ElapsedMs}", requestId, "[redacted]", timer.ElapsedMilliseconds);
		return Results.BadRequest(new { error = "The operation failed. Please retry or review your account settings.", requestId });
	}
});

app.MapPost("/api/outlook/search", async (
	OutlookSearchRequest request,
	OutlookInsightsService insightsService,
	OutlookSearchPreparationService searchPreparationService,
	ILoggerFactory loggerFactory,
	CancellationToken cancellationToken) =>
{
	var logger = loggerFactory.CreateLogger("OutlookSearch");
	var requestId = Guid.NewGuid().ToString("N");
	var timer = Stopwatch.StartNew();
	try
	{
		logger.LogInformation("search_started requestId={RequestId} mailbox={Mailbox} queryChars={QueryChars}", requestId, "[redacted]", request.Query?.Length ?? 0);
		request.Emails = await HydrateTopSearchCandidatesAsync(
			repoRoot,
			request.MailboxAddress,
			request.Emails,
			searchPreparationService,
			request.Query ?? string.Empty,
			cancellationToken);
		var response = await insightsService.SearchAsync(request, requestId, cancellationToken);
		timer.Stop();
		logger.LogInformation("search_completed requestId={RequestId} resultCount={ResultCount} elapsedMs={ElapsedMs}", requestId, response.Results.Count, timer.ElapsedMilliseconds);
		return Results.Ok(response);
	}
	catch (Exception)
	{
		timer.Stop();
		logger.LogError("search_failed requestId={RequestId} mailbox={Mailbox} elapsedMs={ElapsedMs}", requestId, "[redacted]", timer.ElapsedMilliseconds);
		return Results.BadRequest(new { error = "The operation failed. Please retry or review your account settings.", requestId });
	}
});

app.MapPost("/api/outlook/search/start", async (
	OutlookSearchStartRequest request,
	OutlookInsightsService insightsService,
	OutlookSearchPreparationService searchPreparationService,
	OutlookSearchSessionStore sessionStore,
	OutlookSearchFetchCache fetchCache,
	ILoggerFactory loggerFactory,
	CancellationToken cancellationToken) =>
{
	var logger = loggerFactory.CreateLogger("OutlookSearch");
	var requestId = Guid.NewGuid().ToString("N");
	var timer = Stopwatch.StartNew();

	try
	{
		logger.LogInformation(
			"search_start_started requestId={RequestId} mailbox={Mailbox} queryChars={QueryChars} scope={Scope} dateRange={DateRange}",
			requestId,
			"[redacted]",
			request.Query?.Length ?? 0,
			request.Filters.Scope,
			request.Filters.DateRange);

		var baseFetch = await FetchSearchCandidatesAsync(repoRoot, request.MailboxAddress, request.Filters, fetchCache, cancellationToken);
		var baseCandidates = baseFetch.Candidates;
		var session = sessionStore.CreateOrReplace(
			request.MailboxAddress,
			request.ProfileId,
			request.Query ?? string.Empty,
			request.Filters,
			baseCandidates, baseFetch.Diagnostics.IsPartial);
		var filteredCandidates = searchPreparationService.ApplyFilters(baseCandidates, request.Filters);
		var rankedCandidates = searchPreparationService.RankCandidates(request.Query ?? string.Empty, filteredCandidates, 24).ToList();
		var response = await insightsService.SearchAsync(new OutlookSearchRequest
		{
			MailboxAddress = request.MailboxAddress,
			ProfileId = request.ProfileId,
			Query = request.Query ?? string.Empty,
			Emails = rankedCandidates
		}, requestId, cancellationToken);
		sessionStore.UpdateCurrentFilters(session.SessionId, request.Filters);

		timer.Stop();
		logger.LogInformation(
			"search_fetch_coverage requestId={RequestId} stage=start scope={Scope} dateRange={DateRange} windows={WindowCount} rawCandidates={RawCandidateCount} distinctCandidates={DistinctCandidateCount} oldest={OldestCandidate} newest={NewestCandidate} plan={Plan} cacheHit={CacheHit} cacheAgeSeconds={CacheAgeSeconds}",
			requestId,
			baseFetch.Diagnostics.Scope,
			baseFetch.Diagnostics.DateRange,
			baseFetch.Diagnostics.WindowCount,
			baseFetch.Diagnostics.RawCandidateCount,
			baseFetch.Diagnostics.DistinctCandidateCount,
			baseFetch.Diagnostics.OldestCandidateUtc,
			baseFetch.Diagnostics.NewestCandidateUtc,
			$"{baseFetch.Diagnostics.WindowDays}/{baseFetch.Diagnostics.WindowLimit}/{baseFetch.Diagnostics.MaxTotal}",
			baseFetch.Diagnostics.CacheHit,
			baseFetch.Diagnostics.CacheAgeSeconds);
		logger.LogInformation(
			"search_ranked_preview requestId={RequestId} stage=start top={TopRanked} next={NextRanked}",
			requestId,
			"[redacted]",
			"[redacted]");
		logger.LogInformation(
			"search_start_completed requestId={RequestId} sessionId={SessionId} candidates={CandidateCount} elapsedMs={ElapsedMs}",
			requestId,
			session.SessionId,
			filteredCandidates.Count,
			timer.ElapsedMilliseconds);

		return Results.Ok(BuildSearchSessionResponse(
			requestId,
			"start",
			request.Filters,
			response,
			filteredCandidates,
			rankedCandidates,
			session,
			searchPreparationService));
	}
	catch (Exception)
	{
		timer.Stop();
		logger.LogError("search_start_failed requestId={RequestId} mailbox={Mailbox} elapsedMs={ElapsedMs}", requestId, "[redacted]", timer.ElapsedMilliseconds);
		return Results.BadRequest(new { error = "The operation failed. Please retry or review your account settings.", requestId });
	}
});

app.MapPost("/api/outlook/search/refine", async (
	OutlookSearchRefineRequest request,
	OutlookInsightsService insightsService,
	OutlookSearchPreparationService searchPreparationService,
	OutlookSearchSessionStore sessionStore,
	OutlookSearchFetchCache fetchCache,
	ILoggerFactory loggerFactory,
	CancellationToken cancellationToken) =>
{
	var logger = loggerFactory.CreateLogger("OutlookSearch");
	var requestId = Guid.NewGuid().ToString("N");
	var timer = Stopwatch.StartNew();

	try
	{
		var session = sessionStore.Get(request.SessionId);
		if (session is null)
		{
			return Results.NotFound(new { error = "Search session expired. Start a new search.", requestId });
		}

		var normalizedFilters = new OutlookSearchFilters
		{
			Scope = string.IsNullOrWhiteSpace(request.Filters.Scope) ? session.CurrentFilters.Scope : request.Filters.Scope,
			DateRange = string.IsNullOrWhiteSpace(request.Filters.DateRange) ? session.CurrentFilters.DateRange : request.Filters.DateRange,
			Sender = request.Filters.Sender ?? string.Empty,
			Focus = string.IsNullOrWhiteSpace(request.Filters.Focus) ? session.CurrentFilters.Focus : request.Filters.Focus
		};

		var baseScopeChanged =
			!string.Equals(session.BaseFilters.Scope, normalizedFilters.Scope, StringComparison.OrdinalIgnoreCase) ||
			!string.Equals(session.BaseFilters.DateRange, normalizedFilters.DateRange, StringComparison.OrdinalIgnoreCase);
		SearchFetchDiagnostics? fetchDiagnostics = null;
		if (baseScopeChanged)
		{
			var refreshedBase = await FetchSearchCandidatesAsync(repoRoot, session.MailboxAddress, normalizedFilters, fetchCache, cancellationToken);
			fetchDiagnostics = refreshedBase.Diagnostics;
			sessionStore.UpdateBaseCandidates(session.SessionId, normalizedFilters, refreshedBase.Candidates, refreshedBase.Diagnostics.IsPartial);
			session = sessionStore.Get(session.SessionId) ?? session;
		}

		var filteredCandidates = searchPreparationService.ApplyFilters(session.BaseCandidates, normalizedFilters);
		var hydratedCandidates = await HydrateTopSearchCandidatesAsync(
			repoRoot,
			session.MailboxAddress,
			searchPreparationService.RankCandidates(session.Query, filteredCandidates, 24).ToList(),
			searchPreparationService,
			session.Query,
			cancellationToken);
		var narrowedCandidates = searchPreparationService.ApplyFilters(hydratedCandidates, normalizedFilters);
		var rankedCandidates = searchPreparationService.RankCandidates(session.Query, narrowedCandidates, 24).ToList();
		var response = await insightsService.SearchAsync(new OutlookSearchRequest
		{
			MailboxAddress = session.MailboxAddress,
			ProfileId = session.ProfileId,
			Query = session.Query,
			Emails = rankedCandidates
		}, requestId, cancellationToken);
		sessionStore.UpdateCurrentFilters(session.SessionId, normalizedFilters);
		session = sessionStore.Get(session.SessionId) ?? session;

		timer.Stop();
		if (fetchDiagnostics is not null)
		{
			logger.LogInformation(
				"search_fetch_coverage requestId={RequestId} stage=refine scope={Scope} dateRange={DateRange} windows={WindowCount} rawCandidates={RawCandidateCount} distinctCandidates={DistinctCandidateCount} oldest={OldestCandidate} newest={NewestCandidate} plan={Plan} cacheHit={CacheHit} cacheAgeSeconds={CacheAgeSeconds}",
				requestId,
				fetchDiagnostics.Scope,
				fetchDiagnostics.DateRange,
				fetchDiagnostics.WindowCount,
				fetchDiagnostics.RawCandidateCount,
				fetchDiagnostics.DistinctCandidateCount,
				fetchDiagnostics.OldestCandidateUtc,
				fetchDiagnostics.NewestCandidateUtc,
				$"{fetchDiagnostics.WindowDays}/{fetchDiagnostics.WindowLimit}/{fetchDiagnostics.MaxTotal}",
				fetchDiagnostics.CacheHit,
				fetchDiagnostics.CacheAgeSeconds);
		}
		logger.LogInformation(
			"search_ranked_preview requestId={RequestId} stage=refine top={TopRanked} next={NextRanked}",
			requestId,
			"[redacted]",
			"[redacted]");
		logger.LogInformation(
			"search_refine_completed requestId={RequestId} sessionId={SessionId} candidates={CandidateCount} elapsedMs={ElapsedMs}",
			requestId,
			session.SessionId,
			narrowedCandidates.Count,
			timer.ElapsedMilliseconds);

		return Results.Ok(BuildSearchSessionResponse(
			requestId,
			"refine",
			normalizedFilters,
			response,
			narrowedCandidates,
			rankedCandidates,
			session,
			searchPreparationService));
	}
	catch (Exception)
	{
		timer.Stop();
		logger.LogError("search_refine_failed requestId={RequestId} sessionId={SessionId} elapsedMs={ElapsedMs}", requestId, "[redacted]", timer.ElapsedMilliseconds);
		return Results.BadRequest(new { error = "The operation failed. Please retry or review your account settings.", requestId });
	}
});

app.MapGet("/api/outlook/search/session/{sessionId}", (
	string sessionId,
	OutlookSearchSessionStore sessionStore) =>
{
	var session = sessionStore.Get(sessionId);
	return session is null
		? Results.NotFound(new { error = "Search session expired." })
		: Results.Ok(new
		{
			sessionId = session.SessionId,
			query = session.Query,
			filters = session.CurrentFilters,
			lastUpdatedUtc = session.LastUpdatedUtc.ToString("o")
		});
});

app.MapGet("/api/outlook/search-corpus", async (
	string mailboxAddress,
	CancellationToken cancellationToken) =>
{
	var payload = await InvokeOutlookBridgeAsync(
		repoRoot,
		"list-emails-preview",
		new Dictionary<string, string> { ["Limit"] = "80", ["MailboxAddress"] = mailboxAddress },
		cancellationToken);

	using var json = JsonDocument.Parse(string.IsNullOrWhiteSpace(payload) ? "[]" : payload);
	var emails = new List<OutlookSearchCorpusEmail>();
	foreach (var item in json.RootElement.EnumerateArray())
	{
		emails.Add(MapSearchCorpusEmail(item, includeBodyText: false));
	}

	return Results.Ok(new { mailboxAddress, emails });
});

app.MapPost("/api/outlook/client-log", (
	ClientLogEntry entry,
	ILoggerFactory loggerFactory) =>
{
	var logger = loggerFactory.CreateLogger("OutlookClient");
	// Client fields are untrusted; content and arbitrary labels never enter logs.
	logger.LogInformation("client_event received=true");

	return Results.Accepted();
});

app.MapGet("/", () => Results.Redirect("/addin/taskpane.html"));

var startupLogger = app.Services.GetRequiredService<ILoggerFactory>().CreateLogger("OutlookHost");
startupLogger.LogInformation(
	"host_started mode={Mode} contentRoot={ContentRoot} repoRoot={RepoRoot} requestLogPath={RequestLogPath} appLogPath={AppLogPath} deletedExpiredLogs={DeletedExpiredLogs}",
	isWindowsService ? "windows-service" : "interactive",
	hostRoot,
	repoRoot,
	string.IsNullOrWhiteSpace(startupRequestLogPath) ? "(disabled)" : startupRequestLogPath,
	addinLogPath,
	deletedExpiredLogs);

app.Run();

static string ResolveRepoRoot(string hostRoot, string? configuredRepoRoot)
{
	if (!string.IsNullOrWhiteSpace(configuredRepoRoot) && Directory.Exists(configuredRepoRoot))
	{
		return Path.GetFullPath(configuredRepoRoot);
	}

	var candidates = new[]
	{
		hostRoot,
		Directory.GetCurrentDirectory(),
		AppContext.BaseDirectory
	};

	foreach (var start in candidates.Where(path => !string.IsNullOrWhiteSpace(path)).Distinct(StringComparer.OrdinalIgnoreCase))
	{
		var current = new DirectoryInfo(start);
		while (current is not null)
		{
			if (File.Exists(Path.Combine(current.FullName, "maily.sln")) &&
				Directory.Exists(Path.Combine(current.FullName, "smoke_data", "prompts")) &&
				File.Exists(Path.Combine(current.FullName, "scripts", "outlook-bridge.ps1")))
			{
				return current.FullName;
			}

			current = current.Parent;
		}
	}

	throw new DirectoryNotFoundException("Could not resolve the Maily repository root. Set MailyGeneration:RepoRoot in appsettings.Production.json.");
}

static string ResolveHostRoot()
{
	var candidates = new[]
	{
		Directory.GetCurrentDirectory(),
		AppContext.BaseDirectory
	};

	foreach (var start in candidates.Where(path => !string.IsNullOrWhiteSpace(path)).Distinct(StringComparer.OrdinalIgnoreCase))
	{
		var current = new DirectoryInfo(start);
		while (current is not null)
		{
			var wwwrootPath = Path.Combine(current.FullName, "wwwroot");
			var projectPath = Path.Combine(current.FullName, "Maily.OutlookAddin.Web.csproj");
			var executablePath = Path.Combine(current.FullName, "Maily.OutlookAddin.Web.exe");
			var appSettingsPath = Path.Combine(current.FullName, "appsettings.json");
			if (Directory.Exists(wwwrootPath) &&
				(File.Exists(projectPath) || File.Exists(executablePath) || File.Exists(appSettingsPath)))
			{
				return current.FullName;
			}

			current = current.Parent;
		}
	}

	throw new DirectoryNotFoundException("Could not resolve the Maily.OutlookAddin.Web content root.");
}

static (bool Ok, object Details, string UserMessage, string RecommendedAction) BuildCodexReadiness(
	PathResolver paths,
	CodexRuntimeConfigLoader codexConfigLoader)
{
	try
	{
		var config = codexConfigLoader.Load();
		var version = RunCodexProcess(config, "--version");
		var loginStatus = version.Started && version.ExitCode == 0
			? RunCodexProcess(config, "login", "status")
			: new CodexProcessResult(false, -1, string.Empty, string.Empty, string.Empty);
		var status = CodexReadiness.Evaluate(version.Started, version.ExitCode,
			loginStatus.Started, loginStatus.ExitCode, loginStatus.Stdout, loginStatus.Stderr);
		return (status.Ok, status, status.Message, status.RecommendedAction);
	}
	catch (Exception)
	{
		var status = CodexReadiness.Evaluate(false, -1, false, -1, string.Empty, string.Empty);
		return (false, status, status.Message, status.RecommendedAction);
	}
}
static CodexProcessResult RunCodexProcess(CodexRuntimeConfig config, params string[] arguments)
{
	var executablePath = string.IsNullOrWhiteSpace(config.CodexExecutablePath) ? "codex" : config.CodexExecutablePath;
	var startInfo = new ProcessStartInfo
	{
		FileName = executablePath,
		RedirectStandardOutput = true,
		RedirectStandardError = true,
		StandardOutputEncoding = new System.Text.UTF8Encoding(false),
		StandardErrorEncoding = new System.Text.UTF8Encoding(false),
		UseShellExecute = false,
		CreateNoWindow = true
	};
	if (executablePath.EndsWith(".cmd", StringComparison.OrdinalIgnoreCase) ||
		executablePath.EndsWith(".bat", StringComparison.OrdinalIgnoreCase))
	{
		startInfo.FileName = Environment.GetEnvironmentVariable("COMSPEC") ?? "cmd.exe";
		startInfo.Arguments = WindowsBatchCommand.BuildArguments(new[] { executablePath }.Concat(arguments));
	}
	else
	{
		foreach (var argument in arguments)
		{
			startInfo.ArgumentList.Add(argument);
		}
	}

	if (!string.IsNullOrWhiteSpace(config.CodexHome))
	{
		startInfo.Environment["CODEX_HOME"] = config.CodexHome;
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


static int DeleteExpiredLogFiles(string logDirectory, TimeSpan retention)
{
	if (!Directory.Exists(logDirectory))
	{
		return 0;
	}

	var deleted = 0;
	var cutoffUtc = DateTime.UtcNow.Subtract(retention);
	foreach (var path in Directory.EnumerateFiles(logDirectory))
	{
		try
		{
			var lastWriteUtc = File.GetLastWriteTimeUtc(path);
			if (lastWriteUtc >= cutoffUtc)
			{
				continue;
			}

			File.Delete(path);
			deleted++;
		}
		catch
		{
			// Log cleanup must never block host startup.
		}
	}

	return deleted;
}

static string GetCurrentRequestLogPath(string logDirectory) =>
	Path.GetFullPath(Path.Combine(logDirectory, $"addin-web-requests-{DateTime.UtcNow:yyyyMMdd}.log"));

file sealed record CodexProcessResult(
	bool Started,
	int ExitCode,
	string Stdout,
	string Stderr,
	string Error);

file sealed class SearchFetchResult
{
	public List<OutlookSearchCorpusEmail> Candidates { get; set; } = [];
	public SearchFetchDiagnostics Diagnostics { get; set; } = new();
}

file sealed class SearchFetchDiagnostics
{
	public bool IsPartial { get; set; }
	public string Scope { get; set; } = string.Empty;
	public string DateRange { get; set; } = string.Empty;
	public int WindowDays { get; set; }
	public int WindowLimit { get; set; }
	public int MaxTotal { get; set; }
	public int WindowCount { get; set; }
	public int RawCandidateCount { get; set; }
	public int DistinctCandidateCount { get; set; }
	public string StartDateUtc { get; set; } = string.Empty;
	public string EndDateUtc { get; set; } = string.Empty;
	public string OldestCandidateUtc { get; set; } = string.Empty;
	public string NewestCandidateUtc { get; set; } = string.Empty;
	public bool CacheHit { get; set; }
	public int CacheAgeSeconds { get; set; }
	public string CacheKey { get; set; } = string.Empty;
	public List<SearchFetchWindowDiagnostic> WindowSummaries { get; set; } = [];
}

file sealed class SearchFetchWindowResult
{
	public List<OutlookSearchCorpusEmail> Candidates { get; set; } = [];
	public SearchFetchWindowDiagnostic Diagnostic { get; set; } = new();
}

file sealed class SearchFetchWindowDiagnostic
{
	public bool IsPartial { get; set; }
	public string StartDateUtc { get; set; } = string.Empty;
	public string EndDateUtc { get; set; } = string.Empty;
	public int RawCount { get; set; }
	public int DistinctAddedCount { get; set; }
}

file static class CachedCodexReadiness
{
	private static readonly object Gate = new();
	private static DateTimeOffset expiresAtUtc = DateTimeOffset.MinValue;
	private static string cacheKey = string.Empty;
	private static (bool Ok, object Details, string UserMessage, string RecommendedAction) cached;

	public static (bool Ok, object Details, string UserMessage, string RecommendedAction) GetOrRefresh(
		PathResolver paths,
		CodexRuntimeConfigLoader codexConfigLoader,
		TimeSpan ttl,
		Func<(bool Ok, object Details, string UserMessage, string RecommendedAction)> refresh)
	{
		var config = codexConfigLoader.Load();
		var key = string.Join("\n", config.CodexHome, config.CodexExecutablePath, config.Model);
		var now = DateTimeOffset.UtcNow;
		lock (Gate)
		{
			if (now < expiresAtUtc && string.Equals(cacheKey, key, StringComparison.Ordinal))
			{
				return cached;
			}
		}

		var refreshed = refresh();
		lock (Gate)
		{
			cacheKey = key;
			cached = refreshed;
			expiresAtUtc = DateTimeOffset.UtcNow.Add(ttl);
		}

		return refreshed;
	}
}

partial class Program
{
	internal static readonly SemaphoreSlim OutlookBridgeLock = new(1, 1);
}
