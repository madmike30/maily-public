using Maily.Generation.Contracts;
using Maily.Generation.Models;
using Maily.Generation.Services;
using Maily.OutlookAddin.Web.Models;
using Microsoft.Extensions.Logging.Abstractions;
using System.Text.Json;

var failures = new List<string>();
var executed = 0;
var diagnosticSelfTest = args.SequenceEqual(new[] { "--diagnostics-self-test" });
string? filter = null;
if (args.Length == 2 && args[0] == "--filter" && !string.IsNullOrWhiteSpace(args[1]))
    filter = args[1];
else if (args.Length != 0 && !diagnosticSelfTest)
{
    Console.Error.WriteLine("Usage: [--filter <name fragment>] | [--diagnostics-self-test]");
    return 2;
}
if (diagnosticSelfTest)
{
    Run("synthetic diagnostic failure", () =>
        throw new InvalidOperationException("Synthetic outer failure", new Exception("Synthetic inner failure")));
    return Finish();
}

Run("release model defaults use Codex recommendation without invented entitlements", () =>
{
    var paths = new PathResolver(new GenerationOptions());
    using var config = JsonDocument.Parse(File.ReadAllText(Path.Combine(paths.RepoRoot, "services", "Maily.OutlookAddin.Web", "appsettings.json")));
    AssertEqual("", config.RootElement.GetProperty("MailyGeneration").GetProperty("Model").GetString()!);
    using var development = JsonDocument.Parse(File.ReadAllText(Path.Combine(paths.RepoRoot, "services", "Maily.OutlookAddin.Web", "appsettings.Development.json")));
    AssertEqual("", development.RootElement.GetProperty("MailyGeneration").GetProperty("Model").GetString()!);
    var catalog = new OutlookSettingsCatalog(paths);
    var defaults = catalog.BuildSnapshot(null);
    AssertEqual("Codex recommended default", defaults.DefaultModel);
    AssertTrue(defaults.ModelOptions.Count == 0, "No account catalog is available; do not invent model entitlements.");
    var custom = catalog.BuildSnapshot(" synthetic-custom-model ");
    AssertEqual("synthetic-custom-model", custom.ModelOptions.Single());
    var loader = new CodexRuntimeConfigLoader(paths, new GenerationOptions());
    AssertEqual("", loader.Load(" ").Model);
    AssertEqual("synthetic-custom-model", loader.Load(" synthetic-custom-model ").Model);
});
Run("sanitizer preserves substantive inline greetings and commitment details", () =>
{
    var settings = new AddinProfileSettings { UserName = "Maily Tester" };
    foreach (var text in new[] {
        "Hi Alex, I will send the report tomorrow.",
        "Hello, I will confirm by 2026-10-24 at 09:30.",
        "Dear Alex, please call +1 555 123 4567.",
        "I will send the report on 2026-10-24.",
        "Please email support@example.com for the report.",
        "Contact the project manager tomorrow",
        "Please call +1 555 123 4567 tomorrow" })
        AssertEqual(text, ReplySanitizer.NormalizeForOutlookReply(text, settings));
    AssertEqual("The report is ready.", ReplySanitizer.NormalizeForOutlookReply(
        "Hi Alex,\n\nThe report is ready.\n\nRegards,\nMaily Tester\nSupport Team\nmaily.tester@example.com", settings));
});
RunAsync("reply service rejects cleanup-only generated output", async () =>
{
    var root = Path.Combine(new PathResolver(new GenerationOptions()).RepoRoot, "output", "empty-draft", Guid.NewGuid().ToString("N"));
    Directory.CreateDirectory(root);
    var fakeCli = Path.Combine(root, OperatingSystem.IsWindows() ? "fake-codex.cmd" : "fake-codex");
    await File.WriteAllTextAsync(fakeCli, OperatingSystem.IsWindows()
        ? "@echo off\r\nmore > nul\r\necho Hi Alex,\r\nexit /b 0\r\n"
        : "#!/usr/bin/env sh\ncat > /dev/null\nprintf '%s\\n' 'Hi Alex,'\n");
    if (!OperatingSystem.IsWindows()) File.SetUnixFileMode(fakeCli, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
    var options = new GenerationOptions { CodexExecutablePath = fakeCli, CodexHome = Path.Combine(root, "fake-home"), SettingsFilePath = Path.Combine(root, "settings.json") };
    Directory.CreateDirectory(options.CodexHome);
    var paths = new PathResolver(options);
    var service = new OutlookReplyGenerationService(new CodexRuntimeConfigLoader(paths, options),
        new PromptTemplateEngine(paths), new OracleResponsesGateway(NullLogger<OracleResponsesGateway>.Instance), new ProfileSettingsStore(paths), options);
    var request = new GenerateReplyRequest { Client = new GenerateReplyClientContext { MailboxAddress = "synthetic@example.com" } };
    var rejected = false;
    try { await service.GenerateReplyAsync(request, "synthetic-empty-draft"); }
    catch (InvalidOperationException error) { AssertEqual("The generated reply contained no usable draft text after cleanup.", error.Message); rejected = true; }
    AssertTrue(rejected, "Cleanup-only provider output must not return a successful empty draft.");
});
Run("removes trailing signature block", () =>
{
	var settings = new AddinProfileSettings
	{
		UserName = "Maily Tester"
	};

	var result = ReplySanitizer.NormalizeForOutlookReply(
		"""
		Yes please do it.

		Regards,
		Maily Tester | Field CTO
		+1 555 123 4567
		maily.tester@example.com
		""",
		settings,
		"Yes please do it.");

	AssertEqual("Yes please do it.", result);
});

Run("keeps ordinary reply text", () =>
{
	var settings = new AddinProfileSettings
	{
		UserName = "Maily Tester"
	};

	var result = ReplySanitizer.NormalizeForOutlookReply(
		"""
		Thanks for the update.
		I'll review this today and get back to you.
		""",
		settings,
		"Thanks for the update.");

	AssertEqual("Thanks for the update.\nI'll review this today and get back to you.", result);
});

Run("search ranking prefers subject and hydrated body evidence", () =>
{
	var service = new OutlookSearchPreparationService();
	var ranked = service.RankCandidates(
		"pricing Example autonomous",
		[
			new OutlookSearchCorpusEmail
			{
				Id = "weak",
				Subject = "Synthetic weekly project notes",
				FromDisplayName = "Another Sender",
				FromAddress = "other@example.com",
				ReceivedDateTime = "2026-03-01T08:00:00Z",
				BodyPreview = "A generic operational update with no financial details."
			},
			new OutlookSearchCorpusEmail
			{
				Id = "strong",
				Subject = "Example Warehouse pricing",
				FromDisplayName = "Maily Tester",
				FromAddress = "maily.tester@example.com",
				ReceivedDateTime = "2026-03-20T09:00:00Z",
				BodyPreview = "The quote follows below.",
				BodyText = "Autonomous database pricing proposal: estimated monthly total 120 units."
			}
		],
		5);

	AssertEqual("strong", ranked[0].Id);
	AssertTrue(ranked[0].SearchSignals.Count > 0, "Expected ranking signals for the top result.");
});

Run("search ranking understands hebrew pricing intent for autonomous mail", () =>
{
	var service = new OutlookSearchPreparationService();
	var ranked = service.RankCandidates(
		"מחירים לאוטונומ",
		[
			new OutlookSearchCorpusEmail
			{
				Id = "recent-noise",
				Subject = "Synthetic project update",
				FromDisplayName = "Maily Tester",
				FromAddress = "maily.tester@example.com",
				ReceivedDateTime = "2026-03-30T10:00:00Z",
				BodyPreview = "A general project update without commercial details."
			},
			new OutlookSearchCorpusEmail
			{
				Id = "pricing-hit",
				Subject = "Example Warehouse estimate",
				FromDisplayName = "Maily Tester",
				FromAddress = "maily.tester@example.com",
				ReceivedDateTime = "2026-03-11T09:14:00Z",
				BodyPreview = "Synthetic autonomous database estimate with cloud credits.",
				BodyText = "Investment proposal: autonomous database monthly total 1234.5678 units. Quote is for investment proposal only."
			},
			new OutlookSearchCorpusEmail
			{
				Id = "partial-hit",
				Subject = "Example Warehouse questions",
				FromDisplayName = "Maily Tester",
				FromAddress = "maily.tester@example.com",
				ReceivedDateTime = "2026-03-01T08:00:00Z",
				BodyPreview = "A high level consumption estimate for an autonomous database."
			}
		],
		5);

	AssertEqual("pricing-hit", ranked[0].Id);
	AssertContains("pricing", string.Join(" | ", ranked[0].SearchSignals));
	AssertContains("autonomous", string.Join(" | ", ranked[0].SearchSignals));
});

Run("search ranking prefers exact rough cost estimate phrasing", () =>
{
	var service = new OutlookSearchPreparationService();
	var ranked = service.RankCandidates(
		"non binding rough cost estimate adb",
		[
			new OutlookSearchCorpusEmail
			{
				Id = "generic-estimate",
				Subject = "Synthetic planning meeting",
				FromDisplayName = "Maily Tester",
				FromAddress = "maily.tester@example.com",
				ReceivedDateTime = "2026-02-09T17:28:00Z",
				BodyPreview = "A high level consumption estimate for the autonomous database."
			},
			new OutlookSearchCorpusEmail
			{
				Id = "rough-cost-estimate",
				Subject = "Example Warehouse estimate",
				FromDisplayName = "Maily Tester",
				FromAddress = "maily.tester@example.com",
				ReceivedDateTime = "2026-01-29T16:16:00Z",
				BodyPreview = "The attached non binding rough cost estimate for adb outlines the monthly budget."
			}
		],
		5);

	AssertEqual("rough-cost-estimate", ranked[0].Id);
});

Run("search ranking demotes autonomous topic mail without commercial evidence", () =>
{
	var service = new OutlookSearchPreparationService();
	var ranked = service.RankCandidates(
		"autonomous pricing estimate",
		[
			new OutlookSearchCorpusEmail
			{
				Id = "topic-only",
				Subject = "Autonomous learning session",
				FromDisplayName = "Maily Tester",
				FromAddress = "maily.tester@example.com",
				ReceivedDateTime = "2026-03-28T09:00:00Z",
				BodyPreview = "Join a webinar about autonomous database architecture."
			},
			new OutlookSearchCorpusEmail
			{
				Id = "pricing-hit",
				Subject = "Example Warehouse commercial estimate",
				FromDisplayName = "Maily Tester",
				FromAddress = "maily.tester@example.com",
				ReceivedDateTime = "2026-01-29T16:16:00Z",
				BodyPreview = "Your autonomous database quote is ready.",
				BodyText = "Investment proposal. Monthly total 1200 units. Unit price 10. Quote is for investment proposal only."
			}
		],
		5);

	AssertEqual("pricing-hit", ranked[0].Id);
	AssertTrue(ranked[1].SearchScore < ranked[0].SearchScore, "Expected the topic-only autonomous email to score below the commercial evidence hit.");
});

Run("search ranking demotes newsletters for commercial autonomous queries", () =>
{
	var service = new OutlookSearchPreparationService();
	var ranked = service.RankCandidates(
		"autonomous pricing estimate",
		[
			new OutlookSearchCorpusEmail
			{
				Id = "newsletter",
				Subject = "Autonomous database weekly digest",
				FromDisplayName = "Example Vendor Newsletter",
				FromAddress = "synthetic1@example.invalid",
				ReceivedDateTime = "2026-03-30T08:00:00Z",
				BodyPreview = "New autonomous database capabilities and upcoming webinars. Unsubscribe here."
			},
			new OutlookSearchCorpusEmail
			{
				Id = "proposal",
				Subject = "Example Warehouse quote",
				FromDisplayName = "Maily Tester",
				FromAddress = "maily.tester@example.com",
				ReceivedDateTime = "2026-03-17T11:14:00Z",
				BodyPreview = "Please review the autonomous database pricing estimate.",
				BodyText = "Investment proposal: autonomous database monthly total 1200 units and unit price 10."
			}
		],
		5);

	AssertEqual("proposal", ranked[0].Id);
	AssertTrue(ranked[1].SearchScore < ranked[0].SearchScore, "Expected newsletter-like topic mail to score below the commercial evidence hit.");
});

Run("search result cluster trims low-confidence tail after strong evidence", () =>
{
	var service = new OutlookSearchPreparationService();
	var cluster = service.SelectStrongResultCluster(
		[
			new OutlookSearchCorpusEmail { Id = "one", SearchScore = 275, ReceivedDateTime = "2026-03-11T09:00:00Z" },
			new OutlookSearchCorpusEmail { Id = "two", SearchScore = 206, ReceivedDateTime = "2026-03-08T09:00:00Z" },
			new OutlookSearchCorpusEmail { Id = "three", SearchScore = 161, ReceivedDateTime = "2026-01-29T09:00:00Z" },
			new OutlookSearchCorpusEmail { Id = "tail-one", SearchScore = -20, ReceivedDateTime = "2025-09-09T09:00:00Z" },
			new OutlookSearchCorpusEmail { Id = "tail-two", SearchScore = -106, ReceivedDateTime = "2026-01-08T09:00:00Z" }
		],
		5);

	AssertTrue(cluster.Count == 3, $"Expected 3 clustered results, got {cluster.Count}.");
	AssertEqual("three", cluster[^1].Id);
});

Run("search ranking requires real commitment evidence for commitment queries", () =>
{
	var service = new OutlookSearchPreparationService();
	var ranked = service.RankCandidates(
		"what did we commit to",
		[
			new OutlookSearchCorpusEmail
			{
				Id = "newsletter-noise",
				Subject = "Synthetic project overview",
				FromDisplayName = "Events Team",
				FromAddress = "events@example.com",
				ReceivedDateTime = "2026-03-29T09:00:00Z",
				BodyPreview = "A general feature overview without agreed work."
			},
			new OutlookSearchCorpusEmail
			{
				Id = "real-commitment",
				Subject = "Synthetic project commitment",
				FromDisplayName = "Maily Tester",
				FromAddress = "maily.tester@example.com",
				ReceivedDateTime = "2026-03-28T09:00:00Z",
				BodyPreview = "We committed to send the revised plan by Friday."
			}
		],
		5);

	AssertEqual("real-commitment", ranked[0].Id);
});

Run("search ranking requires real decision evidence for decision queries", () =>
{
	var service = new OutlookSearchPreparationService();
	var ranked = service.RankCandidates(
		"what decisions were made",
		[
			new OutlookSearchCorpusEmail
			{
				Id = "topic-only",
				Subject = "Synthetic product announcement",
				FromDisplayName = "Example Vendor Newsletter",
				FromAddress = "synthetic1@example.invalid",
				ReceivedDateTime = "2026-03-30T09:00:00Z",
				BodyPreview = "Read our weekly product update and feature announcements."
			},
			new OutlookSearchCorpusEmail
			{
				Id = "real-decision",
				Subject = "Synthetic project decision",
				FromDisplayName = "Maily Tester",
				FromAddress = "maily.tester@example.com",
				ReceivedDateTime = "2026-03-27T09:00:00Z",
				BodyPreview = "We agreed and approved the revised project plan."
			}
		],
		5);

	AssertEqual("real-decision", ranked[0].Id);
});

Run("search focus filters recognize hebrew commitment and decision signals", () =>
{
	var service = new OutlookSearchPreparationService();
	var filteredCommitments = service.ApplyFilters(
		[
			new OutlookSearchCorpusEmail
			{
				Id = "he-commitment",
				BodyPreview = "אני אשלח מחר את העדכון ונעדכן על השלב הבא."
			},
			new OutlookSearchCorpusEmail
			{
				Id = "he-noise",
				BodyPreview = "תודה על ההזמנה לוובינר."
			}
		],
		new OutlookSearchFilters { Focus = "commitments", Scope = "all-mail", DateRange = "all", Sender = "" });

	var filteredDecisions = service.ApplyFilters(
		[
			new OutlookSearchCorpusEmail
			{
				Id = "he-decision",
				BodyPreview = "הוחלט להתקדם וסוכם שנאשר את הפיילוט."
			},
			new OutlookSearchCorpusEmail
			{
				Id = "he-noise-two",
				BodyPreview = "אשמח אם תעדכנו על הוובינר."
			}
		],
		new OutlookSearchFilters { Focus = "decisions", Scope = "all-mail", DateRange = "all", Sender = "" });

	AssertEqual("he-commitment", filteredCommitments[0].Id);
	AssertTrue(filteredCommitments.Count == 1, $"Expected exactly one Hebrew commitment match, got {filteredCommitments.Count}.");
	AssertEqual("he-decision", filteredDecisions[0].Id);
	AssertTrue(filteredDecisions.Count == 1, $"Expected exactly one Hebrew decision match, got {filteredDecisions.Count}.");
});

Run("search ranking demotes automated newsletter noise for broad ai queries", () =>
{
	var service = new OutlookSearchPreparationService();
	var ranked = service.RankCandidates(
		"ai",
		[
			new OutlookSearchCorpusEmail
			{
				Id = "newsletter",
				Subject = "AI platform update",
				FromDisplayName = "Example Vendor Newsletter",
				FromAddress = "synthetic2@example.invalid",
				ReceivedDateTime = "2026-03-30T09:00:00Z",
				BodyPreview = "Learn about our new AI platform features. View in browser or unsubscribe."
			},
			new OutlookSearchCorpusEmail
			{
				Id = "real-thread",
				Subject = "Synthetic AI workshop plan",
				FromDisplayName = "Maily Tester",
				FromAddress = "maily.tester@example.com",
				ReceivedDateTime = "2026-03-25T09:00:00Z",
				BodyPreview = "Let's review the AI design together and agree on next steps."
			}
		],
		5);

	AssertEqual("real-thread", ranked[0].Id);
});

Run("search ranking still allows exact digest subject searches", () =>
{
	var service = new OutlookSearchPreparationService();
	var ranked = service.RankCandidates(
		"weekly ai digest",
		[
			new OutlookSearchCorpusEmail
			{
				Id = "digest",
				Subject = "Weekly AI digest",
				FromDisplayName = "Postmaster",
				FromAddress = "synthetic3@example.invalid",
				ReceivedDateTime = "2026-03-29T09:00:00Z",
				BodyPreview = "AI digest with recent feature highlights. Unsubscribe here."
			},
			new OutlookSearchCorpusEmail
			{
				Id = "other",
				Subject = "Synthetic AI planning notes",
				FromDisplayName = "Maily Tester",
				FromAddress = "maily.tester@example.com",
				ReceivedDateTime = "2026-03-28T09:00:00Z",
				BodyPreview = "Let's discuss the AI plan for next week."
			}
		],
		5);

	AssertEqual("digest", ranked[0].Id);
});

Run("commitment summary extracts concrete commitment lines", () =>
{
	var service = new OutlookSearchPreparationService();
	var summary = service.BuildCommitmentsSummary(
		[
			new OutlookSearchCorpusEmail
			{
				Id = "one",
				Subject = "RE: OAS in Synthetic Project",
				FromDisplayName = "Maily Tester",
				FromAddress = "maily.tester@example.com",
				ReceivedDateTime = "2026-03-20T09:00:00Z",
				BodyText = "I will send the updated OAS architecture and pricing notes to Synthetic Project tomorrow morning. We also reviewed the open questions."
			}
		]);

	AssertContains("I will send the updated OAS architecture and pricing notes to Synthetic Project tomorrow morning.", summary);
	AssertContains("RE: OAS in Synthetic Project", summary);
});

Run("flexible string converter normalizes arrays and objects", () =>
{
	var options = new JsonSerializerOptions(JsonSerializerDefaults.Web);
	options.Converters.Add(new FlexibleStringJsonConverter());
	var payload = JsonSerializer.Deserialize<FlexiblePayload>(
		"""
		{
		  "mainPoints": ["Point one", "Point two"],
		  "actionItems": { "owner": "Alex", "task": "Reply today" }
		}
		""",
		options) ?? throw new InvalidOperationException("Payload was null.");

	AssertContains("Point one", payload.MainPoints);
	AssertContains("Point two", payload.MainPoints);
	AssertContains("Owner: Alex", payload.ActionItems);
	AssertContains("Task: Reply today", payload.ActionItems);
});

Run("flexible string converter normalizes style refresh payload variants", () =>
{
	var options = new JsonSerializerOptions(JsonSerializerDefaults.Web);
	options.Converters.Add(new FlexibleStringJsonConverter());
	var payload = JsonSerializer.Deserialize<StylePayload>(
		"""
		{
		  "styleSummary": { "summary": "Concise, pragmatic, bilingual." },
		  "analyzedStyle": [
		    "Short paragraphs",
		    { "content": "Lead with the recommendation" }
		  ]
		}
		""",
		options) ?? throw new InvalidOperationException("Payload was null.");

	AssertContains("Concise, pragmatic, bilingual.", payload.StyleSummary);
	AssertContains("Short paragraphs", payload.AnalyzedStyle);
	AssertContains("Lead with the recommendation", payload.AnalyzedStyle);
});

RunAsync("summary language failure returns retryable error instead of invented success", async () =>
{
    foreach (var output in new[] { "not-json", "{}", "{\"executiveSummary\":null}" })
    {
        var root = Path.Combine(new PathResolver(new GenerationOptions()).RepoRoot, "output", "invalid-summary", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        var fakeCli = Path.Combine(root, OperatingSystem.IsWindows() ? "fake-codex.cmd" : "fake-codex");
        await File.WriteAllTextAsync(fakeCli, OperatingSystem.IsWindows()
            ? "@echo off\r\nmore > nul\r\necho " + output + "\r\nexit /b 0\r\n"
            : "#!/usr/bin/env sh\ncat > /dev/null\nprintf '%s\\n' '" + output + "'\n");
        if (!OperatingSystem.IsWindows()) File.SetUnixFileMode(fakeCli, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        var options = new GenerationOptions { CodexExecutablePath = fakeCli, CodexHome = Path.Combine(root, "fake-home"), SettingsFilePath = Path.Combine(root, "settings.json") };
        Directory.CreateDirectory(options.CodexHome);
        var paths = new PathResolver(options);
        var service = new OutlookInsightsService(new CodexRuntimeConfigLoader(paths, options), new OracleResponsesGateway(NullLogger<OracleResponsesGateway>.Instance), new ProfileSettingsStore(paths), new OutlookSearchPreparationService());
        var rejected = false;
        try { await service.SummarizeAsync(new OutlookThreadSummaryRequest { MailboxAddress = "synthetic@example.invalid", BodyText = "Bonjour, merci pour votre message." }, "synthetic-invalid-summary"); }
        catch (InvalidOperationException error) { AssertEqual("The summary could not be completed. Please retry.", error.Message); rejected = true; }
        AssertTrue(rejected, "Invalid summary output must not return invented success.");
    }
});
Run("automatic language summary follows email instead of typed reply", () =>
{
    var systemMethod = typeof(OutlookInsightsService).GetMethod("BuildInsightsSystemPrompt", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
    var systemPrompt = (string)systemMethod.Invoke(null, ["thread-summary"])!;
    AssertContains("email_body", systemPrompt);
    AssertContains("draft_context must never determine", systemPrompt);
    var method = typeof(OutlookInsightsService).GetMethod("BuildSummaryPrompt", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
    var prompt = (string)method.Invoke(null, [new OutlookThreadSummaryRequest {
        BodyText = "שלום לכולם אנא אשרו את התוכנית המעודכנת", UserDraftText = "I will confirm tomorrow.", Subject = "Project plan"
    }, new AddinProfileSettings { Language = "English" }])!;
    AssertContains("appears to be Hebrew", prompt);
    AssertContains("Ignore the typed reply language", prompt);
    AssertContains("saved language settings and writing samples do not override it", prompt);
    AssertContains("<email_body>", prompt);
    AssertContains("<draft_context>", prompt);
});
Run("automatic language reply prioritizes typed text over incoming email", () =>
{
    var engine = new PromptTemplateEngine(new PathResolver(new GenerationOptions { RepoRoot = Directory.GetCurrentDirectory() }));
    foreach (var draft in new[] { "תודה אבדוק ואחזור מחר", "Bonjour, merci pour votre message." })
    {
        var prompt = engine.BuildSystemPrompt(new GenerateReplyRequest {
            Operation = new GenerateReplyOperation { CurrentDraft = draft },
            Compose = new GenerateReplyComposeContext { BodyText = "Please confirm the updated project schedule tomorrow." },
            UserPreferences = new GenerateReplyUserPreferences { Language = "English" }
        }, new AddinProfileSettings { Language = "English" });
        AssertContains("Saved language settings and writing samples do not override content language", prompt);
        AssertContains("Always answer in the same language as the user's typed reply draft", prompt);
        AssertDoesNotContain("appears to be English", prompt);
    }
});
Run("automatic language follows content majority without assuming Latin is English", () =>
{
    AssertTrue(ThreadLanguageHeuristics.DescribeDominantLanguage("Hello team, please review this detailed schedule with דוד before tomorrow.") != "Hebrew", "A Hebrew name must not override the English majority.");
    AssertTrue(ThreadLanguageHeuristics.DescribeDominantLanguage("Bonjour, merci pour votre message. Nous confirmerons demain.") != "English", "Latin script does not imply English.");
    AssertEqual("Hebrew", ThreadLanguageHeuristics.DescribeDominantLanguage("תודה על העדכון, אבדוק ואחזור מחר.\n-- \nEnglish Company International Customer Services"));
    AssertEqual("Hebrew", ThreadLanguageHeuristics.DescribeDominantLanguage("<style>long english styling font family color</style><p>תודה על העדכון אבדוק מחר</p>"));
    AssertTrue(ThreadLanguageHeuristics.DescribeDominantLanguage("I will review this tomorrow.\nFrom: synthetic@example.invalid\nשלום שלום שלום שלום שלום שלום") != "Hebrew", "Quoted text must not override the latest message.");
});
Run("language heuristic recognizes hebrew thread content", () =>
{
	var language = ThreadLanguageHeuristics.DescribeDominantLanguage(
		"שלום לכולם, תודה על העדכון",
		"אני אבדוק ואחזור אליכם מחר.");

	AssertEqual("Hebrew", language);
});

Run("language heuristic lets hebrew draft override long english thread", () =>
{
	var language = ThreadLanguageHeuristics.DescribeDominantLanguage(
		"תודה, אבדוק ואחזור אליכם מחר.",
		"""
		Hello team,
		Thanks for the detailed English update. This quoted thread is intentionally long so it should not
		outvote the user's current Hebrew draft. We discussed budgets, schedules, next steps, and ownership.
		Please review the plan and share feedback by tomorrow.
		""");

	AssertEqual("Hebrew", language);
});

Run("language heuristic ignores reply prefixes before hebrew subject", () =>
{
	var language = ThreadLanguageHeuristics.DescribeDominantLanguage("RE: שלום לגבי התוכנית");

	AssertEqual("Hebrew", language);
});

Run("language heuristic detects hebrew presence", () =>
{
	AssertTrue(ThreadLanguageHeuristics.ContainsHebrew("Hello", "שלום"), "Expected Hebrew detection to return true.");
	AssertTrue(!ThreadLanguageHeuristics.ContainsHebrew("Hello", "Thanks"), "Expected Hebrew detection to return false.");
});

Run("reply formatter emits rtl html for hebrew replies", () =>
{
	var html = ReplyFormatter.ToOutlookHtml("שלום לכולם");
	AssertContains("dir='rtl'", html);
	AssertContains("text-align:right", html);
});

Run("reply prompt honors explicit language override", () =>
{
	var engine = new PromptTemplateEngine(new PathResolver(new GenerationOptions
	{
		RepoRoot = Directory.GetCurrentDirectory()
	}));
	var prompt = engine.BuildSystemPrompt(
		new GenerateReplyRequest
		{
			Client = new GenerateReplyClientContext
			{
				MailboxAddress = "maily.tester@example.com",
				MailboxDisplayName = "Maily Tester"
			},
			Compose = new GenerateReplyComposeContext
			{
				Subject = "Re: Budget update",
				BodyText = "Hello team, thanks for the update."
			},
			Operation = new GenerateReplyOperation
			{
				OutputLanguageOverride = "Hebrew"
			}
		},
		new AddinProfileSettings
		{
			UserName = "Maily Tester"
		});

	AssertContains("The user explicitly selected Hebrew for this reply.", prompt);
	AssertContains("Write the final reply fully in Hebrew.", prompt);
});

Run("reply prompt preserves auto hebrew detection without explicit override", () =>
{
	var engine = new PromptTemplateEngine(new PathResolver(new GenerationOptions
	{
		RepoRoot = Directory.GetCurrentDirectory()
	}));
	var prompt = engine.BuildSystemPrompt(
		new GenerateReplyRequest
		{
			Client = new GenerateReplyClientContext
			{
				MailboxAddress = "maily.tester@example.com",
				MailboxDisplayName = "Maily Tester"
			},
			Compose = new GenerateReplyComposeContext
			{
				Subject = "RE: שלום לגבי התוכנית",
				BodyText = "שלום דור, אפשר להתקדם עם ההצעה כמו שסיכמנו?"
			},
			UserPreferences = new GenerateReplyUserPreferences
			{
				Language = "English"
			},
			Operation = new GenerateReplyOperation()
		},
		new AddinProfileSettings
		{
			UserName = "Maily Tester",
			Language = "English"
		});

	AssertContains("appears to be Hebrew", prompt);
	AssertContains("If there is no typed draft yet, answer in the same language as the latest email message or dominant thread language.", prompt);
	AssertDoesNotContain("The user explicitly selected English", prompt);
});

Run("reply identity respects explicitly selected outgoing shared sender", () =>
{
    var engine = new PromptTemplateEngine(new PathResolver(new GenerationOptions { RepoRoot = Directory.GetCurrentDirectory() }));
    var prompt = engine.BuildSystemPrompt(new() {
        Client = new() { MailboxAddress = "alex@example.invalid", MailboxDisplayName = "Alex Example" },
        Compose = new() { FromAddress = "support@example.invalid", FromDisplayName = "Example Support", FromIsComposeSender = true, BodyText = "From: Ada Other\nPlease help." }
    }, new() { UserName = "Old Name" });
    AssertContains("You are writing as: Example Support <support@example.invalid>", prompt);
    AssertDoesNotContain("You are writing as: Ada Other", prompt);
    AssertContains("Return only the reply body", prompt);
});
Run("reply identity survives stale settings and custom prompt without exchange", () =>
{
    var engine = new PromptTemplateEngine(new PathResolver(new GenerationOptions { RepoRoot = Directory.GetCurrentDirectory() }));
    var request = new GenerateReplyRequest {
        Client = new() { MailboxAddress = "alex@example.invalid", MailboxDisplayName = "Alex Example" },
        Compose = new() { FromAddress = "ada@example.invalid", FromDisplayName = "Ada Other", BodyText = "Please approve this. Regards, Ada Other" }
    };
    foreach (var template in new[] { "", "Help {{userName}} write a reply." }) {
        var prompt = engine.BuildSystemPrompt(request, new() { UserName = "Ada Other", SystemPromptOverride = template });
        AssertContains("You are writing as: Alex Example <alex@example.invalid>", prompt);
        AssertContains("Never write as the author of the incoming or quoted email", prompt);
        AssertContains("Outlook preserves the existing signature separately", prompt);
        AssertDoesNotContain("Help Ada Other write a reply", prompt);
    }
});
RunAsync("reply service uses mailbox identity to remove generated duplicate signature", async () =>
{
    var root = Path.Combine(Directory.GetCurrentDirectory(), "output", "identity", Guid.NewGuid().ToString("N"));
    Directory.CreateDirectory(root);
    var cli = Path.Combine(root, OperatingSystem.IsWindows() ? "fake.cmd" : "fake");
    await File.WriteAllTextAsync(cli, OperatingSystem.IsWindows()
        ? "@echo off\r\nmore > nul\r\necho I will confirm tomorrow.\r\necho Regards,\r\necho Alex Example\r\nexit /b 0\r\n"
        : "#!/bin/sh\ncat > /dev/null\nprintf '%s\\n' 'I will confirm tomorrow.' 'Regards,' 'Alex Example'\n");
    if (!OperatingSystem.IsWindows()) File.SetUnixFileMode(cli, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
    var options = new GenerationOptions { CodexExecutablePath = cli, CodexHome = Path.Combine(root, "fake-home"), SettingsFilePath = Path.Combine(root, "settings.json") };
    Directory.CreateDirectory(options.CodexHome);
    var paths = new PathResolver(options);
    var service = new OutlookReplyGenerationService(new CodexRuntimeConfigLoader(paths, options), new PromptTemplateEngine(paths), new OracleResponsesGateway(NullLogger<OracleResponsesGateway>.Instance), new ProfileSettingsStore(paths), options);
    var result = await service.GenerateReplyAsync(new() {
        Client = new() { MailboxAddress = "alex@example.invalid", MailboxDisplayName = "Alex Example" },
        Compose = new() { FromAddress = "ada@example.invalid", FromDisplayName = "Ada Other", BodyText = "Can you confirm?" }
    }, "synthetic-author");
    AssertEqual("I will confirm tomorrow.", result.Draft.Text);
});
Run("reply prompt writes as mailbox owner instead of incoming sender", () =>
{
	var engine = new PromptTemplateEngine(new PathResolver(new GenerationOptions
	{
		RepoRoot = Directory.GetCurrentDirectory()
	}));
	var prompt = engine.BuildSystemPrompt(
		new GenerateReplyRequest
		{
			Client = new GenerateReplyClientContext
			{
				MailboxAddress = "Alex@example.com",
				MailboxDisplayName = "Alex"
			},
			Compose = new GenerateReplyComposeContext
			{
				Subject = "Project update",
				FromDisplayName = "Ada",
				FromAddress = "ada@example.com",
				To = ["Alex <Alex@example.com>"],
				BodyText = "Hi Test Sender, please send the updated quote by Friday."
			}
		},
		new AddinProfileSettings
		{
			UserName = "Alex"
		});
	var userPrompt = engine.BuildUserPrompt(
		new GenerateReplyRequest
		{
			Client = new GenerateReplyClientContext
			{
				MailboxAddress = "Alex@example.com",
				MailboxDisplayName = "Alex"
			},
			Compose = new GenerateReplyComposeContext
			{
				Subject = "Project update",
				FromDisplayName = "Ada",
				FromAddress = "ada@example.com",
				BodyText = "Hi Test Sender, please send the updated quote by Friday."
			}
		},
		new AddinProfileSettings
		{
			UserName = "Alex"
		});

	AssertContains("You are writing as: Alex <Alex@example.com>", prompt);
	AssertContains("From: Ada <ada@example.com>", prompt);
	AssertDoesNotContain("You are writing as: Ada <ada@example.com>", prompt);
	AssertContains("Do not return only a greeting, closing, acknowledgement, or signature.", userPrompt);
});

Run("reply rewrite prompt honors explicit language override for seeded drafts", () =>
{
	var engine = new PromptTemplateEngine(new PathResolver(new GenerationOptions
	{
		RepoRoot = Directory.GetCurrentDirectory()
	}));
	var prompt = engine.BuildUserPrompt(
		new GenerateReplyRequest
		{
			Client = new GenerateReplyClientContext
			{
				MailboxAddress = "maily.tester@example.com",
				MailboxDisplayName = "Maily Tester"
			},
			Compose = new GenerateReplyComposeContext
			{
				Subject = "Re: Partnership outreach",
				BodyText = "Hello, thanks for reaching out."
			},
			Operation = new GenerateReplyOperation
			{
				CurrentDraft = "Thanks for reaching out, but this is not relevant for us at the moment.",
				OutputLanguageOverride = "Hebrew",
				Mode = "generate-initial"
			}
		},
		new AddinProfileSettings
		{
			UserName = "Maily Tester"
		});

	AssertContains("write the final reply fully in Hebrew", prompt);
	AssertContains("allowed to change the draft language because the user explicitly selected Hebrew", prompt);
	AssertContains("Rewrite the final reply fully in Hebrew", prompt);
});

Run("reply rewrite prompt keeps draft language when no override is selected", () =>
{
	var engine = new PromptTemplateEngine(new PathResolver(new GenerationOptions
	{
		RepoRoot = Directory.GetCurrentDirectory()
	}));
	var prompt = engine.BuildUserPrompt(
		new GenerateReplyRequest
		{
			Client = new GenerateReplyClientContext
			{
				MailboxAddress = "maily.tester@example.com",
				MailboxDisplayName = "Maily Tester"
			},
			Compose = new GenerateReplyComposeContext
			{
				Subject = "Re: Partnership outreach",
				BodyText = "Hello, thanks for reaching out."
			},
			Operation = new GenerateReplyOperation
			{
				CurrentDraft = "Thanks for reaching out, but this is not relevant for us at the moment.",
				Mode = "generate-initial"
			}
		},
		new AddinProfileSettings
		{
			UserName = "Maily Tester"
		});

	AssertContains("Preserve the user's intent, language, and concrete details", prompt);
	AssertContains("Keep the final reply in the same language as the current draft", prompt);
});

Run("summary prompt honors explicit language override", () =>
{
	var method = typeof(OutlookInsightsService).GetMethod("BuildSummaryPrompt", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
		?? throw new InvalidOperationException("BuildSummaryPrompt was not found.");
	var prompt = (string)(method.Invoke(null, [
		new OutlookThreadSummaryRequest
		{
			Surface = "compose",
			OutputLanguageOverride = "English",
			Subject = "שלום",
			BodyText = "שלום לכולם",
			UserDraftText = "תודה"
		},
		new AddinProfileSettings
		{
			UserName = "Maily Tester"
		}
	]) ?? throw new InvalidOperationException("Prompt was null."));

	AssertContains("The user explicitly selected English for this summary.", prompt);
	AssertContains("Write every field fully in English.", prompt);
});

Run("daily briefing prompt honors explicit language override", () =>
{
	var method = typeof(OutlookInsightsService).GetMethod("BuildDailyBriefingPrompt", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
		?? throw new InvalidOperationException("BuildDailyBriefingPrompt was not found.");
	var prompt = (string)(method.Invoke(null, [
		new OutlookDailyBriefingRequest
		{
			MailboxAddress = "maily.tester@example.com",
			OutputLanguageOverride = "Hebrew",
			Meetings =
			[
				new OutlookBriefingMeetingItem
				{
					Subject = "Budget sync",
					StartDateTime = "2026-04-02T09:00:00Z",
					EndDateTime = "2026-04-02T09:30:00Z"
				}
			]
		},
		new AddinProfileSettings
		{
			UserName = "Maily Tester",
			Language = "English"
		}
	]) ?? throw new InvalidOperationException("Prompt was null."));

	AssertContains("The user explicitly selected Hebrew for this briefing.", prompt);
	AssertContains("Write every field fully in Hebrew.", prompt);
});

RunAsync("codex gateway invokes codex exec with constrained stdin prompt", async () =>
{
	var tempRoot = Path.Combine(Path.GetTempPath(), "maily-codex-test-" + Guid.NewGuid().ToString("N"));
	Directory.CreateDirectory(tempRoot);
	var fakeCodexPath = Path.Combine(tempRoot, OperatingSystem.IsWindows() ? "fake-codex.cmd" : "fake-codex");
	var argsPath = Path.Combine(tempRoot, "args.txt");
	var stdinPath = Path.Combine(tempRoot, "stdin.txt");
	var homePath = Path.Combine(tempRoot, "home.txt");
	var codexHome = Path.Combine(tempRoot, "codex-home");
	Directory.CreateDirectory(codexHome);
	if (OperatingSystem.IsWindows())
	{
		await File.WriteAllTextAsync(
			fakeCodexPath,
			$"""
			@echo off
			echo %* > "{argsPath}"
			more > "{stdinPath}"
			echo %CODEX_HOME% > "{homePath}"
			echo Fake reply from Codex
			exit /b 0
			""");
	}
	else
	{
		await File.WriteAllTextAsync(
			fakeCodexPath,
			$"""
			#!/usr/bin/env sh
			printf '%s\n' "$*" > "{argsPath}"
			cat > "{stdinPath}"
			printf '%s\n' "$CODEX_HOME" > "{homePath}"
			printf '%s\n' "Fake reply from Codex"
			""");
		File.SetUnixFileMode(
			fakeCodexPath,
			UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
	}

	var gateway = new OracleResponsesGateway(NullLogger<OracleResponsesGateway>.Instance);
	var output = await gateway.GenerateReplyAsync(
		new CodexRuntimeConfig
		{
			CodexHome = codexHome,
			CodexExecutablePath = fakeCodexPath,
			Model = "test-model"
		},
		"system",
		"user",
		"codex-test");

	AssertEqual("Fake reply from Codex", output);
	var args = await File.ReadAllTextAsync(argsPath);
	AssertContains("exec", args);
	AssertContains("--ephemeral", args);
	AssertContains("--skip-git-repo-check", args);
	AssertContains("--sandbox read-only", args);
	AssertContains("--ignore-user-config", args);
	AssertContains("--ignore-rules", args);
	AssertContains("--disable plugins", args);
	AssertContains("--disable apps", args);
	AssertContains("--disable browser_use", args);
	AssertContains("--disable browser_use_external", args);
	AssertContains("--disable computer_use", args);
	AssertContains("--disable tool_search", args);
	AssertContains("--disable shell_snapshot", args);
	// cmd echo records shell quoting, not the native argv decoded by the target.
	// Maily.CliQuoting.Tests independently asserts the exact native argument values.
	AssertTrue(!args.Contains("model_reasoning_effort="), "Use the runtime reasoning default.");
	AssertContains("-m test-model", args);
	AssertContains("-", args);
	var stdin = await File.ReadAllTextAsync(stdinPath);
	AssertContains("Return only the final answer.", stdin);
	AssertContains("If JSON is requested, return one valid JSON object only.", stdin);
	AssertContains("System:", stdin);
	AssertContains("system", stdin);
	AssertContains("Task:", stdin);
	AssertContains("user", stdin);
	AssertEqual(codexHome, (await File.ReadAllTextAsync(homePath)).Trim());
});

Run("search filters narrow by sender and focus", () =>
{
	var service = new OutlookSearchPreparationService();
	var filtered = service.ApplyFilters(
		[
			new OutlookSearchCorpusEmail
			{
				Id = "1",
				FromDisplayName = "Example Reviewer",
				FromAddress = "synthetic4@example.invalid",
				FolderName = "Inbox",
				ReceivedDateTime = DateTimeOffset.UtcNow.ToString("o"),
				BodyPreview = "I will send the schedule update tomorrow."
			},
			new OutlookSearchCorpusEmail
			{
				Id = "2",
				FromDisplayName = "Another Sender",
				FromAddress = "other@example.com",
				FolderName = "Inbox",
				ReceivedDateTime = DateTimeOffset.UtcNow.ToString("o"),
				BodyPreview = "General status only."
			}
		],
		new OutlookSearchFilters
		{
			Sender = "synthetic4@example.invalid",
			Focus = "commitments",
			Scope = "all-mail",
			DateRange = "90d"
		});

	AssertEqual("1", filtered[0].Id);
	AssertTrue(filtered.Count == 1, "Expected sender + commitments filters to narrow to one email.");
});

Run("search sender refinements include ranked sender labels", () =>
{
	var service = new OutlookSearchPreparationService();
	var refinements = service.BuildSenderRefinements(
		[
			new OutlookSearchCorpusEmail
			{
				Id = "1",
				FromDisplayName = "Example Reviewer",
				FromAddress = "synthetic4@example.invalid"
			},
			new OutlookSearchCorpusEmail
			{
				Id = "2",
				FromDisplayName = "Example Reviewer",
				FromAddress = "synthetic4@example.invalid"
			}
		]);

	AssertEqual("synthetic4@example.invalid", refinements[0].Value);
	AssertContains("Example Reviewer", refinements[0].Label);
});

Run("search date filters accept legacy outlook json date strings", () =>
{
	var service = new OutlookSearchPreparationService();
	var filtered = service.ApplyFilters(
		[
			new OutlookSearchCorpusEmail
			{
				Id = "legacy-date",
				FromDisplayName = "Example Reviewer",
				FromAddress = "synthetic4@example.invalid",
				FolderName = "Inbox",
				ReceivedDateTime = $"/Date({DateTimeOffset.UtcNow.AddDays(-1).ToUnixTimeMilliseconds()})/",
				BodyPreview = "I will send the update tomorrow."
			}
		],
		new OutlookSearchFilters
		{
			Scope = "all-mail",
			DateRange = "180d",
			Focus = "all"
		});

	AssertTrue(filtered.Count == 1, "Expected legacy Outlook date values to survive date-range filtering.");
	AssertEqual("legacy-date", filtered[0].Id);
});

Run("settings schema validation rejects unsupported tone", () =>
{
	var validation = OutlookSettingsSchema.ValidateAndMerge(
		new OutlookSettingsSaveRequestDto
		{
			ProfileId = "default",
			Values = new OutlookSettingsValuesDto
			{
				UserName = "Maily Tester",
				Tone = "wild",
				Length = "concise",
				Language = "English",
				StylePersonalizationEnabled = true
			}
		},
		new AddinProfileSettings());

	AssertTrue(validation.Settings is null, "Expected invalid tone to fail validation.");
	AssertContains("supported options", validation.Errors["tone"]);
});

Run("settings schema preserves hidden fields on merge", () =>
{
	var validation = OutlookSettingsSchema.ValidateAndMerge(
		new OutlookSettingsSaveRequestDto
		{
			ProfileId = "default",
			Values = new OutlookSettingsValuesDto
			{
				UserName = "Maily Tester",
				Tone = "warm",
				Length = "concise",
				Language = "English",
				StylePersonalizationEnabled = true
			}
		},
		new AddinProfileSettings
		{
			Greeting = "Hello",
			Closing = "Best regards",
			StyleProfileSummary = "Existing style"
		});

	AssertEqual("Hello", validation.Settings?.Greeting ?? "");
	AssertEqual("Best regards", validation.Settings?.Closing ?? "");
	AssertEqual("Existing style", validation.Settings?.StyleProfileSummary ?? "");
});

Run("search fetch cache returns cloned candidates", () =>
{
	var cache = new OutlookSearchFetchCache();
	cache.Set(
		"maily.tester@example.com",
		"all-mail",
		"365d",
		[
			new OutlookSearchCorpusEmail
			{
				Id = "cached-1",
				Subject = "Original subject",
				SearchSignals = ["recent"]
			}
		]);

	var firstHit = cache.TryGet("maily.tester@example.com", "all-mail", "365d")
		?? throw new InvalidOperationException("Expected first cache hit.");
	firstHit.Candidates[0].Subject = "Mutated subject";
	firstHit.Candidates[0].SearchSignals.Add("mutated");

	var secondHit = cache.TryGet("maily.tester@example.com", "all-mail", "365d")
		?? throw new InvalidOperationException("Expected second cache hit.");

	AssertEqual("Original subject", secondHit.Candidates[0].Subject);
	AssertTrue(secondHit.Candidates[0].SearchSignals.Count == 1, "Expected cache to return isolated candidate clones.");
});

return Finish();

int Finish()
{
    foreach (var failure in failures) Console.Error.WriteLine(failure);
    Console.WriteLine($"RESULT: {executed - failures.Count} passed, {failures.Count} failed");
    if (executed == 0)
    {
        Console.Error.WriteLine("No tests matched the requested filter.");
        return 1;
    }
    return failures.Count == 0 ? 0 : 1;
}
void Run(string name, Action action)
{
	if (filter is not null && !name.Contains(filter, StringComparison.OrdinalIgnoreCase)) return;
	executed++;
	try
	{
		action();
		Console.WriteLine($"PASS {name}");
	}
	catch (Exception ex)
	{
		failures.Add($"FAIL {name}: {ex}");
	}
}

void RunAsync(string name, Func<Task> action)
{
	if (filter is not null && !name.Contains(filter, StringComparison.OrdinalIgnoreCase)) return;
	executed++;
	try
	{
		action().GetAwaiter().GetResult();
		Console.WriteLine($"PASS {name}");
	}
	catch (Exception ex)
	{
		failures.Add($"FAIL {name}: {ex}");
	}
}

static void AssertEqual(string expected, string actual)
{
	if (!string.Equals(expected, actual, StringComparison.Ordinal))
	{
		throw new InvalidOperationException($"Expected '{expected}' but got '{actual}'.");
	}
}

static void AssertContains(string expectedFragment, string actual)
{
	if ((actual ?? string.Empty).IndexOf(expectedFragment, StringComparison.Ordinal) < 0)
	{
		throw new InvalidOperationException($"Expected '{actual}' to contain '{expectedFragment}'.");
	}
}

static void AssertDoesNotContain(string unexpectedFragment, string actual)
{
	if ((actual ?? string.Empty).IndexOf(unexpectedFragment, StringComparison.Ordinal) >= 0)
	{
		throw new InvalidOperationException($"Expected '{actual}' not to contain '{unexpectedFragment}'.");
	}
}

static void AssertTrue(bool condition, string message)
{
	if (!condition)
	{
		throw new InvalidOperationException(message);
	}
}

file sealed class FlexiblePayload
{
	public string MainPoints { get; set; } = string.Empty;
	public string ActionItems { get; set; } = string.Empty;
}

file sealed class StylePayload
{
	public string StyleSummary { get; set; } = string.Empty;
	public string AnalyzedStyle { get; set; } = string.Empty;
}
