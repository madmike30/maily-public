using Maily.Maui.Models;
using Maily.Maui.Services;

var createDraft = args.Contains("--create-draft", StringComparer.OrdinalIgnoreCase);
var logger = new MailyLogger(new AppPaths());
var settingsService = new UserSettingsService(new AppPaths(), logger);
var codexConfigService = new CodexConfigService(logger);
var promptTemplateService = new PromptTemplateService(new AppPaths(), settingsService);
var outlookBridgeService = new OutlookBridgeService(new AppPaths(), logger);
var oracleClient = new OracleResponsesClient(new HttpClient(), logger);
var draftService = new DraftService(codexConfigService, promptTemplateService, oracleClient, logger);

Console.WriteLine($"Smoke log: {logger.LogPath}");

var originalSettingsPath = new AppPaths().SettingsFile;
var originalSettingsExists = File.Exists(originalSettingsPath);
var originalSettingsContent = originalSettingsExists ? await File.ReadAllTextAsync(originalSettingsPath) : null;

try
{
	Console.WriteLine("1. Loading settings");
	var settings = await settingsService.LoadAsync();
	Console.WriteLine($"   Loaded profile '{settings.UserName}' with inbox limit {settings.InboxLimit}.");

	Console.WriteLine("2. Saving and restoring a probe settings profile");
	var probeName = $"Smoke Probe {DateTime.Now:HHmmss}";
	settings.UserName = probeName;
	await settingsService.SaveAsync(settings);
	var reloadedSettings = await settingsService.LoadAsync();
	if (!string.Equals(reloadedSettings.UserName, probeName, StringComparison.Ordinal))
	{
		throw new InvalidOperationException("Settings probe save/load mismatch.");
	}
	Console.WriteLine($"   Settings round-trip OK for '{probeName}'.");

	Console.WriteLine("3. Loading Codex config");
	var config = codexConfigService.Load();
	Console.WriteLine($"   Model provider ready with default model '{config.Model}'.");

	Console.WriteLine("4. Testing Outlook connection");
	await outlookBridgeService.TestConnectionAsync();
	Console.WriteLine("   Outlook bridge responded.");

	Console.WriteLine("5. Listing inbox preview");
	var inbox = await outlookBridgeService.ListInboxAsync(Math.Min(reloadedSettings.InboxLimit, 10));
	if (inbox.Count == 0)
	{
		throw new InvalidOperationException("Inbox preview returned zero emails.");
	}
	var selected = inbox[0];
	Console.WriteLine($"   Loaded {inbox.Count} emails. First subject: {selected.Subject}");

	Console.WriteLine("6. Loading selected email details");
	selected = await outlookBridgeService.GetEmailAsync(selected.Id);
	Console.WriteLine($"   Detailed body length: {selected.BodyText.Length}. Attachments: {selected.Attachments.Count}.");

	Console.WriteLine("7. Generating draft");
	var effectiveSettings = new UserSettings
	{
		UserName = reloadedSettings.UserName,
		InboxLimit = reloadedSettings.InboxLimit,
		Language = reloadedSettings.Language,
		AnalyzedStyle = reloadedSettings.AnalyzedStyle,
		Style = new UserStyleSettings
		{
			Tone = reloadedSettings.Style.Tone,
			LengthPreference = reloadedSettings.Style.LengthPreference,
			Greeting = reloadedSettings.Style.Greeting,
			Closing = reloadedSettings.Style.Closing,
			CustomInstructions = reloadedSettings.Style.CustomInstructions
		}
	};
	var draft = await draftService.GenerateInitialDraftAsync(selected, effectiveSettings, config.Model);
	if (string.IsNullOrWhiteSpace(draft))
	{
		throw new InvalidOperationException("Draft generation returned empty text.");
	}
	Console.WriteLine($"   Draft generated ({draft.Length} chars).");

	if (createDraft)
	{
		Console.WriteLine("8. Creating Outlook reply draft");
		var html = draftService.ConvertDraftToOutlookHtml(draft);
		await outlookBridgeService.CreateReplyDraftAsync(selected.Id, html, replyAll: false);
		Console.WriteLine("   Outlook draft opened successfully.");
	}

	Console.WriteLine("Smoke runner completed successfully.");
}
finally
{
	if (originalSettingsExists && originalSettingsContent is not null)
	{
		await File.WriteAllTextAsync(originalSettingsPath, originalSettingsContent);
	}
	else if (File.Exists(originalSettingsPath))
	{
		File.Delete(originalSettingsPath);
	}
}
