using System.Text.Json;
using Maily.Maui.Models;

namespace Maily.Maui.Services;

public sealed class UserSettingsService
{
	private readonly AppPaths appPaths;
	private readonly MailyLogger logger;
	private static readonly JsonSerializerOptions JsonOptions = new()
	{
		PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
		WriteIndented = true
	};

	public UserSettingsService(AppPaths appPaths, MailyLogger logger)
	{
		this.appPaths = appPaths;
		this.logger = logger;
	}

	public async Task<UserSettings> LoadAsync(CancellationToken cancellationToken = default)
	{
		if (!File.Exists(appPaths.SettingsFile))
		{
			return new UserSettings();
		}

		await using var stream = File.OpenRead(appPaths.SettingsFile);
		var settings = await JsonSerializer.DeserializeAsync<UserSettings>(stream, JsonOptions, cancellationToken);
		logger.Info("settings_loaded", new { path = appPaths.SettingsFile });
		return MergeWithDefaults(settings);
	}

	public async Task SaveAsync(UserSettings settings, CancellationToken cancellationToken = default)
	{
		Directory.CreateDirectory(Path.GetDirectoryName(appPaths.SettingsFile)!);
		await using var stream = File.Create(appPaths.SettingsFile);
		await JsonSerializer.SerializeAsync(stream, settings, JsonOptions, cancellationToken);
		logger.Info("settings_saved", new { path = appPaths.SettingsFile });
	}

	public string BuildStyleInstructions(UserSettings? settings)
	{
		if (settings is null)
		{
			return string.Empty;
		}

		var lines = new List<string>();
		if (!string.IsNullOrWhiteSpace(settings.Style.Tone))
		{
			lines.Add($"- Tone: {settings.Style.Tone}");
		}

		if (!string.IsNullOrWhiteSpace(settings.Style.LengthPreference))
		{
			lines.Add($"- Reply length: {settings.Style.LengthPreference}");
		}

		if (!string.IsNullOrWhiteSpace(settings.Style.Greeting))
		{
			lines.Add($"- Greeting style: {settings.Style.Greeting}");
		}

		if (!string.IsNullOrWhiteSpace(settings.Language))
		{
			lines.Add($"- Language: {settings.Language}");
		}

		if (!string.IsNullOrWhiteSpace(settings.Style.CustomInstructions))
		{
			lines.Add(string.Empty);
			lines.Add(settings.Style.CustomInstructions);
		}

		if (!string.IsNullOrWhiteSpace(settings.AnalyzedStyle))
		{
			lines.Add(string.Empty);
			lines.Add("Writing style inferred from sent emails:");
			lines.Add(settings.AnalyzedStyle);
		}

		return string.Join(Environment.NewLine, lines);
	}

	private static UserSettings MergeWithDefaults(UserSettings? raw)
	{
		var defaults = new UserSettings();
		if (raw is null)
		{
			return defaults;
		}

		defaults.MailboxAddress = raw.MailboxAddress?.Trim() ?? string.Empty;
		defaults.UserName = string.IsNullOrWhiteSpace(raw.UserName) ? defaults.UserName : raw.UserName;
		defaults.InboxLimit = raw.InboxLimit > 0 ? raw.InboxLimit : defaults.InboxLimit;
		defaults.InboxDatePreset = string.IsNullOrWhiteSpace(raw.InboxDatePreset) ? defaults.InboxDatePreset : raw.InboxDatePreset;
		defaults.InboxStartDate = raw.InboxStartDate;
		defaults.InboxEndDate = raw.InboxEndDate;
		defaults.Language = string.IsNullOrWhiteSpace(raw.Language) ? defaults.Language : raw.Language;
		defaults.AnalyzedStyle = raw.AnalyzedStyle;
		defaults.Style.Tone = string.IsNullOrWhiteSpace(raw.Style?.Tone) ? defaults.Style.Tone : raw.Style.Tone;
		defaults.Style.LengthPreference = string.IsNullOrWhiteSpace(raw.Style?.LengthPreference) ? defaults.Style.LengthPreference : raw.Style.LengthPreference;
		defaults.Style.Greeting = string.IsNullOrWhiteSpace(raw.Style?.Greeting) ? defaults.Style.Greeting : raw.Style.Greeting;
		defaults.Style.Closing = string.IsNullOrWhiteSpace(raw.Style?.Closing) ? defaults.Style.Closing : raw.Style.Closing;
		defaults.Style.CustomInstructions = raw.Style?.CustomInstructions ?? defaults.Style.CustomInstructions;
		return defaults;
	}
}
