using Maily.Generation.Contracts;
using Maily.Generation.Models;

namespace Maily.OutlookAddin.Web.Models;

public static class OutlookSettingsSchema
{
	public static readonly string[] ToneOptions = ["professional", "concise", "warm", "executive summary"];
	public static readonly string[] LengthOptions = ["concise", "standard", "detailed"];
	public static readonly string[] LanguageOptions = ["English", "Hebrew", "Same as thread"];
	private const int UserNameMaxLength = 120;
	private const int ModelMaxLength = 80;
	private const int CustomInstructionsMaxLength = 1200;
	private const int PromptOverrideMaxLength = 4000;

	public static OutlookSettingsDto BuildDto(
		AddinSettingsResponse response,
		OutlookSettingsCatalogSnapshot? catalog = null,
		Dictionary<string, string>? fieldErrors = null)
	{
		var settings = response.Settings;
		var resolvedCatalog = catalog ?? new OutlookSettingsCatalogSnapshot();
		var defaults = BuildDefaults();
		var style = BuildStyleState(settings);
		var values = new OutlookSettingsValuesDto
		{
			UserName = settings.UserName,
			Tone = settings.Tone,
			Length = settings.LengthPreference,
			Language = settings.Language,
			Model = settings.Model ?? string.Empty,
			CustomInstructions = settings.CustomInstructions,
			SystemPromptOverride = settings.SystemPromptOverride,
			DraftPromptOverride = settings.DraftPromptOverride,
			StylePersonalizationEnabled = settings.StylePersonalizationEnabled
		};

		return new OutlookSettingsDto
		{
			ProfileId = response.ProfileId,
			UpdatedUtc = response.UpdatedUtc.ToString("o"),
			Values = values,
			Defaults = defaults,
			Options = new OutlookSettingsOptionsDto
			{
				Tone = ToneOptions.ToList(),
				Length = LengthOptions.ToList(),
				Language = LanguageOptions.ToList(),
				Model = resolvedCatalog.ModelOptions.ToList()
			},
			PromptCatalog = resolvedCatalog.PromptCatalog,
			FieldMetadata = BuildFieldMetadata(defaults, resolvedCatalog),
			FieldErrors = fieldErrors ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
			AdvancedWarnings =
			[
				"Prompt overrides affect all generated replies.",
				"Leave these blank unless you are intentionally changing Maily's baseline."
			],
			Style = style,
			UserName = values.UserName,
			Tone = values.Tone,
			Length = values.Length,
			Language = values.Language,
			Model = values.Model,
			CustomInstructions = values.CustomInstructions,
			SystemPromptOverride = values.SystemPromptOverride,
			DraftPromptOverride = values.DraftPromptOverride,
			StylePersonalizationEnabled = values.StylePersonalizationEnabled,
			StyleProfileSummary = style.Summary,
			AnalyzedStyle = style.AnalyzedStyle,
			StyleSampleCount = style.SampleCount,
			StyleLastRefreshedUtc = style.LastRefreshedUtc,
			StyleLastRefreshStatus = style.LastRefreshStatus,
			StyleLastRefreshError = style.LastRefreshError
		};
	}

	public static (AddinProfileSettings? Settings, Dictionary<string, string> Errors) ValidateAndMerge(
		OutlookSettingsSaveRequestDto request,
		AddinProfileSettings existing)
	{
		var errors = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
		var values = request.Values ?? new OutlookSettingsValuesDto();

		var userName = NormalizeText(values.UserName, existing.UserName, UserNameMaxLength, "userName", errors, allowBlank: false);
		var tone = NormalizeEnum(values.Tone, ToneOptions, existing.Tone, "tone", errors);
		var length = NormalizeEnum(values.Length, LengthOptions, existing.LengthPreference, "length", errors);
		var language = NormalizeEnum(values.Language, LanguageOptions, existing.Language, "language", errors);
		var model = NormalizeOptional(values.Model, ModelMaxLength, "model", errors);
		var customInstructions = NormalizeOptional(values.CustomInstructions, CustomInstructionsMaxLength, "customInstructions", errors);
		var systemPromptOverride = NormalizeOptional(values.SystemPromptOverride, PromptOverrideMaxLength, "systemPromptOverride", errors);
		var draftPromptOverride = NormalizeOptional(values.DraftPromptOverride, PromptOverrideMaxLength, "draftPromptOverride", errors);

		if (errors.Count > 0)
		{
			return (null, errors);
		}

		return (new AddinProfileSettings
		{
			UserName = userName,
			Tone = tone,
			LengthPreference = length,
			Language = language,
			Greeting = existing.Greeting,
			Closing = existing.Closing,
			CustomInstructions = customInstructions,
			SystemPromptOverride = systemPromptOverride,
			DraftPromptOverride = draftPromptOverride,
			AnalyzedStyle = existing.AnalyzedStyle,
			Model = string.IsNullOrWhiteSpace(model) ? null : model,
			StylePersonalizationEnabled = values.StylePersonalizationEnabled,
			StyleProfileSummary = existing.StyleProfileSummary,
			StyleSampleCount = existing.StyleSampleCount,
			StyleLastRefreshedUtc = existing.StyleLastRefreshedUtc,
			StyleLastRefreshStatus = existing.StyleLastRefreshStatus,
			StyleLastRefreshError = existing.StyleLastRefreshError
		}, errors);
	}

	private static OutlookSettingsDefaultsDto BuildDefaults() =>
		new()
		{
			UserName = "User",
			Tone = "professional",
			Length = "concise",
			Language = "English",
			Model = string.Empty,
			CustomInstructions = string.Empty,
			SystemPromptOverride = string.Empty,
			DraftPromptOverride = string.Empty,
			StylePersonalizationEnabled = true
		};

	private static Dictionary<string, OutlookSettingsFieldMetadataDto> BuildFieldMetadata(
		OutlookSettingsDefaultsDto defaults,
		OutlookSettingsCatalogSnapshot catalog) =>
		new(StringComparer.OrdinalIgnoreCase)
		{
			["userName"] = new() { Label = "Your name", Description = "Used as mailbox-owner context for reply generation.", Recommended = defaults.UserName, MaxLength = UserNameMaxLength },
			["tone"] = new() { Label = "Tone", Description = "Sets the default voice Maily should prefer when rewriting or drafting.", Recommended = defaults.Tone, MaxLength = 0 },
			["length"] = new() { Label = "Reply length", Description = "Controls the default depth of generated replies.", Recommended = defaults.Length, MaxLength = 0 },
			["language"] = new() { Label = "Fallback language", Description = "Used only when the thread itself does not clearly establish the language.", Recommended = defaults.Language, MaxLength = 0 },
			["model"] = new() { Label = "Model", Description = "Leave blank for the Codex recommended model, or enter a model ID supported by your account.", Recommended = $"Use {catalog.DefaultModel}", MaxLength = ModelMaxLength },
			["customInstructions"] = new() { Label = "Custom instructions", Description = "Reusable guidance for all replies. Keep this short and stable.", Recommended = "Blank", IsAdvanced = true, MaxLength = CustomInstructionsMaxLength },
			["systemPromptOverride"] = new() { Label = "System prompt override", Description = $"Full expert override for the reply system prompt. Leave blank to keep {catalog.PromptCatalog.SystemPromptSourceLabel}.", Recommended = "Use current Maily baseline", IsAdvanced = true, MaxLength = PromptOverrideMaxLength },
			["draftPromptOverride"] = new() { Label = "Draft rewrite prompt override", Description = $"Expert override for how Maily rewrites an existing user draft. Leave blank to keep {catalog.PromptCatalog.DraftPromptSourceLabel}.", Recommended = "Use current Maily baseline", IsAdvanced = true, MaxLength = PromptOverrideMaxLength },
			["stylePersonalizationEnabled"] = new() { Label = "Use learned style", Description = "When enabled, Maily applies the reusable style profile inferred from Sent Items.", Recommended = defaults.StylePersonalizationEnabled ? "Enabled" : "Disabled", MaxLength = 0 }
		};

	private static OutlookStyleProfileStateDto BuildStyleState(AddinProfileSettings settings)
	{
		var parts = new List<string>();
		if (settings.StyleSampleCount > 0)
		{
			parts.Add($"{settings.StyleSampleCount} sent emails analyzed");
		}
		if (settings.StyleLastRefreshedUtc is not null)
		{
			parts.Add($"Last refresh: {settings.StyleLastRefreshedUtc:O}");
		}
		if (!string.IsNullOrWhiteSpace(settings.StyleLastRefreshStatus))
		{
			parts.Add($"Status: {settings.StyleLastRefreshStatus}");
		}

		return new OutlookStyleProfileStateDto
		{
			Enabled = settings.StylePersonalizationEnabled,
			Summary = settings.StyleProfileSummary ?? "Refresh style from Sent Items to build a reusable style profile.",
			AnalyzedStyle = settings.AnalyzedStyle ?? "Maily will store reusable guidance here after style analysis.",
			SampleCount = settings.StyleSampleCount,
			LastRefreshedUtc = settings.StyleLastRefreshedUtc?.ToString("o") ?? string.Empty,
			LastRefreshStatus = settings.StyleLastRefreshStatus ?? string.Empty,
			LastRefreshError = settings.StyleLastRefreshError ?? string.Empty,
			RefreshStateSummary = parts.Count == 0 ? "No learned style yet." : string.Join(" | ", parts)
		};
	}

	private static string NormalizeText(
		string? value,
		string fallback,
		int maxLength,
		string fieldName,
		Dictionary<string, string> errors,
		bool allowBlank)
	{
		var normalized = (value ?? string.Empty).Trim();
		if (string.IsNullOrWhiteSpace(normalized))
		{
			if (allowBlank)
			{
				return string.Empty;
			}

			if (!string.IsNullOrWhiteSpace(fallback))
			{
				return fallback;
			}

			errors[fieldName] = "This field is required.";
			return string.Empty;
		}

		if (normalized.Length > maxLength)
		{
			errors[fieldName] = $"Keep this under {maxLength} characters.";
		}

		return normalized;
	}

	private static string NormalizeOptional(
		string? value,
		int maxLength,
		string fieldName,
		Dictionary<string, string> errors)
	{
		var normalized = (value ?? string.Empty).Trim();
		if (normalized.Length > maxLength)
		{
			errors[fieldName] = $"Keep this under {maxLength} characters.";
		}

		return normalized;
	}

	private static string NormalizeEnum(
		string? value,
		IReadOnlyCollection<string> allowed,
		string fallback,
		string fieldName,
		Dictionary<string, string> errors)
	{
		var normalized = (value ?? string.Empty).Trim();
		if (allowed.Contains(normalized, StringComparer.OrdinalIgnoreCase))
		{
			return allowed.First(option => string.Equals(option, normalized, StringComparison.OrdinalIgnoreCase));
		}

		errors[fieldName] = "Choose one of the supported options.";
		return fallback;
	}
}
