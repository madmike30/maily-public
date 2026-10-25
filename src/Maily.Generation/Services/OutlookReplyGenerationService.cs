using Maily.Generation.Contracts;
using Maily.Generation.Models;

namespace Maily.Generation.Services;

public sealed class OutlookReplyGenerationService
{
	private readonly CodexRuntimeConfigLoader configLoader;
	private readonly PromptTemplateEngine promptTemplateEngine;
	private readonly OracleResponsesGateway oracleResponsesGateway;
	private readonly ProfileSettingsStore profileSettingsStore;
	private readonly GenerationOptions options;

	public OutlookReplyGenerationService(
		CodexRuntimeConfigLoader configLoader,
		PromptTemplateEngine promptTemplateEngine,
		OracleResponsesGateway oracleResponsesGateway,
		ProfileSettingsStore profileSettingsStore,
		GenerationOptions options)
	{
		this.configLoader = configLoader;
		this.promptTemplateEngine = promptTemplateEngine;
		this.oracleResponsesGateway = oracleResponsesGateway;
		this.profileSettingsStore = profileSettingsStore;
		this.options = options;
	}

	public async Task<GenerateReplyResponse> GenerateReplyAsync(GenerateReplyRequest request, string requestId, CancellationToken cancellationToken = default)
	{
		Validate(request);

		var storedSettings = profileSettingsStore.Get(request.Client.MailboxAddress, request.UserPreferences.ProfileId).Settings;
		var effectiveSettings = MergeSettings(storedSettings, request.UserPreferences);
		ApplyReplyAuthorIdentity(request, effectiveSettings);

		var config = configLoader.Load(effectiveSettings.Model);
		var systemPrompt = promptTemplateEngine.BuildSystemPrompt(request, effectiveSettings);
		var userPrompt = promptTemplateEngine.BuildUserPrompt(request, effectiveSettings);
		var generatedText = await oracleResponsesGateway.GenerateReplyAsync(config, systemPrompt, userPrompt, requestId, cancellationToken);
		var draftText = ReplySanitizer.NormalizeForOutlookReply(generatedText, effectiveSettings, request.Operation.CurrentDraft);
		if (string.IsNullOrWhiteSpace(draftText))
		{
			throw new InvalidOperationException("The generated reply contained no usable draft text after cleanup.");
		}
		var draftHtml = ReplyFormatter.ToOutlookHtml(draftText);

		return new GenerateReplyResponse
		{
			Draft = new GenerateReplyDraft
			{
				Html = draftHtml,
				Text = draftText
			},
			Meta = new GenerateReplyMeta
			{
				Model = config.Model,
				PromptVersion = options.PromptVersion,
				RequestId = requestId
			}
		};
	}

	private static AddinProfileSettings MergeSettings(AddinProfileSettings stored, GenerateReplyUserPreferences prefs)
	{
		var merged = new AddinProfileSettings
		{
			UserName = stored.UserName,
			Tone = stored.Tone,
			LengthPreference = stored.LengthPreference,
			Language = stored.Language,
			Greeting = stored.Greeting,
			Closing = stored.Closing,
			CustomInstructions = stored.CustomInstructions,
			SystemPromptOverride = stored.SystemPromptOverride,
			DraftPromptOverride = stored.DraftPromptOverride,
			AnalyzedStyle = stored.AnalyzedStyle,
			Model = stored.Model,
			StylePersonalizationEnabled = stored.StylePersonalizationEnabled,
			StyleProfileSummary = stored.StyleProfileSummary,
			StyleSampleCount = stored.StyleSampleCount,
			StyleLastRefreshedUtc = stored.StyleLastRefreshedUtc,
			StyleLastRefreshStatus = stored.StyleLastRefreshStatus,
			StyleLastRefreshError = stored.StyleLastRefreshError
		};

		if (!string.IsNullOrWhiteSpace(prefs.Tone))
		{
			merged.Tone = prefs.Tone;
		}

		if (!string.IsNullOrWhiteSpace(prefs.Length))
		{
			merged.LengthPreference = prefs.Length;
		}

		if (!string.IsNullOrWhiteSpace(prefs.Language))
		{
			merged.Language = prefs.Language;
		}

		return merged;
	}

	private static void ApplyReplyAuthorIdentity(GenerateReplyRequest request, AddinProfileSettings settings)
	{
		settings.UserName = ReplyAuthorIdentity.Resolve(request).Name;
	}

	private static void Validate(GenerateReplyRequest request)
	{
		if (string.IsNullOrWhiteSpace(request.Client.MailboxAddress))
		{
			throw new InvalidOperationException("client.mailboxAddress is required.");
		}

		if (string.IsNullOrWhiteSpace(request.UserPreferences.ProfileId))
		{
			request.UserPreferences.ProfileId = "default";
		}
	}

}
