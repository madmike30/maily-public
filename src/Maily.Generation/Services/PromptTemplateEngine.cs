using System.Text.RegularExpressions;
using Maily.Generation.Contracts;
using Maily.Generation.Models;

namespace Maily.Generation.Services;

public sealed class PromptTemplateEngine
{
	private readonly PathResolver paths;

	public PromptTemplateEngine(PathResolver paths)
	{
		this.paths = paths;
	}

	public string BuildSystemPrompt(GenerateReplyRequest request, AddinProfileSettings settings)
	{
		var root = paths.PromptsRoot;
		var systemPrompt = string.IsNullOrWhiteSpace(settings.SystemPromptOverride)
			? ReadOrDefault(
			Path.Combine(root, "prompt-system.md"),
			"""
			- You are assisting {{userName}} in drafting concise professional email replies.
			- Today's date is {{currentDateTime}}.
			- Provide only the reply text without any signature.
			- Match tone, length, and language preferences.

			[Email-Exchange]
			{{emailExchange}}
			[/Email-Exchange]
			""")
			: settings.SystemPromptOverride;

		var emailSamples = ReadOrDefault(Path.Combine(root, "email-samples.md"), "No samples configured.");
		emailSamples = LimitText(emailSamples.Replace("{{userName}}", settings.UserName, StringComparison.Ordinal), 1800);

		var exchange = BuildEmailExchange(request);
		var exchangeLimit = string.IsNullOrWhiteSpace(request.Operation.CurrentDraft) ? 2400 : 1600;
		exchange = LimitText(exchange, exchangeLimit);

		var styleInstructions = BuildStyleInstructions(settings);
		var prompt = systemPrompt
			.Replace("{{emailExchange}}", exchange, StringComparison.Ordinal)
			.Replace("{{userName}}", settings.UserName, StringComparison.Ordinal)
			.Replace("{{emailSamples}}", emailSamples, StringComparison.Ordinal)
			.Replace("{{currentDateTime}}", DateTime.Now.ToString("f"), StringComparison.Ordinal);

		if (!string.IsNullOrWhiteSpace(styleInstructions))
		{
			prompt += $"{Environment.NewLine}{Environment.NewLine}## Writing style preferences{Environment.NewLine}{Environment.NewLine}{styleInstructions}";
		}

		prompt += $"{Environment.NewLine}{Environment.NewLine}## Language rules{Environment.NewLine}{Environment.NewLine}{BuildLanguageInstructions(request)}";

		if (!string.IsNullOrWhiteSpace(request.Operation.CurrentDraft))
		{
			prompt +=
				$"{Environment.NewLine}{Environment.NewLine}## User typed draft{Environment.NewLine}{Environment.NewLine}" +
				"- Treat this as the primary source for intent, facts, commitments, language, and stance." +
				$"{Environment.NewLine}- Rewrite it cleanly without changing concrete details or meaning." +
				$"{Environment.NewLine}{Environment.NewLine}[Current-Draft]{Environment.NewLine}{LimitText(request.Operation.CurrentDraft, 1200)}{Environment.NewLine}[/Current-Draft]";
		}

		return prompt;
	}

	public string BuildUserPrompt(GenerateReplyRequest request, AddinProfileSettings settings)
	{
		var mode = request.Operation.Mode?.Trim().ToLowerInvariant() ?? "generate-initial";
		var explicitLanguageOverride = NormalizeLanguageOverride(request.Operation.OutputLanguageOverride);
		if (!string.IsNullOrWhiteSpace(request.Operation.CurrentDraft))
		{
			var iterate = ReadOrDefault(
				Path.Combine(paths.PromptsRoot, "prompt-iterate-on-draft.md"),
				"""
				[Current draft]
				{{currentDraft}}
				[/Current draft]

				[Request]
				{{userRequest}}
				[/Request]
				""");
			if (!string.IsNullOrWhiteSpace(settings.DraftPromptOverride))
			{
				iterate = settings.DraftPromptOverride;
			}

			var instruction = request.Operation.UserInstruction;
			if (string.IsNullOrWhiteSpace(instruction))
			{
				var draftGuidance = BuildDraftRewriteGuidance(request.Operation.CurrentDraft, explicitLanguageOverride);
				instruction = mode == "regenerate"
					? BuildDraftRewriteInstruction("Refine this draft", explicitLanguageOverride, draftGuidance)
					: BuildDraftRewriteInstruction("Use this existing reply draft as the starting point", explicitLanguageOverride, draftGuidance);
			}

			return iterate
				.Replace("{{currentDraft}}", request.Operation.CurrentDraft, StringComparison.Ordinal)
				.Replace("{{userRequest}}", instruction, StringComparison.Ordinal);
		}

		return string.IsNullOrWhiteSpace(request.Operation.UserInstruction)
			? BuildFreshReplyInstruction(explicitLanguageOverride)
			: request.Operation.UserInstruction!;
	}

	private static string BuildDraftRewriteInstruction(string prefix, string explicitLanguageOverride, string draftGuidance)
	{
		if (!string.IsNullOrWhiteSpace(explicitLanguageOverride))
		{
			return $"{prefix}. Rephrase it in the user's style instead of keeping the same wording. Preserve the user's intent and concrete details, but write the final reply fully in {explicitLanguageOverride}. It is allowed to change the draft language because the user explicitly selected {explicitLanguageOverride}.{Environment.NewLine}{draftGuidance}";
		}

		return $"{prefix}. Rephrase it in the user's style instead of keeping the same wording. Preserve the user's intent, language, and concrete details, complete it only where needed, and improve clarity and flow without changing the meaning.{Environment.NewLine}{draftGuidance}";
	}

	private static string BuildFreshReplyInstruction(string explicitLanguageOverride)
	{
		var instruction =
			"Draft the reply to the latest email in the thread. Include the useful response content in one or two concise paragraphs. " +
			"Do not return only a greeting, closing, acknowledgement, or signature.";
		return string.IsNullOrWhiteSpace(explicitLanguageOverride)
			? instruction
			: $"{instruction} Write the final answer fully in {explicitLanguageOverride}.";
	}

	private static string BuildDraftRewriteGuidance(string currentDraft, string explicitLanguageOverride)
	{
		var normalized = currentDraft.Replace("\r\n", "\n").Trim();
		var wordCount = normalized.Split([' ', '\n', '\t'], StringSplitOptions.RemoveEmptyEntries).Length;
		var endsLikeCompleteSentence =
			normalized.EndsWith(".", StringComparison.Ordinal) ||
			normalized.EndsWith("!", StringComparison.Ordinal) ||
			normalized.EndsWith("?", StringComparison.Ordinal) ||
			normalized.EndsWith("。", StringComparison.Ordinal) ||
			normalized.EndsWith("؟", StringComparison.Ordinal);

		var lines = new List<string>
		{
			"- Rephrase naturally instead of echoing the same wording.",
			"- Preserve meaning and concrete details.",
			"- Do not add a greeting unless the draft already includes one.",
			"- Do not add a sign-off unless the draft already includes one.",
			"- Do not add bullets, placeholders, headings, extra requests, or new facts."
		};

		if (string.IsNullOrWhiteSpace(explicitLanguageOverride))
		{
			lines.Add("- Keep the final reply in the same language as the current draft unless translation was requested.");
		}
		else
		{
			lines.Add($"- Rewrite the final reply fully in {explicitLanguageOverride} because the user explicitly selected that output language.");
			lines.Add("- Preserve the user's meaning and details while translating naturally.");
		}

		if (wordCount >= 10 && endsLikeCompleteSentence)
		{
			lines.Add("- This draft is mostly complete. Prefer a compact stylistic rewrite over expansion.");
		}
		else
		{
			lines.Add("- This draft may be partial. Complete it carefully without changing the intended meaning.");
		}

		if (wordCount <= 3)
		{
			lines.Add("- The draft is very short. Expand it into one or two clear sentences.");
		}

		return string.Join(Environment.NewLine, lines);
	}

	private static string ReadOrDefault(string path, string fallback) =>
		File.Exists(path) ? File.ReadAllText(path) : fallback;

	private static string LimitText(string? value, int maxLength)
	{
		var text = (value ?? string.Empty).Trim();
		return text.Length <= maxLength
			? text
			: text[..maxLength].TrimEnd() + "\n[... truncated ...]";
	}

	private static string BuildStyleInstructions(AddinProfileSettings settings)
	{
		var lines = new List<string>();
		if (!string.IsNullOrWhiteSpace(settings.Tone))
		{
			lines.Add($"- Tone: {settings.Tone}");
		}

		if (!string.IsNullOrWhiteSpace(settings.LengthPreference))
		{
			lines.Add($"- Reply length: {settings.LengthPreference}");
		}

		if (!string.IsNullOrWhiteSpace(settings.Greeting))
		{
			lines.Add($"- Greeting style: {settings.Greeting}");
		}

		if (!string.IsNullOrWhiteSpace(settings.CustomInstructions))
		{
			lines.Add(string.Empty);
			lines.Add(settings.CustomInstructions);
		}

		if (settings.StylePersonalizationEnabled && !string.IsNullOrWhiteSpace(settings.AnalyzedStyle))
		{
			lines.Add(string.Empty);
			lines.Add("Writing style inferred from sent emails:");
			lines.Add(settings.AnalyzedStyle);
		}

		return string.Join(Environment.NewLine, lines);
	}

	private static string BuildLanguageInstructions(GenerateReplyRequest request)
	{
		var dominantLanguage = ThreadLanguageHeuristics.DescribeDominantLanguage(
			request.Operation.CurrentDraft,
			request.Compose.BodyText,
			request.Compose.BodyHtml,
			request.Compose.Subject);
		var explicitOverride = NormalizeLanguageOverride(request.Operation.OutputLanguageOverride);

		var lines = new List<string>
		{
			$"- The dominant language visible in this request appears to be {dominantLanguage}.",
			"- Preserve the user's own wording and phrasing when refining an existing draft."
		};

		if (!string.IsNullOrWhiteSpace(explicitOverride))
		{
			lines.Insert(0, $"- The user explicitly selected {explicitOverride} for this reply. Write the final reply fully in {explicitOverride}.");
			lines.Insert(1, "- This explicit language selection overrides the thread-language default for this generation.");
		}
		else
		{
			lines.Insert(0, "- Saved language settings and writing samples do not override content language. Identify the language from the actual message, ignoring headers, addresses, signatures, and quoted history. Always answer in the same language as the user's typed reply draft when a current draft is present.");
			lines.Insert(1, "- If there is no typed draft yet, answer in the same language as the latest email message or dominant thread language.");
		}

		lines.Insert(2, "- Do not translate to English or any other language unless the email thread explicitly asks for translation or the user explicitly selected a different output language.");

		if (!string.IsNullOrWhiteSpace(request.Operation.CurrentDraft))
		{
			lines.Add("- A current draft is present in this request and should be treated as the primary source of language, tone, and intent.");
		}

		return string.Join(Environment.NewLine, lines);
	}

	private static string NormalizeLanguageOverride(string? value)
	{
		var normalized = value?.Trim().ToLowerInvariant();
		return normalized switch
		{
			"hebrew" => "Hebrew",
			"english" => "English",
			_ => string.Empty
		};
	}

	private static string BuildEmailExchange(GenerateReplyRequest request)
	{
		var bodyText = request.Compose.BodyText;
		if (string.IsNullOrWhiteSpace(bodyText))
		{
			bodyText = StripHtml(request.Compose.BodyHtml ?? string.Empty);
		}

		var to = request.Compose.To.Count == 0 ? "(none)" : string.Join(", ", request.Compose.To);
		var cc = request.Compose.Cc.Count == 0 ? "(none)" : string.Join(", ", request.Compose.Cc);
		var mailboxDisplay = string.IsNullOrWhiteSpace(request.Client.MailboxDisplayName)
			? "Mailbox owner"
			: request.Client.MailboxDisplayName;
		var mailboxAddress = request.Client.MailboxAddress;
		var senderDisplay = string.IsNullOrWhiteSpace(request.Compose.FromDisplayName)
			? "(unknown sender)"
			: request.Compose.FromDisplayName;
		var senderAddress = string.IsNullOrWhiteSpace(request.Compose.FromAddress)
			? "(unknown address)"
			: request.Compose.FromAddress;
		var prefix = string.IsNullOrWhiteSpace(request.Operation.CurrentDraft)
			? string.Empty
			: $"Note: the email exchange below is supporting context only. The current draft is the primary source for the actual reply content.{Environment.NewLine}{Environment.NewLine}";
		return prefix +
			$"You are writing as: {mailboxDisplay} <{mailboxAddress}>{Environment.NewLine}" +
			$"Mailbox: {mailboxAddress}{Environment.NewLine}" +
			$"Subject: {request.Compose.Subject}{Environment.NewLine}" +
			$"From: {senderDisplay} <{senderAddress}>{Environment.NewLine}" +
			$"To: {to}{Environment.NewLine}" +
			$"Cc: {cc}{Environment.NewLine}" +
			$"Platform: {request.Client.Platform}{Environment.NewLine}{Environment.NewLine}" +
			bodyText;
	}

	private static string StripHtml(string value)
	{
		if (string.IsNullOrWhiteSpace(value))
		{
			return string.Empty;
		}

		var withoutTags = Regex.Replace(value, "<[^>]+>", " ");
		return Regex.Replace(withoutTags, "\\s{2,}", " ").Trim();
	}
}
