using Maily.Generation.Services;

namespace Maily.OutlookAddin.Web.Models;

public sealed class OutlookSettingsCatalog
{
	private const string FallbackSystemPrompt =
		"""
		- You are assisting {{userName}} in drafting concise professional email replies.
		- Today's date is {{currentDateTime}}.
		- Provide only the reply text without any signature.
		- Match tone, length, and language preferences.

		[Email-Exchange]
		{{emailExchange}}
		[/Email-Exchange]
		""";

	private const string FallbackDraftPrompt =
		"""
		[Current draft]
		{{currentDraft}}
		[/Current draft]

		[Request]
		{{userRequest}}
		[/Request]
		""";

	private readonly PathResolver paths;

	public OutlookSettingsCatalog(PathResolver paths)
	{
		this.paths = paths;
	}

	public OutlookSettingsCatalogSnapshot BuildSnapshot(string? currentModel)
	{

		var modelOptions = OrderModels(Array.Empty<string>(), currentModel);

		return new OutlookSettingsCatalogSnapshot
		{
			DefaultModel = "Codex recommended default",
			ModelOptions = modelOptions,
			PromptCatalog = new OutlookSettingsPromptCatalogDto
			{
				SystemPromptBaseline = ReadPrompt("prompt-system.md", FallbackSystemPrompt),
				DraftPromptBaseline = ReadPrompt("prompt-iterate-on-draft.md", FallbackDraftPrompt),
				SystemPromptSourceLabel = DescribePromptSource("prompt-system.md"),
				DraftPromptSourceLabel = DescribePromptSource("prompt-iterate-on-draft.md")
			}
		};
	}

	private string ReadPrompt(string fileName, string fallback)
	{
		var path = Path.Combine(paths.PromptsRoot, fileName);
		return File.Exists(path)
			? File.ReadAllText(path)
			: fallback;
	}

	private string DescribePromptSource(string fileName)
	{
		var path = Path.Combine(paths.PromptsRoot, fileName);
		return File.Exists(path)
			? $"Current baseline from {fileName}"
			: "Built-in fallback baseline";
	}

	private static List<string> OrderModels(IEnumerable<string> rawValues, string? currentModel)
	{
		var ordered = new List<string>();
		AddIfPresent(ordered, currentModel);

		foreach (var value in rawValues
			.Where(value => !string.IsNullOrWhiteSpace(value))
			.Distinct(StringComparer.OrdinalIgnoreCase)
			.OrderBy(value => value, StringComparer.OrdinalIgnoreCase))
		{
			AddIfPresent(ordered, value);
		}

		return ordered;
	}

	private static void AddIfPresent(List<string> values, string? candidate)
	{
		var normalized = candidate?.Trim();
		if (string.IsNullOrWhiteSpace(normalized))
		{
			return;
		}

		if (!values.Contains(normalized, StringComparer.OrdinalIgnoreCase))
		{
			values.Add(normalized);
		}
	}

}

public sealed class OutlookSettingsCatalogSnapshot
{
	public string DefaultModel { get; init; } = "Codex recommended default";
	public IReadOnlyList<string> ModelOptions { get; init; } = Array.Empty<string>();
	public OutlookSettingsPromptCatalogDto PromptCatalog { get; init; } = new();
}
