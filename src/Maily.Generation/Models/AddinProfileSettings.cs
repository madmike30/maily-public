namespace Maily.Generation.Models;

public sealed class AddinProfileSettings
{
	public string UserName { get; set; } = "User";
	public string Tone { get; set; } = "professional";
	public string LengthPreference { get; set; } = "concise";
	public string Language { get; set; } = "English";
	public string Greeting { get; set; } = "Hi";
	public string Closing { get; set; } = "Best regards";
	public string CustomInstructions { get; set; } = string.Empty;
	public string SystemPromptOverride { get; set; } = string.Empty;
	public string DraftPromptOverride { get; set; } = string.Empty;
	public string? AnalyzedStyle { get; set; }
	public string? Model { get; set; }
	public bool StylePersonalizationEnabled { get; set; } = true;
	public string? StyleProfileSummary { get; set; }
	public int StyleSampleCount { get; set; }
	public DateTimeOffset? StyleLastRefreshedUtc { get; set; }
	public string? StyleLastRefreshStatus { get; set; }
	public string? StyleLastRefreshError { get; set; }
}
