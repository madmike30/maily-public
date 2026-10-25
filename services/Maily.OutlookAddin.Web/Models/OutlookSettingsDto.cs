namespace Maily.OutlookAddin.Web.Models;

public sealed class OutlookSettingsDto
{
	public string ProfileId { get; init; } = "default";
	public string UpdatedUtc { get; init; } = string.Empty;
	public OutlookSettingsValuesDto Values { get; init; } = new();
	public OutlookSettingsDefaultsDto Defaults { get; init; } = new();
	public OutlookSettingsOptionsDto Options { get; init; } = new();
	public OutlookSettingsPromptCatalogDto PromptCatalog { get; init; } = new();
	public Dictionary<string, OutlookSettingsFieldMetadataDto> FieldMetadata { get; init; } = new(StringComparer.OrdinalIgnoreCase);
	public Dictionary<string, string> FieldErrors { get; init; } = new(StringComparer.OrdinalIgnoreCase);
	public List<string> AdvancedWarnings { get; init; } = [];
	public OutlookStyleProfileStateDto Style { get; init; } = new();

	// Backward-compatible flat fields for existing panes that only need defaults.
	public string UserName { get; init; } = "User";
	public string Tone { get; init; } = "professional";
	public string Length { get; init; } = "concise";
	public string Language { get; init; } = "English";
	public string? Model { get; init; }
	public string CustomInstructions { get; init; } = string.Empty;
	public string SystemPromptOverride { get; init; } = string.Empty;
	public string DraftPromptOverride { get; init; } = string.Empty;
	public bool StylePersonalizationEnabled { get; init; } = true;
	public string StyleProfileSummary { get; init; } = string.Empty;
	public string AnalyzedStyle { get; init; } = string.Empty;
	public int StyleSampleCount { get; init; }
	public string StyleLastRefreshedUtc { get; init; } = string.Empty;
	public string StyleLastRefreshStatus { get; init; } = string.Empty;
	public string StyleLastRefreshError { get; init; } = string.Empty;
}

public sealed class OutlookSettingsValuesDto
{
	public string UserName { get; init; } = "User";
	public string Tone { get; init; } = "professional";
	public string Length { get; init; } = "concise";
	public string Language { get; init; } = "English";
	public string Model { get; init; } = string.Empty;
	public string CustomInstructions { get; init; } = string.Empty;
	public string SystemPromptOverride { get; init; } = string.Empty;
	public string DraftPromptOverride { get; init; } = string.Empty;
	public bool StylePersonalizationEnabled { get; init; } = true;
}

public sealed class OutlookSettingsDefaultsDto
{
	public string UserName { get; init; } = "User";
	public string Tone { get; init; } = "professional";
	public string Length { get; init; } = "concise";
	public string Language { get; init; } = "English";
	public string Model { get; init; } = string.Empty;
	public string CustomInstructions { get; init; } = string.Empty;
	public string SystemPromptOverride { get; init; } = string.Empty;
	public string DraftPromptOverride { get; init; } = string.Empty;
	public bool StylePersonalizationEnabled { get; init; } = true;
}

public sealed class OutlookSettingsOptionsDto
{
	public List<string> Tone { get; init; } = [];
	public List<string> Length { get; init; } = [];
	public List<string> Language { get; init; } = [];
	public List<string> Model { get; init; } = [];
}

public sealed class OutlookSettingsPromptCatalogDto
{
	public string SystemPromptBaseline { get; init; } = string.Empty;
	public string DraftPromptBaseline { get; init; } = string.Empty;
	public string SystemPromptSourceLabel { get; init; } = string.Empty;
	public string DraftPromptSourceLabel { get; init; } = string.Empty;
}

public sealed class OutlookSettingsFieldMetadataDto
{
	public string Label { get; init; } = string.Empty;
	public string Description { get; init; } = string.Empty;
	public string Recommended { get; init; } = string.Empty;
	public bool IsAdvanced { get; init; }
	public int MaxLength { get; init; }
}

public sealed class OutlookStyleProfileStateDto
{
	public bool Enabled { get; init; } = true;
	public string Summary { get; init; } = string.Empty;
	public string AnalyzedStyle { get; init; } = string.Empty;
	public int SampleCount { get; init; }
	public string LastRefreshedUtc { get; init; } = string.Empty;
	public string LastRefreshStatus { get; init; } = string.Empty;
	public string LastRefreshError { get; init; } = string.Empty;
	public string RefreshStateSummary { get; init; } = string.Empty;
}

public sealed class OutlookSettingsSaveRequestDto
{
	public string ProfileId { get; init; } = "default";
	public OutlookSettingsValuesDto Values { get; init; } = new();
}
