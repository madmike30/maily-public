namespace Maily.Generation.Models;

public sealed class GenerationOptions
{
	public string? RepoRoot { get; set; }
	public string? PromptsRoot { get; set; }
	public string? SettingsFilePath { get; set; }
	public string? CodexHome { get; set; }
	public string? CodexExecutablePath { get; set; }
	public string? Model { get; set; }
	public string PromptVersion { get; set; } = "v1";
}
