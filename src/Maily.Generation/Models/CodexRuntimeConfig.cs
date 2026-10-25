namespace Maily.Generation.Models;

public sealed class CodexRuntimeConfig
{
	public string ReasoningEffort { get; set; } = string.Empty;
	public string CodexHome { get; set; } = string.Empty;
	public string CodexExecutablePath { get; set; } = "codex";
	public string Model { get; set; } = string.Empty;
}
