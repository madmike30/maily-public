namespace Maily.Maui.Models;

public sealed class CodexLlmConfig
{
	public required string BaseUrl { get; init; }
	public required string ApiKey { get; init; }
	public required string Model { get; init; }
	public required IReadOnlyDictionary<string, string> ApiHeaders { get; init; }
	public required IReadOnlyList<string> ModelChoices { get; init; }
}
