using Maily.Generation.Models;

namespace Maily.Generation.Services;

public sealed class CodexRuntimeConfigLoader
{
	private readonly PathResolver paths;
	private readonly GenerationOptions options;

	public CodexRuntimeConfigLoader(PathResolver paths, GenerationOptions options)
	{
		this.paths = paths;
		this.options = options;
	}

	public CodexRuntimeConfig Load(string? selectedModel = null)
	{
		return new CodexRuntimeConfig
		{
			CodexHome = paths.CodexHome,
			CodexExecutablePath = string.IsNullOrWhiteSpace(options.CodexExecutablePath)
				? "codex"
				: options.CodexExecutablePath.Trim(),
			Model = !string.IsNullOrWhiteSpace(selectedModel)
				? selectedModel.Trim()
				: options.Model?.Trim() ?? string.Empty
		};
	}
}
