using Maily.Generation.Models;
using Maily.Runtime;

namespace Maily.Generation.Services;

public sealed class CodexRuntimeConfigLoader
{
	private readonly PathResolver paths;
	private readonly GenerationOptions options;
	private readonly Func<string> installedAppResolver;
	public bool UsesInstalledApp => options.CodexExecutablePath?.Trim() == InstalledAppRuntime.Selection;

	public CodexRuntimeConfigLoader(PathResolver paths, GenerationOptions options, Func<string>? installedAppResolver = null)
	{
		this.paths = paths;
		this.options = options;
		this.installedAppResolver = installedAppResolver ?? (() => InstalledAppRuntime.Resolve());
	}

	public CodexRuntimeConfig Load(string? selectedModel = null)
	{
		return new CodexRuntimeConfig
		{
			CodexHome = paths.CodexHome,
			CodexExecutablePath = UsesInstalledApp
				? installedAppResolver()
				: string.IsNullOrWhiteSpace(options.CodexExecutablePath)
				? "codex"
				: options.CodexExecutablePath.Trim(),
			Model = !string.IsNullOrWhiteSpace(selectedModel)
				? selectedModel.Trim()
				: options.Model?.Trim() ?? string.Empty
		};
	}
}
