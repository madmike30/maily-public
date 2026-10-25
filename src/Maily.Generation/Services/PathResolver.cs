using Maily.Generation.Models;

namespace Maily.Generation.Services;

public sealed class PathResolver
{
	private readonly GenerationOptions options;
	private readonly Lazy<string> repoRoot;

	public PathResolver(GenerationOptions options)
	{
		this.options = options;
		repoRoot = new Lazy<string>(ResolveRepoRoot);
	}

	public string RepoRoot => repoRoot.Value;

	public string PromptsRoot
	{
		get
		{
			if (!string.IsNullOrWhiteSpace(options.PromptsRoot))
			{
				return options.PromptsRoot;
			}

			return Path.Combine(RepoRoot, "smoke_data", "prompts");
		}
	}

	public string SettingsFilePath
	{
		get
		{
			if (!string.IsNullOrWhiteSpace(options.SettingsFilePath))
			{
				return options.SettingsFilePath;
			}

			return Path.Combine(RepoRoot, "output", "addin", "addin_profile_settings.json");
		}
	}

	public string CodexHome
	{
		get
		{
			if (!string.IsNullOrWhiteSpace(options.CodexHome))
			{
				return options.CodexHome;
			}

			var env = Environment.GetEnvironmentVariable("CODEX_HOME");
			if (!string.IsNullOrWhiteSpace(env))
			{
				return env;
			}

			return string.Empty;
		}
	}

	private string ResolveRepoRoot()
	{
		if (!string.IsNullOrWhiteSpace(options.RepoRoot))
		{
			return options.RepoRoot;
		}

		var starts = new[]
		{
			AppContext.BaseDirectory,
			Directory.GetCurrentDirectory()
		};

		foreach (var start in starts.Distinct(StringComparer.OrdinalIgnoreCase))
		{
			var current = new DirectoryInfo(start);
			while (current is not null)
			{
				var prompts = Path.Combine(current.FullName, "smoke_data", "prompts");
				if (Directory.Exists(prompts))
				{
					return current.FullName;
				}

				current = current.Parent;
			}
		}

		throw new DirectoryNotFoundException("Could not resolve repository root for prompt templates.");
	}
}
