namespace Maily.Runtime;

// Shared by setup and generation so desktop updates never leave a pinned, removed executable.
public static class InstalledAppRuntime
{
    // Child-only: preserve the desktop authentication home while excluding paid API overrides.
    public static void UseChatGptSubscription(System.Diagnostics.ProcessStartInfo start)
    {
        foreach (var key in new[] { "OPENAI_API_KEY", "CODEX_API_KEY", "AZURE_OPENAI_API_KEY" })
            start.Environment.Remove(key);
    }

    public const string Selection = "installed-app";
    public const string MissingGuidance = "Open the ChatGPT or Codex desktop app, finish setup and sign in, then retry Maily. If the app is missing or outdated, install or update it first.";

    public static string ReadPackageMode(string packageRoot)
    {
        var marker = Path.Combine(packageRoot, "ai-runtime-mode.txt");
        if (!File.Exists(marker)) return "Bundled";
        var value = File.ReadAllText(marker).Trim();
        return value is "InstalledApp" or "Bundled" ? value : throw new InvalidOperationException("Invalid AI runtime package mode. Obtain a complete Maily package.");
    }

    public static IReadOnlyList<string> GetCandidates(string? localAppData = null)
    {
        localAppData ??= Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (string.IsNullOrWhiteSpace(localAppData)) return [];
        var result = new List<string>();
        foreach (var app in new[] { "Codex", "ChatGPT" })
        {
            var bin = Path.Combine(localAppData, "OpenAI", app, "bin");
            try
            {
                if (!Directory.Exists(bin) || IsRedirected(bin)) continue;
                Add(Path.Combine(bin, "codex.exe"));
                foreach (var version in Directory.EnumerateDirectories(bin))
                    if (!IsRedirected(version)) Add(Path.Combine(version, "codex.exe"));
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
        return result.Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(File.GetLastWriteTimeUtc).ThenBy(path => path, StringComparer.OrdinalIgnoreCase).ToArray();

        void Add(string file)
        {
            if (File.Exists(file) && !IsRedirected(file)) result.Add(Path.GetFullPath(file));
        }
    }

    private static bool IsRedirected(string path)
    {
        for (var current = Path.GetFullPath(path); !string.IsNullOrEmpty(current); current = Path.GetDirectoryName(current))
            if ((File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0) return true;
        return false;
    }

    public static string Resolve(string? localAppData = null) =>
        GetCandidates(localAppData).FirstOrDefault() ?? throw new InstalledAppRuntimeUnavailableException();

    public static bool IsChatGptLogin(bool started, int exitCode, string output) =>
        started && exitCode == 0 && string.Equals(output.Trim(), "Logged in using ChatGPT", StringComparison.OrdinalIgnoreCase);

    public static bool IsSupportedVersion(string output)
    {
        var match = System.Text.RegularExpressions.Regex.Match(output.Trim(), @"^codex-cli (?<version>\d+\.\d+\.\d+)(?<suffix>-[A-Za-z0-9.-]+)?$");
        if (!match.Success || !Version.TryParse(match.Groups["version"].Value, out var version)) return false;
        var baseline = new Version(0, 151, 0);
        return version > baseline || version == baseline && !match.Groups["suffix"].Success;
    }
}

public sealed class InstalledAppRuntimeUnavailableException() : InvalidOperationException(InstalledAppRuntime.MissingGuidance);
