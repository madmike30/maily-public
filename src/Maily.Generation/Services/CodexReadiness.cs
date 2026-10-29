namespace Maily.Generation.Services;

/// <summary>Classifies untrusted CLI diagnostics without exposing their contents.</summary>
public static class CodexReadiness
{
    public const string SetupMessage = "Maily needs ChatGPT sign-in before it can generate replies, summaries, search answers, or daily briefings.";
    public const string SetupAction = "Open MailyOutlookAddinSetup.exe from the original extracted installer package and follow its ChatGPT sign-in guidance, then reopen the Maily Outlook pane. The installer includes the required runtime.";

    public static CodexReadinessStatus Evaluate(bool versionStarted, int versionExitCode,
        bool loginStarted, int loginExitCode, string stdout, string stderr, bool installedApp = false, string versionOutput = "")
    {
        var output = string.Join("\n", new[] { stdout.Trim(), stderr.Trim() }.Where(s => s.Length > 0));
        // Fail closed for API keys, unknown formats, extra provider diagnostics and failed probes.
        var signedIn = versionStarted && versionExitCode == 0 && loginStarted && loginExitCode == 0 &&
            (!installedApp || Maily.Runtime.InstalledAppRuntime.IsSupportedVersion(versionOutput)) &&
            string.Equals(output, "Logged in using ChatGPT", StringComparison.OrdinalIgnoreCase);
        return new CodexReadinessStatus(signedIn, versionStarted, versionExitCode, loginExitCode,
            signedIn ? "chatgpt" : "unverified",
            signedIn ? "Maily is ready and signed in with ChatGPT." : SetupMessage,
            signedIn ? string.Empty : installedApp ? Maily.Runtime.InstalledAppRuntime.MissingGuidance : SetupAction);
    }
}

public sealed record CodexReadinessStatus(bool Ok, bool CodexCommandFound, int CodexVersionExitCode,
    int LoginStatusExitCode, string AuthenticationStatus, string Message, string RecommendedAction)
{
    public bool LoginStatusOk => Ok;
}
