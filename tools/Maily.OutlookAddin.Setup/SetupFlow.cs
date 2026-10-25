namespace Maily.OutlookAddin.Setup;

// This model is shared by the form and synthetic tests. It never performs host actions.
internal enum SetupAction { Install, Connect, CheckConnection, OpenOutlook, Close }
internal sealed class SetupFlow
{
    public bool TryResume(string summaryJson, string expectedInstallRoot, bool installedPackageMatches, string expectedPackageFingerprint)
    {
        if (!installedPackageMatches || string.IsNullOrWhiteSpace(expectedPackageFingerprint)) return false;
        try
        {
            using var document = System.Text.Json.JsonDocument.Parse(summaryJson);
            var root = document.RootElement;
            if (root.GetProperty("packageFingerprint").GetString() != expectedPackageFingerprint) return false;
            if (root.GetProperty("status").GetString() != "activation-required" || root.GetProperty("failureCount").GetInt32() != 0) return false;
            if (!string.Equals(Path.GetFullPath(root.GetProperty("installRoot").GetString()!).TrimEnd('\\', '/'), Path.GetFullPath(expectedInstallRoot).TrimEnd('\\', '/'), StringComparison.OrdinalIgnoreCase)) return false;
            ApplicationInstalled = true;
            NextAction = SetupAction.CheckConnection;
            Message = "Maily is already installed. Continue checking its Outlook connection.";
            return true;
        }
        catch (Exception ex) when (ex is System.Text.Json.JsonException or KeyNotFoundException or InvalidOperationException or ArgumentException or NotSupportedException or FormatException or OverflowException) { return false; }
    }
    public SetupAction NextAction { get; private set; } = SetupAction.Install;
    public bool ApplicationInstalled { get; private set; }
    public bool CanConnectAutomatically => NextAction == SetupAction.Connect && !LocalRemovalComplete;
    public string? SaveConnectionProgress(string installRoot, string packageFingerprint) => ApplicationInstalled && !string.IsNullOrWhiteSpace(packageFingerprint)
        ? System.Text.Json.JsonSerializer.Serialize(new { status = "activation-required", failureCount = 0, installRoot, packageFingerprint }) : null;
    public SetupAction AfterConnectionGuide => ApplicationInstalled ? SetupAction.CheckConnection : SetupAction.Install;
    public bool LocalRemovalComplete { get; private set; }
    public bool RepairRequired { get; private set; }
    public string ConnectionInstructions => RepairRequired
        ? "Outlook has an older Maily connection that you installed. In My add-ins, remove only that older Maily entry. Then choose Add a custom add-in > Add from file. Use Copy setup file below and paste it into the file picker. Approve Microsoft's prompt, then reopen Outlook. Do not remove an organization-managed entry."
        : "If you already added Maily, allow Outlook time to update, then check again. Otherwise, open My add-ins and choose Add a custom add-in > Add from file. Use Copy setup file below and paste it into the file picker. Approve Microsoft's prompt, then reopen Outlook. If Maily is already listed or custom add-ins are unavailable, ask your IT administrator for help.";
    public string Message { get; private set; } = "Ready to set up Maily.";
    public string PrimaryText => LocalRemovalComplete ? "Finish removal" : NextAction switch {
        SetupAction.Connect => "Connect Outlook", SetupAction.CheckConnection => "Check connection",
        SetupAction.OpenOutlook => "Open Outlook", SetupAction.Close => "Close", _ => "Install" };
    public void InstallationFinished(int exitCode, bool applicationInstalled, string activationState = "unknown")
    {
        RepairRequired = exitCode == 20 && activationState == "repair-required";
        ApplicationInstalled |= applicationInstalled;
        NextAction = exitCode switch { 0 => SetupAction.OpenOutlook, 20 => SetupAction.Connect,
            21 => SetupAction.CheckConnection, 23 => SetupAction.Close, _ => ApplicationInstalled ? SetupAction.CheckConnection : SetupAction.Install };
        Message = exitCode switch {
            0 => "Ready. Open Maily Settings for the quick tour and optional writing-style learning.",
            20 => "One step remains: connect Maily to your Outlook account.",
            21 => "Maily is installed. Open Outlook, then check the connection.",
            23 => "This Outlook connection needs help from your administrator or Maily support. Do not remove or reinstall the account add-in.",
            _ => ApplicationInstalled ? "Maily is installed. Its Outlook connection needs another check." : "Setup needs attention. Your settings and sign-in are preserved." };
    }
    public void ConnectionFinished(int exitCode, string activationState = "unknown")
    {
        RepairRequired = exitCode == 20 && activationState == "repair-required";
        NextAction = exitCode switch { 0 => SetupAction.OpenOutlook, 20 => SetupAction.Connect, 23 => SetupAction.Close, _ => SetupAction.CheckConnection };
        Message = exitCode == 0 ? "Ready. Open Maily Settings for the quick tour and optional writing-style learning." : exitCode == 23 ? "This Outlook connection needs help from your administrator or Maily support. Do not remove or reinstall the account add-in." : "The connection is not confirmed yet. If you just added Maily, allow Outlook time to update and check again.";
    }
    public void RemovalFinished(int exitCode)
    {
        LocalRemovalComplete = exitCode is 0 or 22;
        NextAction = LocalRemovalComplete ? SetupAction.Close : SetupAction.Install;
        Message = LocalRemovalComplete ? "Maily was removed from this PC. Its Outlook account entry may still need removal." : "Maily could not be removed. Review the details before trying again.";
    }
}
