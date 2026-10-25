using System.Diagnostics;
namespace Maily.OutlookAddin.Setup;

internal static class OutlookShutdown
{
    // Both the GUI and installer/uninstaller scripts use OutlookShutdown.ps1.
    // The helper only attaches to an existing Outlook instance, and never kills it.
    public static OutlookCloseResult Run(string helperPath, int timeoutMilliseconds = 35000)
    {
        if (!File.Exists(helperPath)) return new(false);
        using var process = new Process();
        try
        {
            process.StartInfo = WindowsPowerShellBootstrap.Create(
                "-NoProfile -STA -ExecutionPolicy Bypass -File \"" + helperPath + "\" -CloseOutlook", Path.GetDirectoryName(helperPath)!);
            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            if (!process.WaitForExit(timeoutMilliseconds))
            {
                // Own worker only; it cannot create Outlook or a mailbox session.
                process.Kill(entireProcessTree: true);
                return new(false);
            }
            return new(process.ExitCode == 0);
        }
        catch { return new(false); }
    }
}
