using System.Diagnostics;

namespace Maily.OutlookAddin.Web.Services;

/// <summary>Owns one bridge child until it has exited, including cancellation cleanup.</summary>
public static class BridgeProcessRunner
{
    public static async Task<string> RunAsync(ProcessStartInfo startInfo, SemaphoreSlim gate,
        TimeSpan timeout, CancellationToken cancellationToken)
    {
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        deadline.CancelAfter(timeout);
        var acquired = false;
        Process? process = null;
        var started = false;
        try
        {
            await gate.WaitAsync(deadline.Token);
            acquired = true;
            deadline.Token.ThrowIfCancellationRequested();
            process = new Process { StartInfo = startInfo };
            if (!process.Start()) throw new InvalidOperationException("Could not start the Outlook bridge process.");
            started = true;
            var stdout = process.StandardOutput.ReadToEndAsync(deadline.Token);
            var stderr = process.StandardError.ReadToEndAsync(deadline.Token);
            await Task.WhenAll(process.WaitForExitAsync(deadline.Token), stdout, stderr);
            if (process.ExitCode != 0)
                throw new InvalidOperationException("Outlook bridge action failed. Retry the action.");
            return stdout.Result;
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            throw new TimeoutException("Outlook bridge action timed out. Retry after Outlook responds.");
        }
        finally
        {
            try
            {
                if (started && process is not null && !process.HasExited)
                {
                    process.Kill(entireProcessTree: true);
                    using var cleanup = new CancellationTokenSource(TimeSpan.FromSeconds(10));
                    await process.WaitForExitAsync(cleanup.Token);
                }
            }
            finally
            {
                // If OS cleanup failed, quarantine this gate until restart. An alive
                // mutating child must never overlap a later COM request.
                if (acquired && (!started || process is null || process.HasExited)) gate.Release();
                process?.Dispose();
            }
        }
    }
}