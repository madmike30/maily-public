using System.Diagnostics;
using Maily.Generation.Models;
using Maily.Generation.Services;
using Microsoft.Extensions.Logging.Abstractions;

var evidence = Environment.GetEnvironmentVariable("MAILY_STDIN_TEST_DIR");
if (args.FirstOrDefault() == "--descendant")
{
    File.WriteAllText(Path.Combine(evidence!, "descendant.pid"), Environment.ProcessId.ToString());
    await Task.Delay(Timeout.Infinite);
    return;
}
if (args.FirstOrDefault() == "exec")
{
    File.WriteAllText(Path.Combine(evidence!, "parent.pid"), Environment.ProcessId.ToString());
    using var child = Process.Start(new ProcessStartInfo(Environment.ProcessPath!, "--descendant") { UseShellExecute = false });
    // Intentionally never read stdin: the multi-megabyte prompt must fill its pipe.
    await Task.Delay(Timeout.Infinite);
    return;
}
evidence = Path.Combine(Path.GetTempPath(), "Maily-Stdin-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(evidence);
Environment.SetEnvironmentVariable("MAILY_STDIN_TEST_DIR", evidence);
Console.WriteLine("Artifacts: " + evidence);
using var cancellation = new CancellationTokenSource();
var gateway = new OracleResponsesGateway(NullLogger<OracleResponsesGateway>.Instance);
var operation = gateway.GenerateReplyAsync(new CodexRuntimeConfig { CodexExecutablePath = Environment.ProcessPath! },
    "Synthetic cancellation regression", new string('x', 4 * 1024 * 1024), cancellationToken: cancellation.Token);
try
{
    var ready = Stopwatch.StartNew();
    while (!File.Exists(Path.Combine(evidence, "descendant.pid")))
    {
        if (ready.Elapsed > TimeSpan.FromSeconds(10)) throw new Exception("Fixture did not start.");
        await Task.Delay(25);
    }
    cancellation.Cancel();
    try
    {
        await operation.WaitAsync(TimeSpan.FromSeconds(8));
        throw new Exception("Cancellation was ignored.");
    }
    catch (OperationCanceledException) when (cancellation.IsCancellationRequested) { }
    foreach (var name in new[] { "parent", "descendant" })
    {
        var pid = int.Parse(File.ReadAllText(Path.Combine(evidence, name + ".pid")));
        try
        {
            using var process = Process.GetProcessById(pid);
            await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(3));
        }
        catch (ArgumentException) { }
    }
    Console.WriteLine("PASS: blocked multi-megabyte stdin write cancels promptly and owned process tree exits");
}
finally
{
    // On a red test, terminate only the fixture PIDs recorded by this unique run.
    foreach (var file in Directory.GetFiles(evidence, "*.pid"))
    {
        try { using var process = Process.GetProcessById(int.Parse(File.ReadAllText(file))); if (!process.HasExited) process.Kill(true); }
        catch (ArgumentException) { }
    }
    try { await operation.WaitAsync(TimeSpan.FromSeconds(3)); } catch { }
}
