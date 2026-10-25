using System.Diagnostics;
using System.Reflection;
using Maily.OutlookAddin.Web.Services;



ProcessStartInfo Worker(params string[] arguments)
{
    var start = new ProcessStartInfo(Environment.ProcessPath!)
    {
        UseShellExecute = false, CreateNoWindow = true,
        RedirectStandardOutput = true, RedirectStandardError = true
    };
    if (Path.GetFileNameWithoutExtension(Environment.ProcessPath!).Equals("dotnet", StringComparison.OrdinalIgnoreCase))
        start.ArgumentList.Add(Assembly.GetExecutingAssembly().Location);
    foreach (var argument in arguments) start.ArgumentList.Add(argument);
    return start;
}

if (args.Length > 0 && args[0] == "--worker")
{
    if (args[1] == "success") { Console.Write("synthetic-output"); return 0; }
    if (args[1] == "fail") { Console.Error.Write("SYNTHETIC-PRIVATE-STDERR-CANARY"); return 7; }
    if (args[1] == "sleep")
    {
        File.WriteAllText(args[2], Environment.ProcessId.ToString());
        await Task.Delay(Timeout.Infinite);
    }
    if (args[1] == "tree")
    {
        using var child = Process.Start(Worker("--worker", "sleep", args[2] + ".child"))!;
        while (!File.Exists(args[2] + ".child")) await Task.Delay(10);
        File.WriteAllText(args[2], Environment.ProcessId.ToString());
        await Task.Delay(Timeout.Infinite);
    }
    return 0;
}

var output = Path.GetFullPath(args.FirstOrDefault() ?? Path.Combine("output", "bridge-process-tests", DateTime.UtcNow.ToString("yyyyMMdd-HHmmssfff")));
Directory.CreateDirectory(output);
using var log = new StreamWriter(Path.Combine(output, "results.log")) { AutoFlush = true };
var failures = 0;
async Task Check(string name, Func<Task> action)
{
    try { await action(); Console.WriteLine($"PASS {name}"); log.WriteLine($"PASS {name}"); }
    catch (Exception ex) { failures++; Console.WriteLine($"FAIL {name}: {ex}"); log.WriteLine($"FAIL {name}: {ex}"); }
}
void Assert(bool condition, string message) { if (!condition) throw new Exception(message); }
bool Alive(int id) { try { using var p = Process.GetProcessById(id); return !p.HasExited; } catch (ArgumentException) { return false; } }
async Task Ready(string path)
{
    using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(10));
    while (!File.Exists(path)) await Task.Delay(10, deadline.Token);
}
async Task VerifyCancellation(bool timeout)
{
    using var gate = new SemaphoreSlim(1, 1);
    using var cancel = new CancellationTokenSource();
    var marker = Path.Combine(output, timeout ? "timeout.pid" : "cancel.pid");
    var run = BridgeProcessRunner.RunAsync(Worker("--worker", "tree", marker), gate,
        TimeSpan.FromSeconds(timeout ? 4 : 30), cancel.Token);
    await Ready(marker);
    var parent = int.Parse(File.ReadAllText(marker));
    var child = int.Parse(File.ReadAllText(marker + ".child"));
    try
    {
        Assert(gate.CurrentCount == 0, "Worker does not own gate while running");
        var follower = Task.Run(async () => {
            await gate.WaitAsync();
            try { Assert(!Alive(parent) && !Alive(child), "Next request overlapped a surviving worker"); }
            finally { gate.Release(); }
        });
        if (!timeout) cancel.Cancel();
        Exception? failure = null;
        try { await run.WaitAsync(TimeSpan.FromSeconds(7)); } catch (Exception ex) { failure = ex; }
        Assert(timeout ? failure is TimeoutException : failure is OperationCanceledException, "Wrong cancellation category: " + failure);
        Assert(!Alive(parent) && !Alive(child), "Owned worker survived cancellation/deadline");
        await follower.WaitAsync(TimeSpan.FromSeconds(5));
        Assert(gate.CurrentCount == 1, "Gate was not released after child termination");
        var next = await BridgeProcessRunner.RunAsync(Worker("--worker", "success"), gate, TimeSpan.FromSeconds(5), CancellationToken.None);
        Assert(next == "synthetic-output", "Next request did not succeed");
    }
    finally
    {
        cancel.Cancel(); // Test owns these synthetic workers. Preserve all files and diagnostics.
        foreach (var id in new[] { parent, child })
            if (Alive(id)) { using var p = Process.GetProcessById(id); p.Kill(true); await p.WaitForExitAsync(); }
    }
}
await Check("successful output and nonzero exit", async () =>
{
    using var gate = new SemaphoreSlim(1, 1);
    Assert(await BridgeProcessRunner.RunAsync(Worker("--worker", "success"), gate, TimeSpan.FromSeconds(5), CancellationToken.None) == "synthetic-output", "Output lost");
    try { await BridgeProcessRunner.RunAsync(Worker("--worker", "fail"), gate, TimeSpan.FromSeconds(5), CancellationToken.None); throw new Exception("Expected nonzero exit failure"); }
    catch (InvalidOperationException ex) { Assert(ex.Message == "Outlook bridge action failed. Retry the action.", "Unsafe or unexpected bridge error: " + ex.GetType().Name); Assert(!ex.ToString().Contains("SYNTHETIC-PRIVATE-STDERR-CANARY"), "Process diagnostics leaked to caller"); }
    Assert(gate.CurrentCount == 1, "Failure leaked gate");
});
await Check("request cancellation terminates owned process tree before next call", () => VerifyCancellation(false));
await Check("deadline terminates owned process tree before next call", () => VerifyCancellation(true));
await Check("queued cancellation does not release another request's gate", async () =>
{
    using var gate = new SemaphoreSlim(0, 1);
    using var cancel = new CancellationTokenSource(100);
    try { await BridgeProcessRunner.RunAsync(Worker("--worker", "success"), gate, TimeSpan.FromSeconds(5), cancel.Token); throw new Exception("Expected cancellation"); }
    catch (OperationCanceledException) { }
    Assert(gate.CurrentCount == 0, "Released unowned gate");
});
await Check("start failure releases acquired gate", async () =>
{
    using var gate = new SemaphoreSlim(1, 1);
    var bad = Worker(); bad.FileName = Path.Combine(output, "missing-worker.exe");
    try { await BridgeProcessRunner.RunAsync(bad, gate, TimeSpan.FromSeconds(5), CancellationToken.None); throw new Exception("Expected start failure"); }
    catch (System.ComponentModel.Win32Exception) { }
    Assert(gate.CurrentCount == 1, "Start failure leaked gate");
});
Console.WriteLine($"RESULT failures={failures}");
log.WriteLine($"RESULT failures={failures}");
return failures == 0 ? 0 : 1;