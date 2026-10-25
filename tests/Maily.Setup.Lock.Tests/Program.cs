using System.Diagnostics;
using Maily.OutlookAddin.Setup;

if (args.Length == 3 && args[0] == "--attempt")
{
    using var lease = SetupInstanceLease.TryAcquire(args[1]);
    if (lease is null) return 17;
    File.WriteAllText(args[2], "synthetic protected operation ran");
    return 0;
}

var output = Path.GetFullPath(args.Length == 1 ? args[0] : Path.Combine("output", "setup-lock-" + Guid.NewGuid().ToString("N")));
Directory.CreateDirectory(output);
var name = @"Global\Maily.Setup.Tests." + Guid.NewGuid().ToString("N");
void Check(bool condition, string description) { if (!condition) throw new Exception(description); }
int Attempt(string mutexName, string marker)
{
    var executable = Environment.ProcessPath ?? throw new Exception("Missing process path");
    var start = new ProcessStartInfo(executable) { UseShellExecute = false, CreateNoWindow = true };
    if (Path.GetFileNameWithoutExtension(executable).Equals("dotnet", StringComparison.OrdinalIgnoreCase)) start.ArgumentList.Add(typeof(SetupInstanceLease).Assembly.Location);
    start.ArgumentList.Add("--attempt"); start.ArgumentList.Add(mutexName); start.ArgumentList.Add(marker);
    using var process = Process.Start(start) ?? throw new Exception("Could not start synthetic contender");
    if (!process.WaitForExit(15000)) { process.Kill(true); throw new Exception("Owned synthetic contender timed out"); }
    return process.ExitCode;
}
using (var owner = SetupInstanceLease.TryAcquire(name))
{
    Check(owner is not null, "First setup must acquire ownership");
    var marker = Path.Combine(output, "blocked-operation.txt");
    Check(Attempt(name, marker) == 17, "Concurrent installer/uninstaller must be rejected before protected work");
    Check(!File.Exists(marker), "Rejected contender must not mutate its fixture");
    Check(Attempt(name + ".other-user", Path.Combine(output, "independent-user.txt")) == 0, "Independent user lock must not block");
}
Check(Attempt(name, Path.Combine(output, "after-release.txt")) == 0, "Next setup must run after normal disposal");
var abandoned = name + ".abandoned";
using var abandonedObserver = new Mutex(false, abandoned);
var thread = new Thread(() => { var owner = SetupInstanceLease.TryAcquire(abandoned); Check(owner is not null, "Abandonment fixture ownership"); });
thread.Start(); thread.Join();
using (var recovered = SetupInstanceLease.TryAcquire(abandoned)) Check(recovered is not null, "Abandoned ownership must be recoverable");
Check(Attempt(abandoned, Path.Combine(output, "after-abandoned-release.txt")) == 0, "Recovered ownership must release correctly");
Console.WriteLine("RESULT: 5 installer concurrency checks passed; only synthetic child processes and fixture files used");
Console.WriteLine("Artifacts: " + output);
return 0;
