using Maily.Generation.Services;
using Maily.Generation.Models;
using Microsoft.Extensions.Logging;
const string canary = "SYNTHETIC_PRIVATE_DRAFT_canary@example.invalid";
if (args.Contains("exec")) {
 await Console.In.ReadToEndAsync();
 Console.Error.WriteLine(canary);
 var mode = Environment.GetEnvironmentVariable("MAILY_PRIVACY_TEST_MODE");
 if (mode == "success") Console.WriteLine("Synthetic reply");
 return mode == "failure" ? 1 : 0;
}
foreach (var mode in new[] {"success", "failure", "empty"}) {
 Environment.SetEnvironmentVariable("MAILY_PRIVACY_TEST_MODE", mode);
 var log = new Capture();
 var gateway = new OracleResponsesGateway(log);
 string error = "";
 try { await gateway.GenerateReplyAsync(new CodexRuntimeConfig { CodexExecutablePath = Environment.ProcessPath! }, "synthetic", "synthetic", "test-request"); }
 catch(Exception ex) { error = ex.ToString(); }
 if (log.Text.Contains(canary) || error.Contains(canary)) throw new Exception($"FAIL {mode}: private provider diagnostics exposed");
 if (!log.Text.Contains("test-request")) throw new Exception("Missing correlation ID");
 Console.WriteLine($"PASS provider diagnostics {mode}");
}
var readinessCases = new[] {
 ("ChatGPT stdout", true, 0, true, 0, "Logged in using ChatGPT", "", true),
 ("ChatGPT stderr", true, 0, true, 0, "", "Logged in using ChatGPT\r\n", true),
 ("API key", true, 0, true, 0, "", "Logged in using an API key", false),
 ("failed login", true, 0, true, 1, "Logged in using ChatGPT", "", false),
 ("failed version", true, 1, true, 0, "Logged in using ChatGPT", "", false),
 ("login not started", true, 0, false, 0, "Logged in using ChatGPT", "", false),
 ("version not started", false, 0, true, 0, "Logged in using ChatGPT", "", false),
 ("provider diagnostic", true, 0, true, 0, "Logged in using ChatGPT", canary, false),
 ("ambiguous output", true, 0, true, 0, "Not Logged in using ChatGPT", "", false),
 ("unknown output", true, 0, true, 0, canary, "", false)
};
foreach (var (name, versionStarted, versionExit, loginStarted, loginExit, stdout, stderr, expected) in readinessCases) {
 var status = CodexReadiness.Evaluate(versionStarted, versionExit, loginStarted, loginExit, stdout, stderr);
 if (status.Ok != expected) throw new Exception($"FAIL readiness {name}");
 if (System.Text.Json.JsonSerializer.Serialize(status).Contains(canary)) throw new Exception("FAIL readiness exposes provider output");
 if (!expected && !status.RecommendedAction.Contains("MailyOutlookAddinSetup.exe")) throw new Exception("FAIL guided setup action");
 Console.WriteLine($"PASS readiness {name}");
}
return 0;
sealed class Capture : ILogger<OracleResponsesGateway> {
 public string Text = "";
 public IDisposable? BeginScope<TState>(TState state) where TState:notnull => null;
 public bool IsEnabled(LogLevel level) => true;
 public void Log<TState>(LogLevel level, EventId id, TState state, Exception? error, Func<TState,Exception?,string> formatter) => Text += formatter(state,error) + error;
}
