using Maily.Generation.Models;
using Maily.Generation.Services;
using System.Text.Json;
using Maily.Generation.Contracts;
using Microsoft.Extensions.Logging.Abstractions;

if (args.Contains("exec"))
{
    await Console.In.ReadToEndAsync();
    File.AppendAllText(Environment.GetEnvironmentVariable("MAILY_SYNTHETIC_COUNTER")!, "call\n");
    Console.WriteLine("Yes please do it.");
    return;
}

if (args.Contains("app-server"))
{
    var mode = Environment.GetEnvironmentVariable("MAILY_CATALOG_TEST_MODE");
    if (mode == "hang") { await Task.Delay(30000); return; }
    if (mode == "oversized") { Console.WriteLine(new string('x', 262145)); return; }
    if (mode == "malformed") { Console.WriteLine("not json PRIVATE"); return; }
    while (await Console.In.ReadLineAsync() is { } line)
    {
        using var rpcRequest = JsonDocument.Parse(line);
        var root = rpcRequest.RootElement;
        if (!root.TryGetProperty("id", out var id)) continue;
        if (root.GetProperty("method").GetString() == "initialize")
            Console.WriteLine(JsonSerializer.Serialize(new { id = id.GetInt32(), result = new { } }));
        else
        {
            var cursor = root.GetProperty("params").GetProperty("cursor");
            var first = cursor.ValueKind == JsonValueKind.Null;
            Console.WriteLine(JsonSerializer.Serialize(new { id = id.GetInt32(), result = new { data = new[] { new { model = first ? "provider-default" : "gpt-6-luna", displayName = "Synthetic", description = "Synthetic only", hidden = false, isDefault = first, inputModalities = new[] { "text" }, supportedReasoningEfforts = new[] { new { reasoningEffort = "low" } } } }, nextCursor = first || mode == "repeat" ? "page2" : null } }));
        }
    }
    return;
}
var config = new CodexRuntimeConfig { CodexExecutablePath = Environment.ProcessPath! };
var models = await new CodexModelListSource().ReadAsync(config, CancellationToken.None);
Check(models.Count == 2, "stdio handshake and pagination");
foreach (var mode in new[] { "repeat", "oversized", "malformed", "hang" })
{
    Environment.SetEnvironmentVariable("MAILY_CATALOG_TEST_MODE", mode);
    using var cancel = new CancellationTokenSource(mode == "hang" ? 150 : 5000);
    var failed = false;
    try { await new CodexModelListSource().ReadAsync(config, cancel.Token); }
    catch { failed = true; }
    Check(failed, "reject bounded protocol " + mode);
}
Environment.SetEnvironmentVariable("MAILY_CATALOG_TEST_MODE", null);
var source = new FakeSource(models);
var catalog = new CodexModelCatalog(source);
var snapshots = await Task.WhenAll(Enumerable.Range(0, 8).Select(_ => catalog.GetAsync(config)));
Check(source.Calls == 1, "single flight and successful cache");
Check(snapshots[0].RecommendedModel == "gpt-6-luna", "efficient model intersects catalog");
CodexModelCatalog.ApplySelection(config, snapshots[0]);
Check(config.Model == "gpt-6-luna" && config.ReasoningEffort == "low", "automatic low effort selection");
config.Model = "saved-unavailable";
config.ReasoningEffort = "";
CodexModelCatalog.ApplySelection(config, snapshots[0]);
Check(config.Model == "saved-unavailable" && config.ReasoningEffort == "", "explicit unavailable choice preserved");
var fallback = await new CodexModelCatalog(new FakeSource([models[0]])).GetAsync(config);
Check(fallback.RecommendedModel == "provider-default", "absent preferred model uses provider default");
var olderFastCatalog = await new CodexModelCatalog(new FakeSource([
    new("gpt-5.6-sol", "Workhorse", "", ["low"], true),
    new("gpt-5.6-luna", "Fast", "", ["low"], false)
])).GetAsync(config);
Check(olderFastCatalog.RecommendedModel == "gpt-5.6-luna", "available older efficient model beats workhorse default");
var failing = new FakeSource(models) { Fail = true };
var unavailable = new CodexModelCatalog(failing);
await unavailable.GetAsync(config);
var failure = await unavailable.GetAsync(config);
Check(failing.Calls == 1 && failure.Status == "unavailable" && failure.Models.Count == 0, "failure backoff and honest empty catalog");
Check(!JsonSerializer.Serialize(failure).Contains("PRIVATE"), "failure details redacted");
var testRoot = Path.Combine(Directory.GetCurrentDirectory(), "output", "single-rewrite-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(testRoot);
var counter = Path.Combine(testRoot, "calls.txt");
Environment.SetEnvironmentVariable("MAILY_SYNTHETIC_COUNTER", counter);
var options = new GenerationOptions { CodexExecutablePath = Environment.ProcessPath!, CodexHome = testRoot, SettingsFilePath = Path.Combine(testRoot, "settings.json") };
var paths = new PathResolver(options);
var service = new OutlookReplyGenerationService(new CodexRuntimeConfigLoader(paths, options), new PromptTemplateEngine(paths),
    new OracleResponsesGateway(NullLogger<OracleResponsesGateway>.Instance), new ProfileSettingsStore(paths), options);
var request = new GenerateReplyRequest { Client = new GenerateReplyClientContext { MailboxAddress = "synthetic@example.invalid" } };
request.Operation.CurrentDraft = "Yes please do it.";
var reply = await service.GenerateReplyAsync(request, "synthetic-single-call");
Check(reply.Draft.Text == "Yes please do it." && File.ReadAllLines(counter).Length == 1, "valid short rewrite invokes provider once");
void Check(bool ok, string name) { if (!ok) throw new Exception("FAIL " + name); Console.WriteLine("PASS " + name); }
sealed class FakeSource(IReadOnlyList<CodexModelInfo> models) : ICodexModelListSource
{
    public int Calls; public bool Fail;
    public async Task<IReadOnlyList<CodexModelInfo>> ReadAsync(CodexRuntimeConfig config, CancellationToken token)
    { Interlocked.Increment(ref Calls); await Task.Delay(20, token); if (Fail) throw new Exception("PRIVATE"); return models; }
}
