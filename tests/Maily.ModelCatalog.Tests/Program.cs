using Maily.Generation.Models;
using Maily.Generation.Services;
using System.Text.Json;
using Maily.Generation.Contracts;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.DependencyInjection;

var subscriptionKeys = new[] { "OPENAI_API_KEY", "CODEX_API_KEY", "AZURE_OPENAI_API_KEY" };
if (Environment.GetEnvironmentVariable("MAILY_SUBSCRIPTION_ENV_TEST") == "1")
{
    if (subscriptionKeys.Any(key => Environment.GetEnvironmentVariable(key) is not null) ||
        Environment.GetEnvironmentVariable("CODEX_HOME") != "synthetic-chatgpt-home")
        throw new InvalidOperationException("Subscription child environment was not isolated.");
    if (args.Contains("exec")) { await Console.In.ReadToEndAsync(); Console.WriteLine("Synthetic subscription reply."); return; }
}
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
var previousKeys = subscriptionKeys.ToDictionary(key => key, Environment.GetEnvironmentVariable);
try
{
    foreach (var key in subscriptionKeys) Environment.SetEnvironmentVariable(key, "synthetic-not-a-credential");
    Environment.SetEnvironmentVariable("MAILY_SUBSCRIPTION_ENV_TEST", "1");
    var subscriptionConfig = new CodexRuntimeConfig { CodexExecutablePath = Environment.ProcessPath!, CodexHome = "synthetic-chatgpt-home" };
    var isolatedModels = await new CodexModelListSource().ReadAsync(subscriptionConfig, CancellationToken.None);
    Check(isolatedModels.Count == 2, "model discovery excludes API keys and preserves configured ChatGPT home");
    var isolatedReply = await new OracleResponsesGateway(NullLogger<OracleResponsesGateway>.Instance).GenerateReplyAsync(subscriptionConfig, "Synthetic", "Synthetic");
    Check(isolatedReply == "Synthetic subscription reply.", "generation excludes API keys and preserves configured ChatGPT home");
    Check(subscriptionKeys.All(key => Environment.GetEnvironmentVariable(key) == "synthetic-not-a-credential"), "child isolation leaves parent environment unchanged");
}
finally
{
    foreach (var item in previousKeys) Environment.SetEnvironmentVariable(item.Key, item.Value);
    Environment.SetEnvironmentVariable("MAILY_SUBSCRIPTION_ENV_TEST", null);
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
var clock = new ManualClock();
var delayed = new ControlledSource(models);
var warmingCatalog = new CodexModelCatalog(delayed, clock);
await warmingCatalog.GetAsync(config);
clock.Advance(TimeSpan.FromMinutes(5) + TimeSpan.FromSeconds(1));
delayed.Block = true;
var staleReads = Enumerable.Range(0, 8).Select(_ => warmingCatalog.GetAsync(config)).ToArray();
await Task.WhenAll(staleReads).WaitAsync(TimeSpan.FromSeconds(1));
Check(staleReads.All(t => t.Result.Status == "ready"), "recent successful catalog returns without waiting for refresh");
await delayed.Started.Task.WaitAsync(TimeSpan.FromSeconds(1));
Check(delayed.Calls == 2, "expired catalog refresh is single flight");
clock.Advance(TimeSpan.FromSeconds(30));
var tooOld = warmingCatalog.GetAsync(config);
await Task.Delay(30);
Check(!tooOld.IsCompleted, "catalog older than thirty-second grace must await refresh");
delayed.Release.TrySetResult();
await tooOld.WaitAsync(TimeSpan.FromSeconds(1));
Check(delayed.Calls == 2, "blocking reader reuses successful background refresh");
var changedHome = new CodexRuntimeConfig { CodexExecutablePath = config.CodexExecutablePath, CodexHome = "synthetic-other-home" };
await warmingCatalog.GetAsync(changedHome);
Check(delayed.Calls == 3, "different connection never receives cached prior catalog");
var errorClock = new ManualClock();
var failingRefreshSource = new FakeSource(models);
var failedRefreshCatalog = new CodexModelCatalog(failingRefreshSource, errorClock);
await failedRefreshCatalog.GetAsync(config);
failingRefreshSource.Fail = true;
errorClock.Advance(TimeSpan.FromMinutes(5) + TimeSpan.FromSeconds(1));
Check((await failedRefreshCatalog.GetAsync(config)).Status == "ready", "only recent success can bridge a refresh");
await Task.Delay(80);
Check((await failedRefreshCatalog.GetAsync(config)).Status == "unavailable" && failingRefreshSource.Calls == 2,
    "failed refresh becomes honest unavailable state with failure backoff");
using (var cancelledRead = new CancellationTokenSource())
{
    cancelledRead.Cancel();
    var cancelled = false;
    try { await warmingCatalog.GetAsync(changedHome, cancelledRead.Token); } catch (OperationCanceledException) { cancelled = true; }
    Check(cancelled, "cancelled requests do not start catalog refresh");
}
var startupSource = new ControlledSource(models) { Block = true };
var shutdownClock = new ManualClock();
var shutdownSource = new ControlledSource(models) { BlockCleanup = true };
var shutdownServices = new ServiceCollection();
shutdownServices.AddSingleton(_ => new CodexModelCatalog(shutdownSource, shutdownClock));
var shutdownProvider = shutdownServices.BuildServiceProvider();
var shutdownCatalog = shutdownProvider.GetRequiredService<CodexModelCatalog>();
await shutdownCatalog.GetAsync(config);
shutdownSource.Block = true;
shutdownClock.Advance(TimeSpan.FromMinutes(5) + TimeSpan.FromSeconds(1));
await shutdownCatalog.GetAsync(config);
await shutdownSource.Started.Task.WaitAsync(TimeSpan.FromSeconds(1));
shutdownClock.Advance(TimeSpan.FromSeconds(30));
var queuedBeforeShutdown = shutdownCatalog.GetAsync(config);
var shutdown = shutdownProvider.DisposeAsync().AsTask();
await Task.WhenAny(shutdownSource.CancellationObserved.Task, Task.Delay(500));
Check(shutdownSource.Cancelled, "DI shutdown cancels stale background refresh");
Check(!shutdown.IsCompleted, "DI shutdown awaits source child-process cleanup");
var queuedCancelled = false;
try { await queuedBeforeShutdown.WaitAsync(TimeSpan.FromSeconds(1)); } catch (OperationCanceledException) { queuedCancelled = true; }
Check(queuedCancelled, "shutdown cancels callers queued behind the stale refresh");
shutdownSource.CleanupRelease.TrySetResult();
await shutdown.WaitAsync(TimeSpan.FromSeconds(1));
Check(shutdownSource.CleanupFinished, "stale refresh cleanup completes before disposal returns");
var rejectedAfterShutdown = false;
try { await shutdownCatalog.GetAsync(config); } catch (ObjectDisposedException) { rejectedAfterShutdown = true; }
Check(rejectedAfterShutdown && shutdownSource.Calls == 2, "disposed catalog rejects new discovery");
await shutdownCatalog.DisposeAsync();
Check(shutdownSource.Calls == 2, "catalog disposal is idempotent");
var startupOptions = new GenerationOptions { CodexExecutablePath = Environment.ProcessPath! };
using (var warmup = new Maily.OutlookAddin.Web.Services.CodexCatalogWarmup(
    new CodexModelCatalog(startupSource), new CodexRuntimeConfigLoader(new PathResolver(startupOptions), startupOptions)))
{
    await warmup.StartAsync(CancellationToken.None).WaitAsync(TimeSpan.FromSeconds(1));
    await startupSource.Started.Task.WaitAsync(TimeSpan.FromSeconds(1));
    Check(startupSource.Calls == 1, "startup warms catalog without blocking backend availability");
    using var stopTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(1));
    await warmup.StopAsync(stopTimeout.Token);
    Check(startupSource.Cancelled, "startup catalog warmup cancels on shutdown");
}
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
sealed class ManualClock : TimeProvider
{
    private DateTimeOffset now = DateTimeOffset.UtcNow;
    public override DateTimeOffset GetUtcNow() => now;
    public void Advance(TimeSpan duration) => now += duration;
}
sealed class ControlledSource(IReadOnlyList<CodexModelInfo> models) : ICodexModelListSource
{
    public int Calls; public bool Block; public bool Cancelled; public bool BlockCleanup; public bool CleanupFinished;
    public TaskCompletionSource Started = new(TaskCreationOptions.RunContinuationsAsynchronously);
    public TaskCompletionSource Release = new(TaskCreationOptions.RunContinuationsAsynchronously);
    public TaskCompletionSource CancellationObserved = new(TaskCreationOptions.RunContinuationsAsynchronously);
    public TaskCompletionSource CleanupRelease = new(TaskCreationOptions.RunContinuationsAsynchronously);
    public async Task<IReadOnlyList<CodexModelInfo>> ReadAsync(CodexRuntimeConfig config, CancellationToken token)
    {
        Interlocked.Increment(ref Calls);
        if (Block)
        {
            Started.TrySetResult();
            try { await Release.Task.WaitAsync(token); }
            catch (OperationCanceledException)
            {
                Cancelled = true;
                CancellationObserved.TrySetResult();
                if (BlockCleanup) await CleanupRelease.Task;
                CleanupFinished = true;
                throw;
            }
        }
        return models;
    }
}
