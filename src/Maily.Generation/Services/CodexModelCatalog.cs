using Maily.Generation.Models;

namespace Maily.Generation.Services;

public interface ICodexModelListSource
{
    Task<IReadOnlyList<CodexModelInfo>> ReadAsync(CodexRuntimeConfig config, CancellationToken cancellationToken);
}

public sealed class CodexModelCatalog(ICodexModelListSource source, TimeProvider? timeProvider = null) : IAsyncDisposable
{
    private readonly SemaphoreSlim gate = new(1, 1);
    private readonly object lifetimeLock = new();
    private readonly CancellationTokenSource lifetime = new();
    private readonly HashSet<Task> operations = [];
    private bool disposed;
    private Task? disposal;
    private readonly TimeProvider clock = timeProvider ?? TimeProvider.System;
    private CacheEntry? cache;
    private int backgroundRefresh;
    private sealed record CacheEntry(string Key, CodexModelCatalogSnapshot Snapshot, DateTimeOffset Expires);
    private static readonly string[] Efforts = ["none", "minimal", "low", "medium", "high", "xhigh", "max", "ultra"];
    private static readonly string[] EfficientModels = ["gpt-6-luna", "gpt-5.6-luna"];

    public Task<CodexModelCatalogSnapshot> GetAsync(CodexRuntimeConfig config, CancellationToken cancellationToken = default)
    {
        lock (lifetimeLock)
        {
            ObjectDisposedException.ThrowIf(disposed, this);
            return Track(GetWithLifetimeAsync(config, cancellationToken));
        }
    }

    private async Task<CodexModelCatalogSnapshot> GetWithLifetimeAsync(CodexRuntimeConfig config, CancellationToken cancellationToken)
    {
        using var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, lifetime.Token);
        cancellationToken = linked.Token;
        cancellationToken.ThrowIfCancellationRequested();
        var entry = Volatile.Read(ref cache);
        var key = config.CodexExecutablePath + "\n" + config.CodexHome;
        var now = clock.GetUtcNow();
        // Briefly reuse a successful catalog while a single bounded refresh runs.
        // Never reuse another connection, a failed lookup, or indefinitely stale data.
        if (entry is not null && entry.Key == key && entry.Snapshot.Status == "ready" &&
            now >= entry.Expires && now < entry.Expires.AddSeconds(30))
        {
            if (Interlocked.CompareExchange(ref backgroundRefresh, 1, 0) == 0)
            {
                var connection = new CodexRuntimeConfig { CodexExecutablePath = config.CodexExecutablePath, CodexHome = config.CodexHome };
                _ = Track(RefreshInBackgroundAsync(connection));
            }
            return entry.Snapshot;
        }
        return await GetCoreAsync(config, cancellationToken);
    }

    private async Task RefreshInBackgroundAsync(CodexRuntimeConfig config)
    {
        try { await GetCoreAsync(config, lifetime.Token); }
        catch (OperationCanceledException) when (lifetime.IsCancellationRequested) { }
        finally { Interlocked.Exchange(ref backgroundRefresh, 0); }
    }

    // Called under lifetimeLock while starting work; disposal takes its snapshot
    // under the same lock, so no refresh can escape shutdown tracking.
    private T Track<T>(T operation) where T : Task
    {
        operations.Add(operation);
        _ = RemoveCompletedAsync(operation);
        return operation;
    }

    private async Task RemoveCompletedAsync(Task operation)
    {
        try { await operation; }
        catch { /* Callers still observe their own failure; background tasks are observed here. */ }
        finally { lock (lifetimeLock) operations.Remove(operation); }
    }

    public ValueTask DisposeAsync()
    {
        lock (lifetimeLock)
        {
            if (disposal is not null) return new(disposal);
            disposed = true;
            disposal = DisposeCoreAsync(operations.ToArray());
            return new(disposal);
        }
    }

    private async Task DisposeCoreAsync(Task[] pending)
    {
        // Do not run cancellation callbacks while holding the lifetime lock.
        await Task.Yield();
        lifetime.Cancel();
        try { await Task.WhenAll(pending); }
        catch { /* Normal caller cancellation must not prevent resource cleanup. */ }
        gate.Dispose();
        lifetime.Dispose();
    }

    private async Task<CodexModelCatalogSnapshot> GetCoreAsync(CodexRuntimeConfig config, CancellationToken cancellationToken)
    {
        await gate.WaitAsync(cancellationToken);
        try
        {
            var key = config.CodexExecutablePath + "\n" + config.CodexHome;
            var entry = Volatile.Read(ref cache);
            if (entry is not null && entry.Key == key && clock.GetUtcNow() < entry.Expires) return entry.Snapshot;
            CodexModelCatalogSnapshot snapshot;
            DateTimeOffset expires;
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeout.CancelAfter(TimeSpan.FromSeconds(5));
            try
            {
                var models = await source.ReadAsync(config, timeout.Token);
                if (models.Count == 0) throw new InvalidOperationException();
                // Documented efficient model preference, intersected with the runtime's catalog.
                // Catalog membership is not a guarantee of account entitlement or measured latency.
                var efficient = EfficientModels.Select(id => models.FirstOrDefault(m => m.Id == id)).FirstOrDefault(m => m is not null);
                var recommended = efficient ?? models.FirstOrDefault(m => m.IsDefault) ?? models[0];
                snapshot = new("ready", "codex-model-list", models, recommended.Id,
                    efficient is not null ? "efficient-model-preference" : "provider-default",
                    clock.GetUtcNow(), "Availability can change with your ChatGPT plan and runtime.");
                expires = clock.GetUtcNow().AddMinutes(5);
            }
            catch (Exception) when (!cancellationToken.IsCancellationRequested)
            {
                snapshot = new("unavailable", "codex-model-list", [], null, "runtime-default",
                    clock.GetUtcNow(), "Models could not be loaded. Automatic will use the runtime default. Retry shortly.");
                expires = clock.GetUtcNow().AddSeconds(30);
            }
            Volatile.Write(ref cache, new(key, snapshot, expires));
            return snapshot;
        }
        finally { gate.Release(); }
    }

    public static void ApplySelection(CodexRuntimeConfig config, CodexModelCatalogSnapshot snapshot)
    {
        if (string.IsNullOrWhiteSpace(config.Model)) config.Model = snapshot.RecommendedModel ?? string.Empty;
        var model = snapshot.Models.FirstOrDefault(m => m.Id == config.Model);
        config.ReasoningEffort = model is null ? string.Empty : Efforts.FirstOrDefault(e => model.ReasoningEfforts.Contains(e)) ?? string.Empty;
    }
}
