using Maily.Generation.Models;

namespace Maily.Generation.Services;

public interface ICodexModelListSource
{
    Task<IReadOnlyList<CodexModelInfo>> ReadAsync(CodexRuntimeConfig config, CancellationToken cancellationToken);
}

public sealed class CodexModelCatalog(ICodexModelListSource source)
{
    private readonly SemaphoreSlim gate = new(1, 1);
    private CodexModelCatalogSnapshot? cached;
    private string? cacheKey;
    private DateTimeOffset expires;
    private static readonly string[] Efforts = ["none", "minimal", "low", "medium", "high", "xhigh", "max", "ultra"];
    private static readonly string[] EfficientModels = ["gpt-6-luna", "gpt-5.6-luna"];

    public async Task<CodexModelCatalogSnapshot> GetAsync(CodexRuntimeConfig config, CancellationToken cancellationToken = default)
    {
        await gate.WaitAsync(cancellationToken);
        try
        {
            var key = config.CodexExecutablePath + "\n" + config.CodexHome;
            if (cached is not null && key == cacheKey && DateTimeOffset.UtcNow < expires) return cached;
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
                cached = new("ready", "codex-model-list", models, recommended.Id,
                    efficient is not null ? "efficient-model-preference" : "provider-default",
                    DateTimeOffset.UtcNow, "Availability can change with your ChatGPT plan and runtime.");
                expires = DateTimeOffset.UtcNow.AddMinutes(5);
            }
            catch (Exception) when (!cancellationToken.IsCancellationRequested)
            {
                cached = new("unavailable", "codex-model-list", [], null, "runtime-default",
                    DateTimeOffset.UtcNow, "Models could not be loaded. Automatic will use the runtime default. Retry shortly.");
                expires = DateTimeOffset.UtcNow.AddSeconds(30);
            }
            cacheKey = key;
            return cached!;
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
