using Maily.Generation.Services;
using Microsoft.Extensions.Hosting;

namespace Maily.OutlookAddin.Web.Services;

/// <summary>Discover supported models in parallel with startup; never invokes inference.</summary>
public sealed class CodexCatalogWarmup(CodexModelCatalog catalog, CodexRuntimeConfigLoader configLoader) : BackgroundService
{
    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        try { await catalog.GetAsync(configLoader.Load(), stoppingToken); }
        catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested) { }
        catch (Exception)
        {
            // Readiness and the model endpoint provide the normal recovery path.
            // Optional warmup must neither stop the host nor expose runtime details.
        }
    }
}
