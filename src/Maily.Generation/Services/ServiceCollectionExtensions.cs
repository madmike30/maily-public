using Microsoft.Extensions.DependencyInjection;
using Maily.Generation.Models;

namespace Maily.Generation.Services;

public static class ServiceCollectionExtensions
{
	public static IServiceCollection AddMailyGenerationCore(this IServiceCollection services, Action<GenerationOptions>? configure = null)
	{
		var options = new GenerationOptions();
		configure?.Invoke(options);

		services.AddSingleton(options);
		services.AddSingleton<PathResolver>();
		services.AddSingleton<PromptTemplateEngine>();
		services.AddSingleton<CodexRuntimeConfigLoader>();
		services.AddSingleton<ICodexModelListSource, CodexModelListSource>();
		services.AddSingleton<CodexModelCatalog>();
		services.AddSingleton<ProfileSettingsStore>();
		services.AddSingleton<OutlookSearchPreparationService>();
		services.AddSingleton<OracleResponsesGateway>();
		services.AddSingleton<OutlookReplyGenerationService>();
		services.AddSingleton<OutlookInsightsService>();
		return services;
	}
}
