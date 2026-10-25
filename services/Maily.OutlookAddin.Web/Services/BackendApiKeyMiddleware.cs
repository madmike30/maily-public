using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;

namespace Maily.OutlookAddin.Web.Services;

public static class BackendApiKeyMiddleware
{
    public static IApplicationBuilder UseMailyApiKey(this IApplicationBuilder app, string? backendApiKey)
    {
        return app.Use(async (context, next) =>
        {
            if (!string.IsNullOrWhiteSpace(backendApiKey) && !IsStaticBootstrap(context.Request))
            {
                var provided = context.Request.Headers["X-Maily-Api-Key"].ToString();
                if (!string.Equals(provided, backendApiKey, StringComparison.Ordinal))
                {
                    context.Response.StatusCode = StatusCodes.Status401Unauthorized;
                    await context.Response.WriteAsJsonAsync(new { error = "Unauthorized." });
                    return;
                }
            }
            await next();
        });
    }

    private static bool IsStaticBootstrap(HttpRequest request)
    {
        if (!HttpMethods.IsGet(request.Method) && !HttpMethods.IsHead(request.Method)) return false;
        if (!request.Path.StartsWithSegments("/addin", StringComparison.OrdinalIgnoreCase)) return false;
        var path = request.Path.Value ?? string.Empty;
        // Refuse encoded or traversing paths even when called before server normalization.
        if (path.Contains('%') || path.Contains('\\') || path.Split('/').Any(segment => segment is "." or "..")) return false;
        return Path.GetExtension(path).ToLowerInvariant() is ".html" or ".js" or ".css" or ".png";
    }
}
