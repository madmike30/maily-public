using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;

namespace Maily.OutlookAddin.Web.Services;

public static class BrowserBoundaryMiddleware
{
    public static IApplicationBuilder UseMailyBrowserBoundary(this IApplicationBuilder app, bool allowExternalHosts = false)
    {
        return app.Use(async (context, next) =>
        {
            var request = context.Request;
            var host = request.Host.Host.Trim('[', ']');
            // Local browser pages must not reach the unauthenticated loopback service
            // through an attacker-controlled hostname or cross-origin request.
            var localHost = host.Equals("localhost", StringComparison.OrdinalIgnoreCase) ||
                host == "127.0.0.1" || host == "::1";
            var allowed = allowExternalHosts || localHost;
            if (request.Headers.TryGetValue("Origin", out var origin))
            {
                allowed &= Uri.TryCreate(origin.ToString(), UriKind.Absolute, out var uri) &&
                    Uri.TryCreate($"{request.Scheme}://{request.Host}", UriKind.Absolute, out var target) &&
                    string.IsNullOrEmpty(uri.UserInfo) &&
                    string.Equals(uri.GetLeftPart(UriPartial.Authority),
                        target.GetLeftPart(UriPartial.Authority), StringComparison.OrdinalIgnoreCase);
            }
            context.Response.Headers["X-Content-Type-Options"] = "nosniff";
            context.Response.Headers["Referrer-Policy"] = "no-referrer";
            if (request.Path.StartsWithSegments("/api", StringComparison.OrdinalIgnoreCase))
                context.Response.Headers.CacheControl = "no-store";
            if (!allowed)
            {
                context.Response.StatusCode = StatusCodes.Status403Forbidden;
                await context.Response.WriteAsJsonAsync(new { error = "Request origin is not allowed." });
                return;
            }
            await next();
        });
    }
}
