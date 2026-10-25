using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Maily.OutlookAddin.Web.Services;

var services = new ServiceCollection().AddLogging().BuildServiceProvider();
var builder = new ApplicationBuilder(services);
builder.UseMailyApiKey("synthetic-test-key");
builder.Run(context => { context.Response.StatusCode = 204; return Task.CompletedTask; });
var pipeline = builder.Build();
var failures = 0;
foreach (var test in new[] {
    ("GET", "/addin/taskpane.html", "", 204),
    ("HEAD", "/addin/assets/icon.png", "", 204),
    ("GET", "/api/health", "", 401),
    ("GET", "/API/health", "", 401),
    ("POST", "/api/outlook/rewrite-active-compose", "", 401),
    ("GET", "/api/health", "wrong", 401),
    ("GET", "/api/health", "synthetic-test-key", 204),
    ("POST", "/addin/taskpane.html", "", 401),
    ("GET", "/addin-other/taskpane.html", "", 401),
    ("GET", "/addin/../api/health", "", 401),
    ("GET", "/addin/%2e%2e/api/health", "", 401),
    ("GET", "/unknown", "", 401)
}) {
    var context = new DefaultHttpContext { RequestServices = services };
    context.Response.Body = new MemoryStream();
    context.Request.Method = test.Item1;
    context.Request.Path = test.Item2;
    context.Request.Headers["X-Maily-Api-Key"] = test.Item3;
    await pipeline(context);
    if (context.Response.StatusCode != test.Item4) { failures++; Console.WriteLine($"FAIL {test.Item1} {test.Item2}: {context.Response.StatusCode}, expected {test.Item4}"); }
}
Console.WriteLine($"RESULT: {12-failures} passed, {failures} failed");
var browserBuilder = new ApplicationBuilder(services);
browserBuilder.UseMailyBrowserBoundary();
browserBuilder.Run(context => { context.Response.StatusCode = 204; return Task.CompletedTask; });
var browserPipeline = browserBuilder.Build();
foreach (var test in new[] {
    ("localhost:7097", "https://localhost:7097", 204),
    ("localhost:443", "https://localhost", 204),
    ("localhost:7097", "", 204),
    ("localhost:7097", "https://attacker.invalid", 403),
    ("localhost:7097", "null", 403),
    ("localhost:7097", "https://localhost:7098", 403),
    ("attacker.invalid:7097", "https://attacker.invalid:7097", 403)
}) {
    var context = new DefaultHttpContext { RequestServices = services };
    context.Response.Body = new MemoryStream();
    context.Request.Method = "POST";
    context.Request.Scheme = "https";
    context.Request.Host = new HostString(test.Item1);
    context.Request.Path = "/api/outlook/generate-reply";
    if (test.Item2.Length > 0) context.Request.Headers.Origin = test.Item2;
    await browserPipeline(context);
    if (context.Response.StatusCode != test.Item3) { failures++; Console.WriteLine("FAIL browser-origin boundary"); }
}
Console.WriteLine($"Browser boundary: {failures} total failures");
return failures == 0 ? 0 : 1;
