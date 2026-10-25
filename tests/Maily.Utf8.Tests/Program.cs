using System.Reflection;
using System.Text;
using Maily.Generation.Models;
using Maily.Generation.Services;
using Microsoft.Extensions.Logging.Abstractions;

const string unicode = "שלום café 日本語 😀";
if (args.FirstOrDefault() == "exec")
{
    // Use raw streams so BOMs and encoding corruption cannot be hidden by a reader.
    using var bytes = new MemoryStream();
    Console.OpenStandardInput().CopyTo(bytes);
    var input = new UTF8Encoding(false, true).GetString(bytes.ToArray());
    if (input.StartsWith('\uFEFF') || !input.Contains(unicode)) Environment.Exit(41);
    Console.OpenStandardOutput().Write(Encoding.UTF8.GetBytes(unicode));
    return;
}
if (args.FirstOrDefault() == "--version")
{
    Console.OpenStandardOutput().Write(Encoding.UTF8.GetBytes(unicode));
    Console.OpenStandardError().Write(Encoding.UTF8.GetBytes(unicode));
    return;
}
Console.InputEncoding = Encoding.ASCII;
Console.OutputEncoding = Encoding.ASCII;
var config = new CodexRuntimeConfig { CodexExecutablePath = Environment.ProcessPath! };
var gateway = new OracleResponsesGateway(NullLogger<OracleResponsesGateway>.Instance);
var result = await gateway.GenerateReplyAsync(config, "Synthetic encoding test", unicode);
if (result != unicode) throw new Exception("Unicode response did not round-trip under ASCII host encoding.");
Console.WriteLine("PASS: BOM-free Unicode prompt and response under ASCII host encoding");
var probe = Assembly.Load("Maily.OutlookAddin.Web").GetTypes()
    .SelectMany(t => t.GetMethods(BindingFlags.Static | BindingFlags.NonPublic))
    .Single(m => m.Name.Contains("g__RunCodexProcess|") && m.GetParameters().Length == 2);
var probeResult = probe.Invoke(null, [config, new[] { "--version" }])!;
foreach (var property in new[] { "Stdout", "Stderr" })
    if ((string)probeResult.GetType().GetProperty(property)!.GetValue(probeResult)! != unicode)
        throw new Exception("Probe " + property + " was not decoded as UTF-8.");
Console.WriteLine("PASS: Unicode probe stdout and stderr under ASCII host encoding");
