using System.Reflection;
using System.Text.Json;
using Maily.Generation.Models;
using Maily.Generation.Services;
using Microsoft.Extensions.Logging.Abstractions;

if (args.FirstOrDefault() is "--fixture" or "exec")
{
    if (args.Contains("exec")) Console.In.ReadToEnd();
    Console.WriteLine(JsonSerializer.Serialize(args[0] == "--fixture" ? args.Skip(1) : args));
    return;
}
if (!OperatingSystem.IsWindows()) throw new PlatformNotSupportedException("Windows cmd behavior must run on Windows.");
var artifact = Path.Combine(Path.GetTempPath(), "Maily CLI quoting " + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(artifact);
foreach (var file in Directory.GetFiles(AppContext.BaseDirectory)) File.Copy(file, Path.Combine(artifact, Path.GetFileName(file)));
var shim = Path.Combine(artifact, "codex fixture.cmd");
File.WriteAllText(shim, "@echo off\r\n\"%~dp0Maily.CliQuoting.Tests.exe\" --fixture %*\r\n", new System.Text.UTF8Encoding(false));
Console.WriteLine("Artifacts: " + artifact);
var gateway = new OracleResponsesGateway(NullLogger<OracleResponsesGateway>.Instance);
var result = await gateway.GenerateReplyAsync(new CodexRuntimeConfig { CodexExecutablePath = shim, Model = "synthetic model" }, "synthetic", "synthetic");
var actual = JsonSerializer.Deserialize<string[]>(result)!;
AssertToolIsolation(actual);
string[] expected = ["exec", "--ephemeral", "--skip-git-repo-check", "--sandbox", "read-only", "--ignore-user-config", "--ignore-rules", "--disable", "plugins", "--disable", "apps", "--disable", "browser_use", "--disable", "browser_use_external", "--disable", "computer_use", "--disable", "tool_search", "--disable", "shell_snapshot",
    "--disable", "shell_tool", "--disable", "view_image", "--disable", "multi_agent", "--disable", "image_generation", "--disable", "hooks", "--disable", "workspace_dependencies", "--disable", "skill_search", "--disable", "memories", "--disable", "code_mode", "--disable", "code_mode_host", "--disable", "tool_suggest", "--enable", "skip_host_skill_discovery", "-c", "web_search=\"disabled\"", "-c", "tools.view_image=false", "-c", "project_doc_max_bytes=0",
    "-c", "approval_policy=\"never\"", "--color", "never", "-m", "synthetic model", "-"];
if (!actual.SequenceEqual(expected))
    throw new Exception("Generation arguments were not preserved: " + result);
Console.WriteLine("PASS: batch path with spaces preserves exact generation arguments");
foreach (var executable in new[] { shim, Path.Combine(artifact, "Maily.CliQuoting.Tests.exe") })
{
    // Direct executable fixture understands exec without the batch shim's --fixture marker.
    var defaultResult = await gateway.GenerateReplyAsync(new CodexRuntimeConfig { CodexExecutablePath = executable, Model = " " }, "synthetic", "synthetic");
    var defaultArgs = JsonSerializer.Deserialize<string[]>(defaultResult)!;
    AssertToolIsolation(defaultArgs);
    if (defaultArgs.Contains("-m") || defaultArgs.Contains("--model") || defaultArgs.Any(a => a.StartsWith("model_reasoning_effort=")))
        throw new Exception("Default generation overrides the runtime's recommended model or reasoning.");
}
Console.WriteLine("PASS: native and batch default model omit model and reasoning overrides");
foreach (var executable in new[] { shim, Path.Combine(artifact, "Maily.CliQuoting.Tests.exe") })
{
    var selectedResult = await gateway.GenerateReplyAsync(new CodexRuntimeConfig { CodexExecutablePath = executable, Model = "synthetic-model", ReasoningEffort = "low" }, "synthetic", "synthetic");
    var selectedArgs = JsonSerializer.Deserialize<string[]>(selectedResult)!;
    if (!selectedArgs.Contains("model_reasoning_effort=\"low\"")) throw new Exception("Selected reasoning effort was not preserved.");
    if (selectedArgs.Any(a => a.Contains("service_tier"))) throw new Exception("Paid priority tier must not be enabled.");
}
Console.WriteLine("PASS: native and batch preserve selected low reasoning without paid priority tier");
var web = Assembly.Load("Maily.OutlookAddin.Web");
var probe = web.GetTypes().SelectMany(t => t.GetMethods(BindingFlags.Static | BindingFlags.NonPublic))
    .Single(m => m.Name.Contains("g__RunCodexProcess|") && m.GetParameters().Length == 2);
var probeResult = probe.Invoke(null, [new CodexRuntimeConfig { CodexExecutablePath = shim }, new[] { "--version" }])!;
var output = (string)probeResult.GetType().GetProperty("Stdout")!.GetValue(probeResult)!;
if (JsonSerializer.Deserialize<string[]>(output) is not ["--version"]) throw new Exception("Probe arguments changed.");
Console.WriteLine("PASS: availability probe runs the same batch path");
foreach (var unsafeModel in new[] { "%PATH%", "!PATH!", "a&b", "a|b", "a<b", "a>b", "a^b", "a\rb", "a\nb" })
{
    try
    {
        await gateway.GenerateReplyAsync(new CodexRuntimeConfig { CodexExecutablePath = shim, Model = unsafeModel }, "synthetic", "synthetic");
        throw new Exception("Shell control syntax was accepted.");
    }
    catch (ArgumentException) { }
}
Console.WriteLine("PASS: shell expansion and control syntax rejected before process start");

static void AssertToolIsolation(string[] arguments)
{
    foreach (var feature in new[] { "shell_tool", "view_image", "multi_agent", "image_generation", "hooks", "workspace_dependencies", "skill_search", "memories", "code_mode", "code_mode_host", "tool_suggest" })
        if (!Enumerable.Range(0, arguments.Length - 1).Any(index => arguments[index] == "--disable" && arguments[index + 1] == feature))
            throw new Exception("Generation must disable tool capability: " + feature);
    foreach (var option in new[] { "web_search=\"disabled\"", "tools.view_image=false", "project_doc_max_bytes=0", "approval_policy=\"never\"" })
        if (!Enumerable.Range(0, arguments.Length - 1).Any(index => arguments[index] == "-c" && arguments[index + 1] == option))
            throw new Exception("Generation must enforce context/tool policy: " + option);
    if (!Enumerable.Range(0, arguments.Length - 1).Any(index => arguments[index] == "--enable" && arguments[index + 1] == "skip_host_skill_discovery"))
        throw new Exception("Generation must not discover private host skills.");
}
