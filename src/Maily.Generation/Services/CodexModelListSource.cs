using System.Diagnostics;
using System.Text;
using System.Text.Json;
using Maily.Generation.Models;

namespace Maily.Generation.Services;

/// <summary>Read-only supported app-server RPC. Never creates a thread or invokes inference.</summary>
public sealed class CodexModelListSource : ICodexModelListSource
{
    public async Task<IReadOnlyList<CodexModelInfo>> ReadAsync(CodexRuntimeConfig config, CancellationToken cancellationToken)
    {
        var executable = config.CodexExecutablePath;
        if (executable.EndsWith(".cmd", StringComparison.OrdinalIgnoreCase) || executable.EndsWith(".bat", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Model discovery requires the bundled native runtime.");
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(TimeSpan.FromSeconds(5));
        var token = timeout.Token;
        var start = new ProcessStartInfo(executable) { WorkingDirectory = Path.GetTempPath(),
            UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = true,
            RedirectStandardOutput = true, RedirectStandardError = true,
            StandardInputEncoding = new UTF8Encoding(false), StandardOutputEncoding = new UTF8Encoding(false) };
        Maily.Runtime.InstalledAppRuntime.UseChatGptSubscription(start);
        start.ArgumentList.Add("app-server"); start.ArgumentList.Add("--stdio");
        if (!string.IsNullOrWhiteSpace(config.CodexHome)) start.Environment["CODEX_HOME"] = config.CodexHome;
        using var process = new Process { StartInfo = start };
        Task? drain = null;
        try
        {
            if (!process.Start()) throw new InvalidOperationException("Runtime did not start.");
            drain = DrainAsync(process.StandardError, token);
            var frames = 0;
            async Task<JsonElement> CallAsync(int id, string method, object parameters)
            {
                await process.StandardInput.WriteLineAsync(JsonSerializer.Serialize(new { id, method, @params = parameters }).AsMemory(), token);
                await process.StandardInput.FlushAsync(token);
                while (++frames <= 200)
                {
                    using var json = JsonDocument.Parse(await ReadFrameAsync(process.StandardOutput, token));
                    var root = json.RootElement;
                    if (!root.TryGetProperty("id", out var responseId) || !responseId.TryGetInt32(out var value) || value != id) continue;
                    if (root.TryGetProperty("error", out _)) throw new InvalidOperationException("Model RPC failed.");
                    return root.GetProperty("result").Clone();
                }
                throw new InvalidOperationException("Too many runtime frames.");
            }
            await CallAsync(1, "initialize", new { clientInfo = new { name = "maily-model-catalog", version = "1.0" }, capabilities = new { experimentalApi = false } });
            await process.StandardInput.WriteLineAsync("{\"method\":\"initialized\"}".AsMemory(), token);
            await process.StandardInput.FlushAsync(token);
            var models = new List<CodexModelInfo>();
            var cursors = new HashSet<string>(StringComparer.Ordinal);
            string? cursor = null;
            for (var page = 0; page < 10; page++)
            {
                var result = await CallAsync(page + 2, "model/list", new { cursor, limit = 50, includeHidden = false });
                foreach (var model in result.GetProperty("data").EnumerateArray())
                {
                    if (model.TryGetProperty("hidden", out var hidden) && hidden.GetBoolean()) continue;
                    if (model.TryGetProperty("inputModalities", out var modalities) && !modalities.EnumerateArray().Any(m => m.GetString() == "text")) continue;
                    var id = model.GetProperty("model").GetString();
                    if (string.IsNullOrWhiteSpace(id) || id.Length > 200 || id.Any(char.IsControl)) continue;
                    if (models.Any(m => m.Id == id)) continue;
                    var efforts = model.TryGetProperty("supportedReasoningEfforts", out var supported)
                        ? supported.EnumerateArray().Select(e => e.GetProperty("reasoningEffort").GetString() ?? "").ToArray() : [];
                    models.Add(new(id, Text(model, "displayName", id), Text(model, "description", ""), efforts,
                        model.TryGetProperty("isDefault", out var isDefault) && isDefault.GetBoolean()));
                    if (models.Count > 500) throw new InvalidOperationException("Catalog too large.");
                }
                cursor = result.TryGetProperty("nextCursor", out var next) ? next.GetString() : null;
                if (string.IsNullOrEmpty(cursor)) return models;
                if (!cursors.Add(cursor)) throw new InvalidOperationException("Repeated catalog cursor.");
            }
            throw new InvalidOperationException("Too many catalog pages.");
        }
        finally
        {
            timeout.Cancel();
            try { if (!process.HasExited) process.Kill(entireProcessTree: true); } catch (InvalidOperationException) { }
            try { await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(2)); } catch (InvalidOperationException) { } catch (TimeoutException) { }
            if (drain is not null) { try { await drain.WaitAsync(TimeSpan.FromSeconds(2)); } catch { } }
        }
    }

    private static string Text(JsonElement element, string name, string fallback)
    {
        var text = element.TryGetProperty(name, out var value) ? value.GetString() ?? fallback : fallback;
        return text.Length <= 2000 ? text : text[..2000];
    }
    private static async Task DrainAsync(StreamReader reader, CancellationToken token)
    {
        var buffer = new char[4096];
        while (await reader.ReadAsync(buffer.AsMemory(), token) != 0) { }
    }
    private static async Task<string> ReadFrameAsync(StreamReader reader, CancellationToken token)
    {
        var buffer = new char[1];
        var line = new StringBuilder();
        while (line.Length < 262144)
        {
            if (await reader.ReadAsync(buffer.AsMemory(), token) == 0) throw new InvalidOperationException("Runtime closed its output.");
            if (buffer[0] == '\n') return line.ToString();
            line.Append(buffer[0]);
        }
        throw new InvalidOperationException("Runtime frame too large.");
    }
}
