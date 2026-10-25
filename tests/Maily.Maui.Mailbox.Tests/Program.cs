using System.Diagnostics;
using System.Text;
using System.Text.Json;
using Maily.Maui.Services;

var root = Path.Combine(Directory.GetCurrentDirectory(), "output", "maui-mailbox-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);
var paths = new AppPaths(root);
File.WriteAllText(paths.SettingsFile, "{\"mailbox_address\":\" fixture@example.invalid \"}");
var service = new OutlookBridgeService(paths, new MailyLogger());
await service.ListInboxAsync(5);
Check(BridgeMock.Calls.Single().GetValueOrDefault("MailboxAddress") == "fixture@example.invalid", "inbox forwards explicitly configured mailbox");
Console.WriteLine("PASS inbox forwards explicitly configured mailbox");

await service.TestConnectionAsync();
await service.GetEmailBodyAsync("synthetic-id");
await service.GetAttachmentsAsync("synthetic-id");
await service.GetEmailAsync("synthetic-id");
await service.CreateReplyDraftAsync("synthetic-id", "<p>Synthetic fixture</p>", false);
await service.CreateReplyDraftAsync("synthetic-id", "<p>Synthetic fixture</p>", true);
Check(BridgeMock.Calls.Count == 8, "all MAUI bridge commands reached mock boundary");
Check(BridgeMock.Calls.All(call => call.GetValueOrDefault("MailboxAddress") == "fixture@example.invalid"), "every command retains explicit scope");
Check(BridgeMock.Calls.Any(call => call["Action"] == "test"), "connection readiness is mailbox-scoped");
Console.WriteLine("PASS all 8 calls forward the configured mailbox, including readiness and both reply modes");

var settingsService = new UserSettingsService(paths, new MailyLogger());
var settings = await settingsService.LoadAsync();
Check(settings.MailboxAddress == "fixture@example.invalid", "settings load preserves and trims mailbox");
settings.UserName = "Synthetic User";
settings.Style.Tone = "warm";
await settingsService.SaveAsync(settings);
var saved = await settingsService.LoadAsync();
Check(saved.MailboxAddress == "fixture@example.invalid" && saved.UserName == "Synthetic User" && saved.Style.Tone == "warm", "settings round-trip preserves mailbox and other settings");
File.WriteAllText(paths.SettingsFile, "{\"mailbox_address\":\"other@example.invalid\"}");
await service.ListInboxAsync(1);
Check(BridgeMock.Calls.Last()["MailboxAddress"] == "fixture@example.invalid", "editing settings cannot switch account under cached item IDs");
Console.WriteLine("PASS settings round-trip and account binding until restart");

var absentPaths = new AppPaths(Path.Combine(root, "missing"));
Directory.CreateDirectory(absentPaths.RepoRoot);
var absent = new OutlookBridgeService(absentPaths, new MailyLogger());
Func<Task>[] commands = [
    () => absent.TestConnectionAsync(),
    () => absent.ListInboxAsync(1),
    () => absent.GetEmailBodyAsync("synthetic-id"),
    () => absent.GetAttachmentsAsync("synthetic-id"),
    () => absent.GetEmailAsync("synthetic-id"),
    () => absent.CreateReplyDraftAsync("synthetic-id", "synthetic", false),
    () => absent.CreateReplyDraftAsync("synthetic-id", "synthetic", true)
];
var before = BridgeMock.Calls.Count;
foreach (var command in commands)
{
    try { await command(); throw new Exception("Missing mailbox unexpectedly succeeded"); }
    catch (InvalidOperationException ex) { Check(ex.Message.Contains("mailbox_address"), "actionable missing mailbox error"); }
}
Check(BridgeMock.Calls.Count == before, "missing mailbox fails before starting any bridge process");
Console.WriteLine("PASS all 7 public operations reject missing mailbox before process access");
File.WriteAllText(absentPaths.SettingsFile, "{\"mailbox_address\":\"  \"}");
var blank = new OutlookBridgeService(absentPaths, new MailyLogger());
try { await blank.TestConnectionAsync(); throw new Exception("Blank mailbox unexpectedly succeeded"); }
catch (InvalidOperationException) { }
Check(BridgeMock.Calls.Count == before, "blank mailbox does not probe Outlook");
BridgeMock.RejectMailbox = true;
try { await service.TestConnectionAsync(); throw new Exception("Unavailable mailbox reported ready"); }
catch (InvalidOperationException) { }
BridgeMock.RejectMailbox = false;
Console.WriteLine("PASS readiness rejects blank and unavailable accounts");
Console.WriteLine("RESULT: 5 MAUI mailbox contract groups passed; all process access mocked");
static void Check(bool condition, string message) { if (!condition) throw new Exception(message); }

namespace Maily.Maui.Services
{
    public sealed class AppPaths(string root)
    {
        public string RepoRoot => root;
        public string SettingsFile => Path.Combine(root, "user_settings.json");
        public string OutlookBridgeScript => Path.Combine(root, "mock-only.ps1");
    }
    public sealed class MailyLogger
    {
        public void Info(string name, object? data = null) { }
        public void Error(string name, Exception error, object? data = null) { }
    }
    public static class BridgeMock
    {
        public static readonly List<Dictionary<string, string>> Calls = [];
        public static bool RejectMailbox;
    }
    // Replaces only the OS process boundary in this linked-source test assembly.
    // No PowerShell, COM, Graph, account or mailbox access can occur.
    public sealed class Process : IDisposable
    {
        public static Process Start(ProcessStartInfo info)
        {
            var args = info.ArgumentList;
            var call = new Dictionary<string, string>();
            foreach (var name in new[] { "Action", "OutputPath", "MailboxAddress", "MailId", "HtmlPath" })
            {
                var index = args.IndexOf("-" + name);
                if (index >= 0) call[name] = args[index + 1];
            }
            BridgeMock.Calls.Add(call);
            var content = call["Action"] switch
            {
                "list-emails-preview" or "list-attachments" => "[]",
                "get-email" => "{\"id\":\"synthetic-id\",\"from\":{\"emailAddress\":{\"address\":\"sender@example.invalid\"}},\"body\":{\"text\":\"synthetic\"}}",
                _ => "{\"ok\":true}"
            };
            File.WriteAllText(call["OutputPath"], content);
            return new Process();
        }
        public int ExitCode => BridgeMock.RejectMailbox ? 1 : 0;
        public StreamReader StandardError { get; } = new(new MemoryStream(Encoding.UTF8.GetBytes("")));
        public Task WaitForExitAsync(CancellationToken token) => Task.CompletedTask;
        public void Dispose() => StandardError.Dispose();
    }
}
