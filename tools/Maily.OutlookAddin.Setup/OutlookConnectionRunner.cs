using System.Diagnostics;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
namespace Maily.OutlookAddin.Setup;

internal sealed record OutlookConnectionResult(bool Registered, string Code, string Message)
{
    public static OutlookConnectionResult For(string code) => code switch
    {
        "registered" => new(true,code,"Maily was added to the selected account. Reopen Outlook, then check the connection below."),
        "cancelled" => new(false,code,"Connection cancelled. Check Outlook before trying again; Microsoft may have completed the request."),
        "credential-storage" => new(false,code,"Secure sign-in storage is unavailable. Use the manual connection steps or contact support."),
        "sign-in-timeout" or "sign-in-incomplete" => new(false,code,"Microsoft sign-in did not finish. Try again, or use the manual steps for an unsupported account."),
        "timeout" or "registration-timeout" => new(false,code,"Connection took too long. Check Outlook before retrying, or use the manual steps."),
        "invalid-package" => new(false,code,"Connection files are missing or could not be verified. Extract a complete Maily package and retry."),
        "registration-failed" => new(false,code,"Microsoft could not add Maily. Use the manual steps or ask your administrator whether add-ins are allowed."),
        "connection-unavailable" => new(false,code,"Microsoft connection is unavailable. Try again or use the manual steps."),
        _ => new(false,"unverified-result","Connection is not confirmed. Check Outlook or use the manual steps; do not assume Maily was installed.")
    };
}

internal static class OutlookConnectionRunner
{
    private const string ManifestId = "683ea353-6d71-4315-afd4-147e968605a9";
    public static ProcessStartInfo CreateStartInfo(string packageRoot, string sessionRoot)
    {
        var runtime = Path.Combine(Path.GetFullPath(packageRoot), "payload", "runtime", "atk");
        var node = Path.Combine(runtime,"node.exe");
        var script = Path.Combine(packageRoot,"payload","scripts","connect-outlook.cjs");
        var manifest = Path.Combine(packageRoot,"payload","addin","manifest.xml");
        if(!File.Exists(node) || !File.Exists(script) || !File.Exists(manifest)) throw new FileNotFoundException("Connection files are missing.");
        var start = new ProcessStartInfo(node) { WorkingDirectory=Path.GetFullPath(packageRoot) };
        foreach(var argument in new[]{script,"--runtime-root",runtime,"--manifest",manifest,"--session-root",Path.GetFullPath(sessionRoot)}) start.ArgumentList.Add(argument);
        // Never allow inherited Node preload hooks to run ahead of the verified adapter.
        start.Environment.Remove("NODE_OPTIONS"); start.Environment.Remove("NODE_PATH");
        return start;
    }

    public static async Task<OutlookConnectionResult> RunAsync(ProcessStartInfo start, CancellationToken cancellationToken, TimeSpan? limit=null)
    {
        if(cancellationToken.IsCancellationRequested) return OutlookConnectionResult.For("cancelled");
        using var timeout=CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(limit ?? TimeSpan.FromMinutes(11));
        using var process=new Process { StartInfo=start };
        start.UseShellExecute=false;start.CreateNoWindow=true;start.RedirectStandardOutput=true;start.RedirectStandardError=true;
        var started=false;
        try
        {
            timeout.Token.ThrowIfCancellationRequested();
            process.Start();started=true;
            using var termination=timeout.Token.Register(()=>StopOwnedTree(process));
            var stdout=ReadBoundedAsync(process.StandardOutput,16384,true,timeout.Token);
            var stderr=ReadBoundedAsync(process.StandardError,65536,false,timeout.Token);
            var stdoutGuard=CancelOnFailure(stdout,timeout);
            var stderrGuard=CancelOnFailure(stderr,timeout);
            await Task.WhenAll(stdout,stderr,process.WaitForExitAsync(timeout.Token)).WaitAsync(timeout.Token).ConfigureAwait(false);
            await Task.WhenAll(stdoutGuard,stderrGuard).ConfigureAwait(false);
            using var json=JsonDocument.Parse(await stdout.ConfigureAwait(false));
            var root=json.RootElement;
            var status=root.GetProperty("status").GetString();
            var code=root.GetProperty("code").GetString();
            if(status=="registered" && code=="registered" && process.ExitCode==0 &&
                string.Equals(root.GetProperty("manifestId").GetString(),ManifestId,StringComparison.OrdinalIgnoreCase) &&
                (!root.TryGetProperty("titleId",out var title) || Regex.IsMatch(title.GetString() ?? "","^[A-Za-z0-9._-]{1,200}$")))
                return OutlookConnectionResult.For("registered");
            if(status=="failed" && process.ExitCode!=0 && code is "credential-storage" or "sign-in-timeout" or "sign-in-incomplete" or "registration-timeout" or "registration-failed" or "connection-unavailable" or "invalid-package")
                return OutlookConnectionResult.For(code);
            return OutlookConnectionResult.For("unverified-result");
        }
        catch(OperationCanceledException) { return OutlookConnectionResult.For(cancellationToken.IsCancellationRequested?"cancelled":"timeout"); }
        catch { return OutlookConnectionResult.For("unverified-result"); }
        finally { if(started) StopOwnedTree(process); }
    }
    private static Task CancelOnFailure(Task task,CancellationTokenSource timeout) => task.ContinueWith(done =>
    { if(done.IsFaulted) { _=done.Exception; try { timeout.Cancel(); } catch(ObjectDisposedException){} } },CancellationToken.None,TaskContinuationOptions.ExecuteSynchronously,TaskScheduler.Default);
    private static async Task<string> ReadBoundedAsync(StreamReader reader,int maximum,bool retain,CancellationToken token)
    {
        var text=new StringBuilder();var buffer=new char[1024];var total=0;
        while(true)
        {
            var count=await reader.ReadAsync(buffer.AsMemory(),token).ConfigureAwait(false);if(count==0)break;
            total+=count;if(total>maximum)throw new InvalidDataException("Connector output exceeded its bound.");
            if(retain)text.Append(buffer,0,count);
        }
        return text.ToString();
    }
    private static void StopOwnedTree(Process process) { try { if(!process.HasExited)process.Kill(entireProcessTree:true); } catch(InvalidOperationException){} catch(System.ComponentModel.Win32Exception){} }
}
