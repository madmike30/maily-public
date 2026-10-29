using System.Diagnostics;
using System.Windows.Forms;
using Maily.OutlookAddin.Setup;
internal static class ConnectionTests
{
 [STAThread] static void Main(string[] args)
 {
  if(args.Length>0 && args[0]=="fake") { Fake(args[1],args.ElementAtOrDefault(2)); return; }
  Run().GetAwaiter().GetResult();
  TestGuide(); TestIntegrity(); TestInstalledApp();
 }
 static void TestInstalledApp()
 {
  var root=Path.Combine(Path.GetTempPath(),"maily-installed-app-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(root);
  var local=Path.Combine(root,"synthetic-local");Directory.CreateDirectory(local);
  File.WriteAllText(Path.Combine(root,"ai-runtime-mode.txt"),"InstalledApp");
  if(CodexSetupPreflight.AllowsBrowserSignIn(root))throw new Exception("Installed app mode permits browser sign-in mutation");
  var calls=0;
  var absent=CodexSetupPreflight.Check(root,local,(path,command)=>{calls++;return (true,0,"Logged in using ChatGPT");});
  if(absent.Ok||calls!=0||!absent.RecommendedAction.Contains("desktop app"))throw new Exception("Missing app did not fail before process probe");
  var executable=Path.Combine(local,"OpenAI","ChatGPT","bin","v1","codex.exe");Directory.CreateDirectory(Path.GetDirectoryName(executable)!);File.WriteAllText(executable,"synthetic; never execute");
  var signedOut=CodexSetupPreflight.Check(root,local,(path,command)=>{calls++;return command=="--version"?(true,0,"codex-cli 0.164.0-alpha.11.1"):(true,1,"PRIVATE synthetic provider output");});
  if(signedOut.Ok||signedOut.ToDisplayText().Contains("PRIVATE")||!signedOut.RecommendedAction.Contains("sign in"))throw new Exception("Signed-out app guidance/privacy mismatch");
  var ready=CodexSetupPreflight.Check(root,local,(path,command)=>(true,0,command=="--version"?"codex-cli 0.164.0-alpha.11.1":"Logged in using ChatGPT"));
  if(!ready.Ok||ready.CodexExecutablePath!=executable)throw new Exception("Installed app sign-in not recognized");
  var newer=Path.Combine(local,"OpenAI","ChatGPT","bin","v2","codex.exe");Directory.CreateDirectory(Path.GetDirectoryName(newer)!);File.WriteAllText(newer,"synthetic broken runtime");File.SetLastWriteTimeUtc(newer,DateTime.UtcNow.AddMinutes(1));
  var probed=new List<string>();
  var brokenLatest=CodexSetupPreflight.Check(root,local,(path,command)=>{probed.Add(path);return path==newer?(true,1,"unavailable"):(true,0,command=="--version"?"codex-cli 0.164.0":"Logged in using ChatGPT");});
  if(brokenLatest.Ok||probed.Count!=1||probed[0]!=newer)throw new Exception("Setup approved an older runtime than generation would select");
  var files=new[]{"ai-runtime-mode.txt","setup/install.ps1","setup/uninstall.ps1","setup/OutlookShutdown.ps1","payload/addin/manifest.xml","payload/runtime/atk/node.exe","payload/runtime/atk/atk/cli.js","payload/scripts/connect-outlook.cjs","payload/scripts/outlook-credential-preflight.cjs","payload/app/Maily.OutlookAddin.Web/Maily.OutlookAddin.Web.exe"};
  var inventory=new List<object>();
  foreach(var file in files){var target=Path.Combine(root,file);Directory.CreateDirectory(Path.GetDirectoryName(target)!);if(file!="ai-runtime-mode.txt")File.WriteAllText(target,"synthetic");var bytes=File.ReadAllBytes(target);inventory.Add(new{path=file,bytes=bytes.Length,sha256=Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes))});}
  var inventoryPath=Path.Combine(root,"SHA256SUMS.json");File.WriteAllText(inventoryPath,System.Text.Json.JsonSerializer.Serialize(inventory));
  PackageIntegrity.Verify(root,false);
  File.WriteAllText(inventoryPath,System.Text.Json.JsonSerializer.Serialize(inventory.Skip(1)));
  var rejected=false;try{PackageIntegrity.Verify(root,false);}catch(InvalidOperationException){rejected=true;}
  if(!rejected)throw new Exception("Unverified runtime-mode marker accepted");
  Console.WriteLine("PASS installed app missing/signed-out/ready preflight, no browser sign-in, compact integrity and verified mode marker");
 }
 static void TestGuide()
 {
  using var owner=new SetupForm(SetupOptions.Parse([]),"Local",null);owner.ApplyInstallationResult(20,true);
  var calls=0;
  using var guide=owner.CreateConnectionGuide(token=>{calls++;return Task.FromResult(OutlookConnectionResult.For("registered"));});
  guide.ShowInTaskbar=false;guide.StartPosition=FormStartPosition.Manual;guide.Location=new System.Drawing.Point(-10000,-10000);
  guide.Show();Application.DoEvents();
  if(calls!=0)throw new Exception("Guide opening triggered sign-in without an explicit click");
  guide.Controls.OfType<Button>().Single(button=>button.Name=="connectionSignIn").PerformClick();Application.DoEvents();
  if(calls!=1)throw new Exception("Explicit sign-in did not invoke connector");
  if(owner.Controls.OfType<Button>().Any(button=>button.Text=="Open Outlook"))throw new Exception("Registration falsely verified Outlook cache");
  if(!guide.Controls.OfType<Label>().Single(label=>label.Name=="connectionStatus").Text.Contains("Reopen Outlook"))throw new Exception("Registration omitted verification guidance");
  guide.Close();
  owner.ApplyInstallationResult(23,false,"needs-support");
  using var blocked=owner.CreateConnectionGuide(token=>throw new Exception("Managed connection launched"));
  if(blocked.Controls.OfType<Button>().Single(button=>button.Name=="connectionSignIn").Enabled)throw new Exception("Support-required connection permits sign-in");
  Console.WriteLine("PASS explicit-only GUI connection and no false verification");
 }
 static void TestIntegrity()
 {
  var files=new[]{"setup/install.ps1","setup/uninstall.ps1","setup/OutlookShutdown.ps1","payload/addin/manifest.xml","payload/runtime/atk/node.exe","payload/runtime/atk/atk/cli.js","payload/scripts/connect-outlook.cjs","payload/scripts/outlook-credential-preflight.cjs"};
  foreach(var missing in files.Skip(2).Where(file=>file!="payload/addin/manifest.xml").Prepend(""))
  {
   var root=Path.Combine(Path.GetTempPath(),"maily-connection-integrity-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(root);
   var inventory=new List<object>();
   foreach(var file in files.Where(file=>file!=missing))
   {
    var target=Path.Combine(root,file);Directory.CreateDirectory(Path.GetDirectoryName(target)!);File.WriteAllText(target,"synthetic");
    var bytes=File.ReadAllBytes(target);inventory.Add(new{path=file,bytes=bytes.Length,sha256=Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes))});
   }
   File.WriteAllText(Path.Combine(root,"SHA256SUMS.json"),System.Text.Json.JsonSerializer.Serialize(inventory));
   var rejected=false;try{PackageIntegrity.Verify(root,true);}catch(InvalidOperationException){rejected=true;}
   if(rejected!=(missing!=""))throw new Exception("Connector integrity missing-file guard: "+missing);
   if(missing=="")
   {
    var oldOptions=Environment.GetEnvironmentVariable("NODE_OPTIONS");var oldPath=Environment.GetEnvironmentVariable("NODE_PATH");
    try {
     Environment.SetEnvironmentVariable("NODE_OPTIONS","synthetic-preload");Environment.SetEnvironmentVariable("NODE_PATH","synthetic-modules");
     var start=OutlookConnectionRunner.CreateStartInfo(root,Path.Combine(root,"sessions"));
     if(start.Environment.ContainsKey("NODE_OPTIONS") || start.Environment.ContainsKey("NODE_PATH") || start.ArgumentList.Count!=7)throw new Exception("Node environment or argument boundary failed");
    } finally {Environment.SetEnvironmentVariable("NODE_OPTIONS",oldOptions);Environment.SetEnvironmentVariable("NODE_PATH",oldPath);}
   }
  }
  Console.WriteLine("PASS connector and shutdown files required by package integrity");
 }
 static void Fake(string mode,string? pidPath)
 {
  if(pidPath is not null)File.WriteAllText(pidPath,Environment.ProcessId.ToString());
  if(mode=="hang"){Thread.Sleep(30000);return;}
  if(mode=="oversized"){Console.Write(new string('x',70000));return;}
  if(mode=="noise"){Console.WriteLine("PRIVATE RAW TOKEN");return;}
  if(mode=="failed"){Console.Error.WriteLine("PRIVATE STDERR TOKEN");Console.WriteLine("{\"status\":\"failed\",\"code\":\"credential-storage\",\"message\":\"PRIVATE RAW TOKEN\"}");Environment.ExitCode=1;return;}
  var id=mode=="wrong-id"?"wrong":"683ea353-6d71-4315-afd4-147e968605a9";
  Console.WriteLine("{\"status\":\"registered\",\"code\":\"registered\",\"manifestId\":\""+id+"\",\"message\":\"PRIVATE RAW TOKEN\"}");
 }
 static async Task Run()
 {
  foreach(var mode in new[]{"registered","wrong-id","noise","oversized","failed","hang"})
  {
   var start=new ProcessStartInfo(Environment.ProcessPath!);start.ArgumentList.Add("fake");start.ArgumentList.Add(mode);
   var result=await OutlookConnectionRunner.RunAsync(start,CancellationToken.None,TimeSpan.FromMilliseconds(mode=="hang"?300:10000));
   if(result.Registered!=(mode=="registered") || result.Message.Contains("PRIVATE"))throw new Exception(mode+" safety mismatch");
   Console.WriteLine("PASS connector runner "+mode);
  }
  using var cancelled=new CancellationTokenSource();cancelled.Cancel();
  var cancelledResult=await OutlookConnectionRunner.RunAsync(new ProcessStartInfo("MUST-NOT-START"),cancelled.Token);
  if(cancelledResult.Code!="cancelled")throw new Exception("Cancellation started helper");
  Console.WriteLine("PASS cancelled connector does not launch");
  using var midCancel=new CancellationTokenSource(TimeSpan.FromMilliseconds(600));
  var pidPath=Path.Combine(Path.GetTempPath(),"maily-connector-pid-"+Guid.NewGuid().ToString("N")+".txt");
  var hanging=new ProcessStartInfo(Environment.ProcessPath!);foreach(var argument in new[]{"fake","hang",pidPath})hanging.ArgumentList.Add(argument);
  var cancelledDuringRun=await OutlookConnectionRunner.RunAsync(hanging,midCancel.Token);
  if(cancelledDuringRun.Code!="cancelled")throw new Exception("Running cancellation was ignored");
  if(File.Exists(pidPath)){try{using var child=Process.GetProcessById(int.Parse(File.ReadAllText(pidPath)));if(!child.WaitForExit(1000))throw new Exception("Cancelled helper remained alive");}catch(ArgumentException){}}
  Console.WriteLine("PASS running cancellation terminates owned helper");
  var flow=new SetupFlow();flow.InstallationFinished(23,false,"needs-support");
  if(flow.CanConnectAutomatically)throw new Exception("Managed ownership allowed connection");
  flow.InstallationFinished(20,true,"activation-required");
  if(!flow.CanConnectAutomatically || flow.NextAction!=SetupAction.Connect)throw new Exception("Pending connection cannot connect");
  Console.WriteLine("PASS connector ownership and verification flow");
 }
}
