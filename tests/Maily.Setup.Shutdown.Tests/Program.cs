using Maily.OutlookAddin.Setup;
internal static class ShutdownTests
{
 static void Main()
 {
  var root = Path.Combine(Path.GetTempPath(), "maily-shutdown-test-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(root);
  foreach (var (name, script, expected) in new[] { ("success", "exit 0", true), ("failure", "exit 7", false), ("blocked", "Start-Sleep -Seconds 30", false) })
  {
   var path = Path.Combine(root, name + ".ps1"); File.WriteAllText(path, script);
   var result = OutlookShutdown.Run(path, name == "blocked" ? 500 : 10000);
   if(result.Success != expected) throw new Exception(name + " result");
   Console.WriteLine("PASS GUI shared shutdown helper: " + name);
  }
  if(OutlookShutdown.Run(Path.Combine(root,"missing.ps1"),500).Success) throw new Exception("missing helper accepted");
  Console.WriteLine("PASS missing shutdown helper fails closed");
 }
}
