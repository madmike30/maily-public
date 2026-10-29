using Maily.Runtime;
using Maily.Generation.Models;
using Maily.Generation.Services;
var root=Path.GetFullPath(Path.Combine("output","installed-app-tests-"+Guid.NewGuid().ToString("N")));
Directory.CreateDirectory(root);
var count=0;
void Check(bool condition,string name){if(!condition)throw new Exception(name);count++;Console.WriteLine("PASS "+name);}
void Reject(Action action,string name){try{action();}catch(InvalidOperationException){Check(true,name);return;}throw new Exception(name);}
string Runtime(string relative,DateTime timestamp){var path=Path.GetFullPath(Path.Combine(root,relative));Directory.CreateDirectory(Path.GetDirectoryName(path)!);File.WriteAllText(path,"synthetic executable; do not run");File.SetLastWriteTimeUtc(path,timestamp);return path;}
Check(InstalledAppRuntime.GetCandidates(root).Count==0,"missing app has no fallback to PATH");
Reject(()=>InstalledAppRuntime.Resolve(root),"missing app actionable failure");
var old=Runtime("OpenAI/Codex/bin/old/codex.exe",DateTime.UtcNow.AddDays(-2));
var other=Runtime("arbitrary/codex.exe",DateTime.UtcNow.AddDays(1));
Check(InstalledAppRuntime.Resolve(root)==old,"known app directory only");
var options=new GenerationOptions{CodexExecutablePath=InstalledAppRuntime.Selection};
var loader=new CodexRuntimeConfigLoader(new PathResolver(options),options,()=>InstalledAppRuntime.Resolve(root));
Check(loader.Load().CodexExecutablePath==old,"sentinel resolves installed runtime");
var fresh=Runtime("OpenAI/Codex/bin/new/codex.exe",DateTime.UtcNow);
File.Move(old,old+".retained");
Check(loader.Load().CodexExecutablePath==fresh,"app update re-resolves runtime without stale pinned path");
var chat=Runtime("OpenAI/ChatGPT/bin/codex.exe",DateTime.UtcNow.AddHours(1));
Check(loader.Load().CodexExecutablePath==chat,"ChatGPT desktop runtime recognized");
options.CodexExecutablePath="explicit-pinned.exe";Check(loader.Load().CodexExecutablePath=="explicit-pinned.exe","explicit legacy path preserved");
options.CodexExecutablePath="";Check(loader.Load().CodexExecutablePath=="codex","legacy empty configuration preserved");
Check(InstalledAppRuntime.ReadPackageMode(root)=="Bundled","missing package marker legacy default");
var marker=Path.Combine(root,"ai-runtime-mode.txt");File.WriteAllText(marker,"InstalledApp\n");
Check(InstalledAppRuntime.ReadPackageMode(root)=="InstalledApp","installed-app marker recognized");
File.WriteAllText(marker,"unexpected");Reject(()=>InstalledAppRuntime.ReadPackageMode(root),"unknown package mode rejected");
Check(InstalledAppRuntime.IsChatGptLogin(true,0,"Logged in using ChatGPT"),"ChatGPT sign-in accepted");
Check(!InstalledAppRuntime.IsChatGptLogin(true,0,"Not logged in")&&!InstalledAppRuntime.IsChatGptLogin(true,1,"Logged in using ChatGPT")&&!InstalledAppRuntime.IsChatGptLogin(true,0,"Logged in using an API key"),"signed-out failed and API-key sessions rejected");
var missingStatus=CodexReadiness.Evaluate(false,-1,false,-1,"","",true);
Check(!missingStatus.Ok&&missingStatus.RecommendedAction.Contains("desktop app")&&!missingStatus.RecommendedAction.Contains("installer includes"),"missing installed runtime readiness guides desktop repair");
var signedOutStatus=CodexReadiness.Evaluate(true,0,true,1,"synthetic private diagnostic","",true);
Check(!signedOutStatus.Ok&&signedOutStatus.RecommendedAction==InstalledAppRuntime.MissingGuidance&&!signedOutStatus.ToString().Contains("private diagnostic"),"signed-out installed app readiness has safe desktop guidance");
Check(InstalledAppRuntime.IsSupportedVersion("codex-cli 0.151.0")&&InstalledAppRuntime.IsSupportedVersion("codex-cli 0.164.0-alpha.11.1"),"supported stable and newer desktop preview versions accepted");
Check(!InstalledAppRuntime.IsSupportedVersion("codex-cli 0.150.9")&&!InstalledAppRuntime.IsSupportedVersion("codex-cli 0.151.0-alpha.1")&&!InstalledAppRuntime.IsSupportedVersion("unknown version")&&!InstalledAppRuntime.IsSupportedVersion("codex-cli 0.164.0 extra text"),"old preview baseline and malformed versions rejected");
Check(!CodexReadiness.Evaluate(true,0,true,0,"Logged in using ChatGPT","",true,"codex-cli 0.150.9").Ok&&CodexReadiness.Evaluate(true,0,true,0,"Logged in using ChatGPT","",true,"codex-cli 0.164.0-alpha.11.1").Ok,"readiness requires supported desktop runtime version");
Console.WriteLine($"RESULT: {count} passed, 0 failed.");