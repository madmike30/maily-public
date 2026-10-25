using Maily.OutlookAddin.Web.Services;

var directory = Path.GetFullPath(args.Length > 0 ? args[0] : Path.Combine("output", "temp-file-test-" + Guid.NewGuid().ToString("N")));
Directory.CreateDirectory(directory);
Environment.SetEnvironmentVariable("TMP", directory);
Environment.SetEnvironmentVariable("TEMP", directory);
Environment.SetEnvironmentVariable("TMPDIR", directory);
var paths = Enumerable.Range(0, 32).Select(_ => ComposeTempFile.CreatePath()).ToArray();
var failures = 0;
void Check(bool condition, string name) { Console.WriteLine($"{(condition ? "PASS" : "FAIL")}: {name}"); if (!condition) failures++; }
Check(paths.Distinct().Count() == 32, "concurrent request names are distinct");
Check(paths.All(path => Path.GetDirectoryName(path) == directory && Path.GetExtension(path) == ".html"), "paths stay in isolated temporary directory");
Check(Directory.GetFiles(directory).Length == 0, "allocating paths creates no orphan base files");
await File.WriteAllTextAsync(paths[0], "<p>Synthetic draft</p>");
Check(Directory.GetFiles(directory).Length == 1 && await File.ReadAllTextAsync(paths[0]) == "<p>Synthetic draft</p>", "writing a draft creates exactly the intended file");
Console.WriteLine($"RESULT: {4-failures} passed, {failures} failed");
Console.WriteLine($"Retained synthetic fixture directory: {directory}");
return failures == 0 ? 0 : 1;
