namespace Maily.OutlookAddin.Web.Services;

public static class ComposeTempFile
{
    public static string CreatePath() => Path.Combine(Path.GetTempPath(), "maily-compose-" + Guid.NewGuid().ToString("N") + ".html");
}
