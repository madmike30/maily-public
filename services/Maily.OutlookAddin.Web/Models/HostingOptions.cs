namespace Maily.OutlookAddin.Web.Models;

public sealed class HostingOptions
{
	public string ServiceName { get; set; } = "MailyOutlookAddinHost";
	public string HttpsUrl { get; set; } = "https://localhost:7097";
	public string HttpUrl { get; set; } = "http://localhost:5289";
	public string? LogDirectory { get; set; }
	public string? CertificatePath { get; set; }
	public string? CertificatePassword { get; set; }
	public bool RequireOutlookBridge { get; set; } = true;
	public bool EnableHttp { get; set; }
	public bool EnableRequestLog { get; set; }
	public bool EnableEventLog { get; set; }
	public int ReadinessCacheSeconds { get; set; } = 30;
}
