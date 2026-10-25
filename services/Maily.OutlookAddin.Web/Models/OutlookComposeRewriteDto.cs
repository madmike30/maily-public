namespace Maily.OutlookAddin.Web.Models;

public sealed class OutlookComposeRewriteRequest
{
	public string Subject { get; init; } = string.Empty;
	public string UserName { get; init; } = string.Empty;
	public string MailboxAddress { get; init; } = string.Empty;
	public string SeedText { get; init; } = string.Empty;
	public string ComposeType { get; init; } = "unknown";
	public string Text { get; init; } = string.Empty;
	public string Html { get; init; } = string.Empty;
}
