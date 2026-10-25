namespace Maily.OutlookAddin.Web.Models;

public sealed class ClientLogEntry
{
	public string EventName { get; init; } = string.Empty;
	public string MailboxAddress { get; init; } = string.Empty;
	public string Pane { get; init; } = string.Empty;
	public string Status { get; init; } = string.Empty;
	public string Detail { get; init; } = string.Empty;
}
