using Maily.Generation.Models;

namespace Maily.Generation.Contracts;

public sealed class AddinSettingsRequest
{
	public string MailboxAddress { get; set; } = string.Empty;
	public string ProfileId { get; set; } = "default";
}

public sealed class AddinSettingsUpsertRequest
{
	public string MailboxAddress { get; set; } = string.Empty;
	public string ProfileId { get; set; } = "default";
	public AddinProfileSettings Settings { get; set; } = new();
}

public sealed class AddinSettingsResponse
{
	public string MailboxAddress { get; set; } = string.Empty;
	public string ProfileId { get; set; } = "default";
	public AddinProfileSettings Settings { get; set; } = new();
	public DateTimeOffset UpdatedUtc { get; set; } = DateTimeOffset.UtcNow;
}
