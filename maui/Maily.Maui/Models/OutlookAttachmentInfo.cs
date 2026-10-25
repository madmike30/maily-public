namespace Maily.Maui.Models;

public sealed class OutlookAttachmentInfo
{
	public required string Id { get; init; }
	public required string Name { get; init; }
	public int Index { get; init; }
	public long Size { get; init; }
}
