namespace Maily.Maui.Models;

public sealed class OutlookEmail
{
	public required string Id { get; init; }
	public required string Subject { get; init; }
	public required string SenderName { get; init; }
	public required string SenderAddress { get; init; }
	public required string SenderLine { get; init; }
	public required string Importance { get; init; }
	public required string ReceivedAgo { get; init; }
	public required string ReceivedFull { get; init; }
	public required string Preview { get; init; }
	public required string BodyText { get; init; }
	public required string BodyHtml { get; init; }
	public required string Category { get; init; }
	public required string AccountName { get; init; }
	public required string TopicSummary { get; init; }
	public required string PriorityLabel { get; init; }
	public bool IsUnread { get; init; }
	public bool HasAttachments { get; init; }
	public IReadOnlyList<OutlookAttachmentInfo> Attachments { get; init; } = Array.Empty<OutlookAttachmentInfo>();
	public string BodyPreview => string.IsNullOrWhiteSpace(BodyText) ? Preview : BodyText;
	public string AttachmentsSummary => !HasAttachments
		? "No attachments"
		: Attachments.Count > 0
			? $"{Attachments.Count} attachment{(Attachments.Count == 1 ? string.Empty : "s")} included"
			: "Attachments loading...";

	public OutlookEmail WithAttachments(IReadOnlyList<OutlookAttachmentInfo> attachments) =>
		new()
		{
			Id = Id,
			Subject = Subject,
			SenderName = SenderName,
			SenderAddress = SenderAddress,
			SenderLine = SenderLine,
			Importance = Importance,
			ReceivedAgo = ReceivedAgo,
			ReceivedFull = ReceivedFull,
			Preview = Preview,
			BodyText = BodyText,
			BodyHtml = BodyHtml,
			Category = Category,
			AccountName = AccountName,
			TopicSummary = TopicSummary,
			PriorityLabel = PriorityLabel,
			IsUnread = IsUnread,
			HasAttachments = HasAttachments || attachments.Count > 0,
			Attachments = attachments
		};
}
