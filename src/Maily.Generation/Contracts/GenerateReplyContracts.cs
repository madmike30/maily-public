namespace Maily.Generation.Contracts;

public sealed class GenerateReplyRequest
{
	public GenerateReplyClientContext Client { get; set; } = new();
	public GenerateReplyComposeContext Compose { get; set; } = new();
	public GenerateReplyUserPreferences UserPreferences { get; set; } = new();
	public GenerateReplyOperation Operation { get; set; } = new();
}

public sealed class GenerateReplyClientContext
{
	public string Platform { get; set; } = "outlook-classic";
	public string HostVersion { get; set; } = string.Empty;
	public string MailboxAddress { get; set; } = string.Empty;
	public string MailboxDisplayName { get; set; } = string.Empty;
}

public sealed class GenerateReplyComposeContext
{
	public string Subject { get; set; } = string.Empty;
	public string FromAddress { get; set; } = string.Empty;
	public string FromDisplayName { get; set; } = string.Empty;
	public bool FromIsComposeSender { get; set; }
	public List<string> To { get; set; } = [];
	public List<string> Cc { get; set; } = [];
	public string BodyType { get; set; } = "html";
	public string? BodyHtml { get; set; }
	public string? BodyText { get; set; }
}

public sealed class GenerateReplyUserPreferences
{
	public string ProfileId { get; set; } = "default";
	public string? Tone { get; set; }
	public string? Length { get; set; }
	public string? Language { get; set; }
}

public sealed class GenerateReplyOperation
{
	public string Mode { get; set; } = "generate-initial";
	public string? CurrentDraft { get; set; }
	public string? UserInstruction { get; set; }
	public string? OutputLanguageOverride { get; set; }
}

public sealed class GenerateReplyResponse
{
	public GenerateReplyDraft Draft { get; set; } = new();
	public GenerateReplyMeta Meta { get; set; } = new();
}

public sealed class GenerateReplyDraft
{
	public string Html { get; set; } = string.Empty;
	public string Text { get; set; } = string.Empty;
}

public sealed class GenerateReplyMeta
{
	public string Model { get; set; } = string.Empty;
	public string PromptVersion { get; set; } = string.Empty;
	public string RequestId { get; set; } = string.Empty;
}
