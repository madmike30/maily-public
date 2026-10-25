namespace Maily.Maui.Models;

public sealed class UserSettings
{
	public string MailboxAddress { get; set; } = string.Empty;
	public string UserName { get; set; } = "User";
	public int InboxLimit { get; set; } = 50;
	public string InboxDatePreset { get; set; } = "Last 14 days";
	public string? InboxStartDate { get; set; }
	public string? InboxEndDate { get; set; }
	public string Language { get; set; } = "English";
	public UserStyleSettings Style { get; set; } = new();
	public string? AnalyzedStyle { get; set; }
}

public sealed class UserStyleSettings
{
	public string Tone { get; set; } = "professional";
	public string LengthPreference { get; set; } = "concise";
	public string Greeting { get; set; } = "Hi";
	public string Closing { get; set; } = "Best wishes";
	public string CustomInstructions { get; set; } = string.Empty;
}
