namespace Maily.Generation.Contracts;

public sealed class StyleRefreshEmailSample
{
	public string Subject { get; set; } = string.Empty;
	public string BodyText { get; set; } = string.Empty;
	public string SentDateTime { get; set; } = string.Empty;
}

public sealed class OutlookStyleRefreshRequest
{
	public string MailboxAddress { get; set; } = string.Empty;
	public string ProfileId { get; set; } = "default";
	public string? Model { get; set; }
	public List<StyleRefreshEmailSample> Samples { get; set; } = [];
}

public sealed class OutlookStyleRefreshResponse
{
	public string RequestId { get; set; } = string.Empty;
	public string Model { get; set; } = string.Empty;
	public int SampleCount { get; set; }
	public string StyleSummary { get; set; } = string.Empty;
	public string AnalyzedStyle { get; set; } = string.Empty;
	public DateTimeOffset RefreshedUtc { get; set; } = DateTimeOffset.UtcNow;
}

public sealed class OutlookThreadSummaryRequest
{
	public string Surface { get; set; } = "compose";
	public string MailboxAddress { get; set; } = string.Empty;
	public string ProfileId { get; set; } = "default";
	public string? OutputLanguageOverride { get; set; }
	public string Subject { get; set; } = string.Empty;
	public string FromDisplayName { get; set; } = string.Empty;
	public string FromAddress { get; set; } = string.Empty;
	public List<string> To { get; set; } = [];
	public List<string> Cc { get; set; } = [];
	public string BodyText { get; set; } = string.Empty;
	public string? UserDraftText { get; set; }
}

public sealed class OutlookThreadSummaryResponse
{
	public OutlookSummarySections Sections { get; set; } = new();
	public OutlookSummaryMeta Meta { get; set; } = new();
}

public sealed class OutlookSummarySections
{
	public string ExecutiveSummary { get; set; } = string.Empty;
	public string MainPoints { get; set; } = string.Empty;
	public string ActionItems { get; set; } = string.Empty;
	public string RecommendedNextSteps { get; set; } = string.Empty;
}

public sealed class OutlookSummaryMeta
{
	public string RequestId { get; set; } = string.Empty;
	public string Model { get; set; } = string.Empty;
	public string Surface { get; set; } = string.Empty;
	public bool UsedDraftContext { get; set; }
}

public sealed class OutlookDailyBriefingRequest
{
	public string MailboxAddress { get; set; } = string.Empty;
	public string ProfileId { get; set; } = "default";
	public string? OutputLanguageOverride { get; set; }
	public List<OutlookBriefingMeetingItem> Meetings { get; set; } = [];
	public List<OutlookBriefingMailItem> OpenActionItems { get; set; } = [];
	public List<OutlookBriefingMailItem> ImportantUnreadEmails { get; set; } = [];
}

public sealed class OutlookBriefingMeetingItem
{
	public string Subject { get; set; } = string.Empty;
	public string StartDateTime { get; set; } = string.Empty;
	public string EndDateTime { get; set; } = string.Empty;
	public string Location { get; set; } = string.Empty;
	public string Organizer { get; set; } = string.Empty;
	public string Preview { get; set; } = string.Empty;
}

public sealed class OutlookBriefingMailItem
{
	public string Id { get; set; } = string.Empty;
	public string Subject { get; set; } = string.Empty;
	public string FromLine { get; set; } = string.Empty;
	public string ReceivedDateTime { get; set; } = string.Empty;
	public string FolderName { get; set; } = string.Empty;
	public string BodyPreview { get; set; } = string.Empty;
	public string WhyItMatters { get; set; } = string.Empty;
}

public sealed class OutlookDailyBriefingResponse
{
	public OutlookDailyBriefingSections Sections { get; set; } = new();
	public OutlookDailyBriefingMeta Meta { get; set; } = new();
}

public sealed class OutlookDailyBriefingSections
{
	public string RightNow { get; set; } = string.Empty;
	public string MeetingsToday { get; set; } = string.Empty;
	public string OpenActionItems { get; set; } = string.Empty;
	public string ImportantUnread { get; set; } = string.Empty;
	public string NextFocus { get; set; } = string.Empty;
}

public sealed class OutlookDailyBriefingMeta
{
	public string RequestId { get; set; } = string.Empty;
	public string Model { get; set; } = string.Empty;
	public int MeetingCount { get; set; }
	public int OpenActionItemCount { get; set; }
	public int ImportantUnreadCount { get; set; }
	public string GeneratedAtUtc { get; set; } = string.Empty;
}

public sealed class OutlookSearchCorpusEmail
{
	public string Id { get; set; } = string.Empty;
	public string Subject { get; set; } = string.Empty;
	public string FromDisplayName { get; set; } = string.Empty;
	public string FromAddress { get; set; } = string.Empty;
	public string FolderName { get; set; } = string.Empty;
	public string ReceivedDateTime { get; set; } = string.Empty;
	public string BodyPreview { get; set; } = string.Empty;
	public string BodyText { get; set; } = string.Empty;
	public bool IsRead { get; set; } = true;
	public string Importance { get; set; } = string.Empty;
	public double SearchScore { get; set; }
	public List<string> SearchSignals { get; set; } = [];
}

public sealed class OutlookSearchRequest
{
	public string MailboxAddress { get; set; } = string.Empty;
	public string ProfileId { get; set; } = "default";
	public string Query { get; set; } = string.Empty;
	public List<OutlookSearchCorpusEmail> Emails { get; set; } = [];
}

public sealed class OutlookSearchResponse
{
	public string RequestId { get; set; } = string.Empty;
	public string Model { get; set; } = string.Empty;
	public string AnswerSummary { get; set; } = string.Empty;
	public List<OutlookSearchResultItem> Results { get; set; } = [];
	public string Commitments { get; set; } = string.Empty;
}

public sealed class OutlookSearchResultItem
{
	public string Id { get; set; } = string.Empty;
	public string Subject { get; set; } = string.Empty;
	public string WhyItMatches { get; set; } = string.Empty;
	public string FromLine { get; set; } = string.Empty;
	public string FolderName { get; set; } = string.Empty;
	public string ReceivedDateTime { get; set; } = string.Empty;
	public string BodyPreview { get; set; } = string.Empty;
}

public sealed class OutlookSearchFilters
{
	public string Scope { get; set; } = "all-mail";
	public string DateRange { get; set; } = "90d";
	public string Sender { get; set; } = string.Empty;
	public string Focus { get; set; } = "all";
}

public sealed class OutlookSearchStartRequest
{
	public string MailboxAddress { get; set; } = string.Empty;
	public string ProfileId { get; set; } = "default";
	public string Query { get; set; } = string.Empty;
	public OutlookSearchFilters Filters { get; set; } = new();
}

public sealed class OutlookSearchRefineRequest
{
	public string SessionId { get; set; } = string.Empty;
	public OutlookSearchFilters Filters { get; set; } = new();
}

public sealed class OutlookSearchFacetOption
{
	public string Value { get; set; } = string.Empty;
	public string Label { get; set; } = string.Empty;
	public int Count { get; set; }
}

public sealed class OutlookSearchRefinements
{
	public List<OutlookSearchFacetOption> Scope { get; set; } = [];
	public List<OutlookSearchFacetOption> DateRange { get; set; } = [];
	public List<OutlookSearchFacetOption> Focus { get; set; } = [];
	public List<OutlookSearchFacetOption> Sender { get; set; } = [];
}

public sealed class OutlookSearchSessionMeta
{
	public bool IsPartial { get; set; }
	public string Warning { get; set; } = string.Empty;
	public string RequestId { get; set; } = string.Empty;
	public string SessionId { get; set; } = string.Empty;
	public string Model { get; set; } = string.Empty;
	public string QueryStage { get; set; } = string.Empty;
	public string ScopeLabel { get; set; } = string.Empty;
	public string EvidenceMode { get; set; } = string.Empty;
	public int CandidateCount { get; set; }
	public int ResultCount { get; set; }
	public int HydratedCount { get; set; }
	public string LastUpdatedUtc { get; set; } = string.Empty;
}

public sealed class OutlookSearchSessionResponse
{
	public string Query { get; set; } = string.Empty;
	public OutlookSearchFilters Filters { get; set; } = new();
	public OutlookSearchRefinements AvailableRefinements { get; set; } = new();
	public string AnswerSummary { get; set; } = string.Empty;
	public string Commitments { get; set; } = string.Empty;
	public List<OutlookSearchResultItem> Results { get; set; } = [];
	public OutlookSearchSessionMeta Meta { get; set; } = new();
}
