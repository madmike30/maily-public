namespace Maily.OutlookAddin.Web.Services;

public static class SearchCoverage
{
    // A full window may have additional unseen items, including equal-date ties.
    // Count before deduplication and before ranking/filtering.
    public static bool IsLimited(int rawCount, int limit) => limit > 0 && rawCount >= limit;

    public static string MergeWarning(string? warning, bool isPartial)
    {
        const string coverage = "Results cover a limited subset of this mailbox. Some messages may be missing; narrow the date range or scope.";
        return isPartial ? string.Join(" ", new[] { warning, coverage }.Where(value => !string.IsNullOrWhiteSpace(value))) : warning ?? string.Empty;
    }
}
