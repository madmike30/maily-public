using Maily.Generation.Contracts;
using Maily.OutlookAddin.Web.Models;

var cache = new OutlookSearchFetchCache();
var source = new OutlookSearchCorpusEmail
{
    Id = "synthetic-message", IsRead = false, Importance = "high",
    Subject = "Synthetic priority reminder", SearchSignals = ["unread"]
};
cache.Set("maily-test@example.invalid", "inbox", "week", [source]);
source.IsRead = true;
source.Importance = "low";
source.Subject = "Changed after storing";
source.SearchSignals.Clear();
var first = cache.TryGet("maily-test@example.invalid", "inbox", "week")
    ?? throw new Exception("Expected cache hit");
AssertMetadata(first.Candidates.Single());
first.Candidates[0].IsRead = true;
first.Candidates[0].Importance = "normal";
first.Candidates[0].SearchSignals.Add("changed");
first.Candidates.Clear();
var second = cache.TryGet("maily-test@example.invalid", "inbox", "week")
    ?? throw new Exception("Expected second cache hit");
AssertMetadata(second.Candidates.Single());
Console.WriteLine("PASS: cached unread/high-priority metadata survives store/read round trips and caller mutation");

static void AssertMetadata(OutlookSearchCorpusEmail item)
{
    if (item.IsRead || item.Importance != "high")
        throw new Exception("Unread/high-priority metadata was lost in cache round trip");
    if (item.Subject != "Synthetic priority reminder" || !item.SearchSignals.SequenceEqual(["unread"]))
        throw new Exception("Caller mutation leaked into cached candidate");
}
