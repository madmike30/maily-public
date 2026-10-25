using System.Collections.Concurrent;
using Maily.Generation.Contracts;

namespace Maily.OutlookAddin.Web.Models;

public sealed class OutlookSearchFetchCache
{
	private static readonly TimeSpan EntryLifetime = TimeSpan.FromMinutes(2);
	private const int MaxEntries = 12;
	private readonly ConcurrentDictionary<string, OutlookSearchFetchCacheEntry> entries = new(StringComparer.OrdinalIgnoreCase);

	public OutlookSearchFetchCacheHit? TryGet(string mailboxAddress, string scope, string dateRange)
	{
		PruneExpired();
		var key = BuildKey(mailboxAddress, scope, dateRange);
		if (!entries.TryGetValue(key, out var entry))
		{
			return null;
		}

		var age = DateTimeOffset.UtcNow - entry.CreatedUtc;
		if (age > EntryLifetime)
		{
			entries.TryRemove(key, out _);
			return null;
		}

		return new OutlookSearchFetchCacheHit
		{
			Key = key,
			Age = age,
			Candidates = CloneCandidates(entry.Candidates),
			IsPartial = entry.IsPartial
		};
	}

	public void Set(string mailboxAddress, string scope, string dateRange, IReadOnlyList<OutlookSearchCorpusEmail> candidates, bool isPartial = false)
	{
		PruneExpired();
		var key = BuildKey(mailboxAddress, scope, dateRange);
		entries[key] = new OutlookSearchFetchCacheEntry
		{
			CreatedUtc = DateTimeOffset.UtcNow,
			Candidates = CloneCandidates(candidates),
			IsPartial = isPartial
		};

		PruneOverflow();
	}

	internal static List<OutlookSearchCorpusEmail> CloneCandidates(IReadOnlyList<OutlookSearchCorpusEmail> candidates) =>
		candidates.Select(email => new OutlookSearchCorpusEmail
		{
			Id = email.Id,
			Subject = email.Subject,
			FromDisplayName = email.FromDisplayName,
			FromAddress = email.FromAddress,
			FolderName = email.FolderName,
			ReceivedDateTime = email.ReceivedDateTime,
			BodyPreview = email.BodyPreview,
			BodyText = email.BodyText,
			IsRead = email.IsRead,
			Importance = email.Importance,
			SearchScore = email.SearchScore,
			SearchSignals = email.SearchSignals?.ToList() ?? []
		}).ToList();

	private void PruneExpired()
	{
		var cutoff = DateTimeOffset.UtcNow - EntryLifetime;
		foreach (var pair in entries)
		{
			if (pair.Value.CreatedUtc < cutoff)
			{
				entries.TryRemove(pair.Key, out _);
			}
		}
	}

	private void PruneOverflow()
	{
		if (entries.Count <= MaxEntries)
		{
			return;
		}

		foreach (var pair in entries
			.OrderBy(entry => entry.Value.CreatedUtc)
			.Take(Math.Max(0, entries.Count - MaxEntries)))
		{
			entries.TryRemove(pair.Key, out _);
		}
	}

	private static string BuildKey(string mailboxAddress, string scope, string dateRange) =>
		$"{(mailboxAddress ?? string.Empty).Trim().ToLowerInvariant()}|{(scope ?? string.Empty).Trim().ToLowerInvariant()}|{(dateRange ?? string.Empty).Trim().ToLowerInvariant()}";
}

public sealed class OutlookSearchFetchCacheHit
{
	public string Key { get; init; } = string.Empty;
	public TimeSpan Age { get; init; }
	public List<OutlookSearchCorpusEmail> Candidates { get; init; } = [];
	public bool IsPartial { get; init; }
}

internal sealed class OutlookSearchFetchCacheEntry
{
	public DateTimeOffset CreatedUtc { get; init; }
	public List<OutlookSearchCorpusEmail> Candidates { get; init; } = [];
	public bool IsPartial { get; init; }
}
