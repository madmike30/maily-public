using System.Collections.Concurrent;
using Maily.Generation.Contracts;

namespace Maily.OutlookAddin.Web.Models;

public sealed class OutlookSearchSessionStore
{
	private readonly ConcurrentDictionary<string, OutlookSearchSessionState> sessions = new(StringComparer.OrdinalIgnoreCase);

	public OutlookSearchSessionState CreateOrReplace(
		string mailboxAddress,
		string profileId,
		string query,
		OutlookSearchFilters filters,
		IReadOnlyList<OutlookSearchCorpusEmail> candidates, bool isPartial = false)
	{
		PruneExpired();
		var session = new OutlookSearchSessionState
		{
			SessionId = Guid.NewGuid().ToString("N"),
			MailboxAddress = mailboxAddress,
			ProfileId = profileId,
			Query = query,
			BaseFilters = CloneFilters(filters),
			CurrentFilters = CloneFilters(filters),
			BaseCandidates = candidates.ToList(),
			IsPartial = isPartial,
			CreatedUtc = DateTimeOffset.UtcNow,
			LastUpdatedUtc = DateTimeOffset.UtcNow
		};

		sessions[session.SessionId] = session;
		return session;
	}

	public OutlookSearchSessionState? Get(string sessionId)
	{
		PruneExpired();
		if (!sessions.TryGetValue(sessionId, out var session))
		{
			return null;
		}

		session.LastUpdatedUtc = DateTimeOffset.UtcNow;
		return session;
	}

	public void UpdateBaseCandidates(string sessionId, OutlookSearchFilters filters, IReadOnlyList<OutlookSearchCorpusEmail> candidates, bool isPartial = false)
	{
		if (!sessions.TryGetValue(sessionId, out var session))
		{
			return;
		}

		session.BaseFilters = CloneFilters(filters);
		session.CurrentFilters = CloneFilters(filters);
		session.BaseCandidates = candidates.ToList();
		session.IsPartial = isPartial;
		session.LastUpdatedUtc = DateTimeOffset.UtcNow;
	}

	public void UpdateCurrentFilters(string sessionId, OutlookSearchFilters filters)
	{
		if (!sessions.TryGetValue(sessionId, out var session))
		{
			return;
		}

		session.CurrentFilters = CloneFilters(filters);
		session.LastUpdatedUtc = DateTimeOffset.UtcNow;
	}

	private void PruneExpired()
	{
		var cutoff = DateTimeOffset.UtcNow.AddMinutes(-30);
		foreach (var pair in sessions)
		{
			if (pair.Value.LastUpdatedUtc < cutoff)
			{
				sessions.TryRemove(pair.Key, out _);
			}
		}
	}

	private static OutlookSearchFilters CloneFilters(OutlookSearchFilters filters) =>
		new()
		{
			Scope = filters.Scope,
			DateRange = filters.DateRange,
			Sender = filters.Sender,
			Focus = filters.Focus
		};
}

public sealed class OutlookSearchSessionState
{
	public string SessionId { get; init; } = string.Empty;
	public string MailboxAddress { get; init; } = string.Empty;
	public string ProfileId { get; init; } = "default";
	public string Query { get; init; } = string.Empty;
	public OutlookSearchFilters BaseFilters { get; set; } = new();
	public OutlookSearchFilters CurrentFilters { get; set; } = new();
	public List<OutlookSearchCorpusEmail> BaseCandidates { get; set; } = [];
	public bool IsPartial { get; set; }
	public DateTimeOffset CreatedUtc { get; init; } = DateTimeOffset.UtcNow;
	public DateTimeOffset LastUpdatedUtc { get; set; } = DateTimeOffset.UtcNow;
}
