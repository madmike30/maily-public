using System.Globalization;
using System.Text.RegularExpressions;
using Maily.Generation.Contracts;

namespace Maily.Generation.Services;

public sealed class OutlookSearchPreparationService
{
	private static readonly HashSet<string> StopWords = new(StringComparer.OrdinalIgnoreCase)
	{
		"a", "an", "and", "are", "as", "at", "be", "by", "did", "do", "for", "from", "have", "i",
		"in", "is", "it", "me", "my", "of", "on", "or", "our", "please", "re", "regarding", "the",
		"their", "them", "they", "this", "to", "was", "we", "were", "what", "when", "where", "who",
		"will", "with", "you", "your",
		"את", "אתה", "אתם", "אני", "אם", "גם", "היא", "הם", "זה", "זו", "טוב", "כל", "כבר", "כי",
		"לא", "לי", "לו", "מה", "מי", "מייל", "מיילים", "נשלח", "נשלחו", "על", "עם", "רק", "של",
		"שלחתי", "שנשלחתי", "ששלחתי", "תודה"
	};

	private static readonly SearchConcept[] SearchConcepts =
	[
		new(
			"pricing",
			[
				"price",
				"prices",
				"pricing",
				"cost",
				"costing",
				"estimate",
				"estimates",
				"estimation",
				"estimations",
				"quote",
				"quotes",
				"quotation",
				"proposal",
				"proposals",
				"commercial",
				"budget",
				"investment proposal",
				"rough cost estimate",
				"unit price",
				"monthly total",
				"quote is for investment proposal only",
				"מחיר",
				"מחירים",
				"תמחור",
				"הצעת מחיר",
				"הצעות מחיר",
				"אומדן",
				"הערכה",
				"הערכות"
			]),
		new(
			"autonomous",
			[
				"autonomous",
				"autonomous db",
				"autonomous database",
				"autonomous ai database",
				"autonomous ai lakehouse",
				"adb",
				"lakehouse",
				"אוטונומ",
				"אטונומ",
				"לקהאוס"
			]),
		new(
			"commitments",
			[
				"action",
				"actions",
				"commit",
				"commitment",
				"follow up",
				"next step",
				"task",
				"tasks",
				"משימה",
				"משימות",
				"פעולה",
				"פעולות",
				"המשך טיפול"
			]),
		new(
			"decisions",
			[
				"decision",
				"decisions",
				"decide",
				"approved",
				"agreed",
				"resolve",
				"החלטה",
				"הוחלט",
				"מאושר"
			])
	];

	private static readonly string[] CommitmentTerms =
	[
		"commit",
		"committed",
		"deliver",
		"delivered",
		"delivering",
		"follow up",
		"follow-up",
		"i will",
		"we will",
		"will send",
		"will share",
		"will update",
		"will follow up",
		"i'll",
		"we'll",
		"next step",
		"owner",
		"promise",
		"ownership",
		"אשלח",
		"נשלח",
		"אעביר",
		"נעביר",
		"אשתף",
		"נשתף",
		"אעדכן",
		"נעדכן",
		"נחזור",
		"נבדוק",
		"באחריות",
		"לטיפול",
		"השלב הבא",
		"צעד הבא"
	];

	private static readonly string[] DecisionTerms =
	[
		"agree",
		"agreed",
		"agreed to",
		"approved",
		"approved to",
		"confirm",
		"confirmed",
		"decided",
		"resolved",
		"final decision",
		"decision was",
		"we decided",
		"it was decided",
		"הוחלט",
		"סוכם",
		"הוסכם",
		"מאושר",
		"אושר",
		"אישרנו"
	];

	private static readonly string[] CommercialEvidenceTerms =
	[
		"price",
		"prices",
		"pricing",
		"cost",
		"costing",
		"estimate",
		"estimates",
		"estimation",
		"estimations",
		"quote",
		"quotes",
		"quotation",
		"proposal",
		"proposals",
		"commercial",
		"budget",
		"investment proposal",
		"rough cost estimate",
		"non binding rough cost estimate",
		"monthly total",
		"unit price",
		"oci credits",
		"cpq",
		"pricing notes",
		"מחיר",
		"מחירים",
		"תמחור",
		"הצעת מחיר",
		"הצעות מחיר",
		"אומדן",
		"הערכה",
		"הערכות",
		"עלות",
		"תקציב"
	];

	private static readonly string[] NumericCommercialEvidenceTerms =
	[
		"usd",
		"ils",
		"nis",
		"eur",
		"dollar",
		"dollars",
		"monthly total",
		"per month",
		"price list",
		"credits"
	];

	private static readonly string[] GenericTopicNoiseTerms =
	[
		"what's new",
		"welcome",
		"newsletter",
		"learning session",
		"update",
		"updates",
		"webinar",
		"event",
		"events"
	];

	private static readonly string[] AutomatedMailboxNoiseTerms =
	[
		"digest",
		"new messages",
		"accountability message",
		"action recommended",
		"register now",
		"last chance to register",
		"curated list",
		"partner showcase",
		"learning material",
		"subscription update",
		"subscriptions updates",
		"broadcast system",
		"product update webcast",
		"announcements"
	];

	private static readonly string[] AutomatedSenderNoiseTerms =
	[
		"postmaster",
		"noreply",
		"no-reply",
		"do-not-reply",
		"mailer-daemon",
		"notification",
		"notifications",
		"alerts"
	];

	private static readonly Regex SentenceBoundaryRegex = new(@"(?<=[\.\!\?])\s+", RegexOptions.Compiled);
	private static readonly Regex TokenRegex = new(@"[\p{L}\p{Nd}][\p{L}\p{Nd}_\-\.@]{1,}", RegexOptions.Compiled | RegexOptions.CultureInvariant);
	private static readonly Regex MoneyAmountRegex = new(@"(?:(?:[$€£₪]|usd|eur|ils|nis)\s*\d[\d,\.]*|\d[\d,\.]*\s*(?:usd|eur|ils|nis|k|m))", RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
	private static readonly Regex AutomatedDigestSubjectRegex = new(@"\b(?:digest|new messages|action recommended|register now)\b", RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

	public IReadOnlyList<OutlookSearchCorpusEmail> RankCandidates(string query, IEnumerable<OutlookSearchCorpusEmail> emails, int maxCount)
	{
		var queryProfile = BuildQueryProfile(query);

		return emails
			.Where(email => !string.IsNullOrWhiteSpace(email.Id))
			.Select(email => ScoreEmail(email, queryProfile))
			.OrderByDescending(email => email.SearchScore)
			.ThenByDescending(ParseReceivedUtcOrMin)
			.Take(maxCount)
			.ToList();
	}

	public IReadOnlyList<OutlookSearchCorpusEmail> SelectStrongResultCluster(
		IReadOnlyList<OutlookSearchCorpusEmail> rankedEmails,
		int maxCount = 5)
	{
		if (rankedEmails.Count <= 1 || maxCount <= 0)
		{
			return rankedEmails.Take(Math.Max(0, maxCount)).ToList();
		}

		var selected = new List<OutlookSearchCorpusEmail>(Math.Min(maxCount, rankedEmails.Count));
		var topScore = rankedEmails[0].SearchScore;
		double? previousScore = null;

		foreach (var email in rankedEmails)
		{
			if (selected.Count >= maxCount)
			{
				break;
			}

			if (selected.Count == 0)
			{
				selected.Add(email);
				previousScore = email.SearchScore;
				continue;
			}

			var currentScore = email.SearchScore;
			var previousGap = previousScore is null ? 0 : previousScore.Value - currentScore;
			var topGap = topScore - currentScore;
			var shouldStop =
				(selected.Count >= 2 && currentScore <= 0 && topScore >= 100) ||
				(selected.Count >= 2 && previousGap >= 140) ||
				(selected.Count >= 3 && topGap >= 220);

			if (shouldStop)
			{
				break;
			}

			selected.Add(email);
			previousScore = currentScore;
		}

		return selected;
	}

	public string BuildCommitmentsSummary(IEnumerable<OutlookSearchCorpusEmail> emails)
	{
		var lines = emails
			.SelectMany(email => ExtractCommitmentLines(email).Select(line => new
			{
				Email = email,
				Line = line
			}))
			.GroupBy(entry => entry.Line, StringComparer.OrdinalIgnoreCase)
			.Select(group =>
			{
				var newest = group
					.OrderByDescending(entry => ParseReceivedUtcOrMin(entry.Email))
					.First();
				return $"- {newest.Line} ({FormatEmailReference(newest.Email)})";
			})
			.Take(5)
			.ToList();

		return lines.Count == 0
			? string.Empty
			: string.Join(Environment.NewLine, lines);
	}

	public IReadOnlyList<OutlookSearchCorpusEmail> ApplyFilters(
		IEnumerable<OutlookSearchCorpusEmail> emails,
		OutlookSearchFilters filters)
	{
		var scope = NormalizeScope(filters.Scope);
		var sender = (filters.Sender ?? string.Empty).Trim();
		var focus = NormalizeFocus(filters.Focus);
		var startDate = ResolveStartDate(filters.DateRange);

		return emails
			.Where(email => MatchesScope(email, scope))
			.Where(email => MatchesSender(email, sender))
			.Where(email => MatchesDate(email, startDate))
			.Where(email => MatchesFocus(email, focus))
			.ToList();
	}

	public List<OutlookSearchFacetOption> BuildSenderRefinements(IEnumerable<OutlookSearchCorpusEmail> emails, int maxCount = 8)
	{
		return emails
			.Where(email => !string.IsNullOrWhiteSpace(email.FromAddress) || !string.IsNullOrWhiteSpace(email.FromDisplayName))
			.GroupBy(GetSenderKey, StringComparer.OrdinalIgnoreCase)
			.Select(group =>
			{
				var sample = group.First();
				var label = string.IsNullOrWhiteSpace(sample.FromDisplayName)
					? sample.FromAddress
					: $"{sample.FromDisplayName} <{sample.FromAddress}>";
				return new OutlookSearchFacetOption
				{
					Value = group.Key,
					Label = label,
					Count = group.Count()
				};
			})
			.OrderByDescending(option => option.Count)
			.ThenBy(option => option.Label, StringComparer.OrdinalIgnoreCase)
			.Take(maxCount)
			.ToList();
	}

	public static List<OutlookSearchFacetOption> BuildStaticScopeOptions() =>
	[
		new() { Value = "all-mail", Label = "All mail", Count = 0 },
		new() { Value = "inbox", Label = "Inbox", Count = 0 },
		new() { Value = "sent-items", Label = "Sent Items", Count = 0 }
	];

	public static List<OutlookSearchFacetOption> BuildStaticDateRangeOptions() =>
	[
		new() { Value = "30d", Label = "Last 30 days", Count = 0 },
		new() { Value = "90d", Label = "Last 90 days", Count = 0 },
		new() { Value = "180d", Label = "Last 180 days", Count = 0 },
		new() { Value = "365d", Label = "Last year", Count = 0 },
		new() { Value = "all", Label = "Any time", Count = 0 }
	];

	public static List<OutlookSearchFacetOption> BuildStaticFocusOptions() =>
	[
		new() { Value = "all", Label = "All matches", Count = 0 },
		new() { Value = "commitments", Label = "Commitments", Count = 0 },
		new() { Value = "decisions", Label = "Decisions", Count = 0 }
	];

	private OutlookSearchCorpusEmail ScoreEmail(
		OutlookSearchCorpusEmail email,
		SearchQueryProfile queryProfile)
	{
		var subject = NormalizeText(email.Subject);
		var from = NormalizeText($"{email.FromDisplayName} {email.FromAddress}");
		var preview = NormalizeText(email.BodyPreview);
		var body = NormalizeText(email.BodyText);
		var evidenceText = string.IsNullOrWhiteSpace(body) ? preview : $"{preview}\n{body}";
		var subjectWithoutReplyPrefix = StripReplyPrefixes(subject);
		var combined = $"{subject}\n{subjectWithoutReplyPrefix}\n{from}\n{preview}\n{body}";
		var emailTokens = Tokenize(combined);
		var genericNoiseHits = CountMatches(subject, GenericTopicNoiseTerms) + CountMatches(evidenceText, GenericTopicNoiseTerms);
		var senderNoiseHits = CountMatches(from, AutomatedSenderNoiseTerms);
		var automatedNoiseHits = CountMatches(subject, AutomatedMailboxNoiseTerms) + senderNoiseHits;
		if (AutomatedDigestSubjectRegex.IsMatch(subject) || AutomatedDigestSubjectRegex.IsMatch(subjectWithoutReplyPrefix))
		{
			automatedNoiseHits += 1;
		}

		var overlapCount = queryProfile.Tokens.Count == 0
			? 0
			: queryProfile.Tokens.Count(token => emailTokens.Contains(token));

		double score = 0;
		var signals = new List<string>();

		if (!string.IsNullOrWhiteSpace(queryProfile.NormalizedQuery))
		{
			if (subject.Contains(queryProfile.NormalizedQuery, StringComparison.Ordinal) ||
				subjectWithoutReplyPrefix.Contains(queryProfile.NormalizedQuery, StringComparison.Ordinal))
			{
				score += 80;
				signals.Add("subject phrase");
			}

			if (from.Contains(queryProfile.NormalizedQuery, StringComparison.Ordinal))
			{
				score += 45;
				signals.Add("sender phrase");
			}

			if (preview.Contains(queryProfile.NormalizedQuery, StringComparison.Ordinal) || body.Contains(queryProfile.NormalizedQuery, StringComparison.Ordinal))
			{
				score += 55;
				signals.Add("body phrase");
			}
		}

		if (overlapCount > 0)
		{
			score += overlapCount * 10;
			signals.Add(overlapCount == 1 ? "1 query term" : $"{overlapCount} query terms");
		}

		var subjectOverlap = queryProfile.Tokens.Count(token =>
			subject.Contains(token, StringComparison.Ordinal) ||
			subjectWithoutReplyPrefix.Contains(token, StringComparison.Ordinal));
		if (subjectOverlap > 0)
		{
			score += subjectOverlap * 16;
			signals.Add(subjectOverlap == 1 ? "subject keyword" : "subject keywords");
		}

		var senderOverlap = queryProfile.Tokens.Count(token => from.Contains(token, StringComparison.Ordinal));
		if (senderOverlap > 0)
		{
			score += senderOverlap * 12;
			signals.Add(senderOverlap == 1 ? "sender keyword" : "sender keywords");
		}

		var matchedConcepts = 0;
		foreach (var conceptMatch in queryProfile.Concepts)
		{
			var subjectHits = CountMatches(subject, conceptMatch.Terms) + CountMatches(subjectWithoutReplyPrefix, conceptMatch.Terms);
			var bodyHits = CountMatches(preview, conceptMatch.Terms) + CountMatches(body, conceptMatch.Terms);
			var exactSubjectHits = CountMatches(subject, conceptMatch.QueryTerms) + CountMatches(subjectWithoutReplyPrefix, conceptMatch.QueryTerms);
			var exactBodyHits = CountMatches(preview, conceptMatch.QueryTerms) + CountMatches(body, conceptMatch.QueryTerms);
			if (subjectHits == 0 && bodyHits == 0)
			{
				score -= GetMissingConceptPenalty(conceptMatch);
				continue;
			}

			matchedConcepts += 1;
			score += subjectHits * (conceptMatch.IsHighIntent ? 42 : 22);
			score += bodyHits * (conceptMatch.IsHighIntent ? 28 : 14);
			score += exactSubjectHits * (conceptMatch.IsHighIntent ? 55 : 20);
			score += exactBodyHits * (conceptMatch.IsHighIntent ? 36 : 14);

			if (subjectHits > 0 && bodyHits > 0)
			{
				score += conceptMatch.IsHighIntent ? 30 : 14;
			}

			if (conceptMatch.QueryTerms.Length > 0 && exactSubjectHits == 0 && exactBodyHits == 0)
			{
				score -= conceptMatch.Name == "pricing" ? 90 : conceptMatch.IsHighIntent ? 45 : 8;
			}

			signals.Add(conceptMatch.SignalLabel);
		}

		if (queryProfile.Concepts.Count > 1)
		{
			if (matchedConcepts == queryProfile.Concepts.Count)
			{
				score += 120;
				signals.Add("matched full intent");
			}
			else if (matchedConcepts == 0)
			{
				score -= 90;
			}
			else
			{
				var missingConcepts = queryProfile.Concepts.Count - matchedConcepts;
				var missingHighIntentConcepts = queryProfile.Concepts.Count(concept => concept.IsHighIntent) - queryProfile.Concepts.Count(concept =>
					concept.IsHighIntent &&
					(CountMatches(subject, concept.Terms) + CountMatches(subjectWithoutReplyPrefix, concept.Terms) + CountMatches(preview, concept.Terms) + CountMatches(body, concept.Terms)) > 0);
				score -= (missingConcepts * 24) + (missingHighIntentConcepts * 70);
			}
		}

		if (!string.IsNullOrWhiteSpace(email.BodyText))
		{
			score += 8;
			signals.Add("full body");
		}

		var commitmentHits = CountCommitmentSignals(preview + "\n" + body);
		var decisionHits = CountDecisionSignals(preview + "\n" + body);

		if (queryProfile.CommitmentQuery)
		{
			if (commitmentHits > 0)
			{
				score += Math.Min(42, 12 + (commitmentHits * 8));
				signals.Add(commitmentHits == 1 ? "commitment signal" : "commitment signals");
			}
			else
			{
				score -= genericNoiseHits > 0 ? 145 : 90;
			}
		}

		if (queryProfile.DecisionQuery)
		{
			if (decisionHits > 0)
			{
				score += Math.Min(40, 10 + (decisionHits * 8));
				signals.Add(decisionHits == 1 ? "decision signal" : "decision signals");
			}
			else
			{
				score -= genericNoiseHits > 0 ? 165 : 105;
			}
		}

		var exactSubjectQueryMatch = !string.IsNullOrWhiteSpace(queryProfile.NormalizedQuery) &&
			(subject.Contains(queryProfile.NormalizedQuery, StringComparison.Ordinal) ||
			 subjectWithoutReplyPrefix.Contains(queryProfile.NormalizedQuery, StringComparison.Ordinal));
		var nearExactSubjectTokenMatch = queryProfile.Tokens.Count >= 2 &&
			subjectOverlap >= queryProfile.Tokens.Count;
		var strongSubjectQueryMatch = exactSubjectQueryMatch || nearExactSubjectTokenMatch;
		var tokenCoverage = queryProfile.Tokens.Count == 0
			? 0
			: (double)overlapCount / queryProfile.Tokens.Count;
		var replyThreadSubject = !string.IsNullOrWhiteSpace(email.Subject) &&
			(email.Subject.StartsWith("RE:", StringComparison.OrdinalIgnoreCase) ||
			 email.Subject.StartsWith("FW:", StringComparison.OrdinalIgnoreCase) ||
			 email.Subject.StartsWith("FWD:", StringComparison.OrdinalIgnoreCase));

		if (automatedNoiseHits > 0)
		{
			var focusIntentQuery = queryProfile.CommitmentQuery || queryProfile.DecisionQuery || queryProfile.CommercialIntent || queryProfile.AutonomousIntent;
			var strongEvidenceHits = matchedConcepts + commitmentHits + decisionHits + subjectOverlap;
			if (focusIntentQuery && strongEvidenceHits == 0)
			{
				score -= strongSubjectQueryMatch ? 28 : 120 + (automatedNoiseHits * 10);
				signals.Add("automated noise");
			}
			else if (!strongSubjectQueryMatch && strongEvidenceHits == 0 && overlapCount <= 1)
			{
				score -= 48 + (automatedNoiseHits * 10);
				signals.Add("automated noise");
			}

			var broadTopicQuery = !focusIntentQuery && queryProfile.Tokens.Count >= 2;
			var weakThreadContext = !replyThreadSubject && string.IsNullOrWhiteSpace(body);
			if (!strongSubjectQueryMatch &&
				broadTopicQuery &&
				(senderNoiseHits > 0 || genericNoiseHits > 0) &&
				(tokenCoverage < 0.95 || weakThreadContext))
			{
				score -= 90 + (automatedNoiseHits * 22) + (genericNoiseHits * 12);
				signals.Add("automated topic noise");
			}
		}

		if (queryProfile.CommercialIntent)
		{
			var subjectCommercialHits = CountMatches(subject, CommercialEvidenceTerms) + CountMatches(subjectWithoutReplyPrefix, CommercialEvidenceTerms);
			var bodyCommercialHits = CountMatches(evidenceText, CommercialEvidenceTerms);
			var numericCommercialHits = CountMatches(evidenceText, NumericCommercialEvidenceTerms);
			var hasMoneyAmount = MoneyAmountRegex.IsMatch(evidenceText);
			var autonomousHits = CountMatches(subject, SearchConcepts[1].Terms) + CountMatches(subjectWithoutReplyPrefix, SearchConcepts[1].Terms) + CountMatches(evidenceText, SearchConcepts[1].Terms);
			if (subjectCommercialHits > 0 || bodyCommercialHits > 0)
			{
				score += (subjectCommercialHits * 34) + (bodyCommercialHits * 18);
				signals.Add(subjectCommercialHits > 0 ? "commercial evidence" : "commercial hints");
			}

			if (numericCommercialHits > 0 || hasMoneyAmount)
			{
				score += 55 + (numericCommercialHits * 10);
				signals.Add("pricing evidence");
			}

			if ((subjectCommercialHits + bodyCommercialHits) == 0)
			{
				score -= queryProfile.AutonomousIntent && autonomousHits > 0 ? 210 : 120;
			}

			if (queryProfile.AutonomousIntent && autonomousHits > 0 && (subjectCommercialHits + bodyCommercialHits) == 0)
			{
				score -= 110;
			}

			if (genericNoiseHits > 0 && (subjectCommercialHits + bodyCommercialHits + numericCommercialHits) == 0)
			{
				score -= 65;
				signals.Add("topic-only");
			}
		}

		var ageDays = (DateTimeOffset.UtcNow - ParseReceivedUtcOrMin(email)).TotalDays;
		if (ageDays <= 14)
		{
			score += 6;
			signals.Add("recent");
		}
		else if (ageDays <= 45)
		{
			score += 3;
			signals.Add("fairly recent");
		}

		email.SearchScore = score;
		email.SearchSignals = signals
			.Distinct(StringComparer.OrdinalIgnoreCase)
			.Take(4)
			.ToList();
		return email;
	}

	private static IEnumerable<string> ExtractCommitmentLines(OutlookSearchCorpusEmail email)
	{
		var source = string.IsNullOrWhiteSpace(email.BodyText) ? email.BodyPreview : email.BodyText;
		if (string.IsNullOrWhiteSpace(source))
		{
			yield break;
		}

		foreach (var sentence in SentenceBoundaryRegex.Split(source))
		{
			var normalized = CollapseWhitespace(sentence);
			if (normalized.Length < 18 || normalized.Length > 220)
			{
				continue;
			}

			var lower = normalized.ToLowerInvariant();
			if (!CommitmentTerms.Any(term => lower.Contains(term, StringComparison.Ordinal)))
			{
				continue;
			}

			yield return normalized;
		}
	}

	private static int CountCommitmentSignals(string value)
	{
		var normalized = NormalizeText(value);
		return CommitmentTerms.Count(term => normalized.Contains(term, StringComparison.Ordinal));
	}

	private static int CountDecisionSignals(string value)
	{
		var normalized = NormalizeText(value);
		return DecisionTerms.Count(term => normalized.Contains(term, StringComparison.Ordinal));
	}

	private static bool LooksLikeCommitmentQuery(string normalizedQuery, HashSet<string> queryTokens)
	{
		if (CommitmentTerms.Any(term => normalizedQuery.Contains(term, StringComparison.Ordinal)))
		{
			return true;
		}

		return queryTokens.Contains("action") ||
			queryTokens.Contains("actions") ||
			queryTokens.Contains("owner") ||
			queryTokens.Contains("owners") ||
			queryTokens.Contains("follow") ||
			queryTokens.Contains("next") ||
			queryTokens.Contains("promised") ||
			queryTokens.Contains("task") ||
			queryTokens.Contains("tasks");
	}

	private static bool LooksLikeDecisionQuery(string normalizedQuery, HashSet<string> queryTokens)
	{
		if (DecisionTerms.Any(term => normalizedQuery.Contains(term, StringComparison.Ordinal)))
		{
			return true;
		}

		return queryTokens.Contains("decision") ||
			queryTokens.Contains("decisions") ||
			queryTokens.Contains("approve") ||
			queryTokens.Contains("approved") ||
			queryTokens.Contains("confirm") ||
			queryTokens.Contains("confirmed") ||
			queryTokens.Contains("resolved") ||
			queryTokens.Contains("החלטה") ||
			queryTokens.Contains("הוחלט") ||
			queryTokens.Contains("מאושר");
	}

	private static bool LooksLikeCommercialQuery(string normalizedQuery, HashSet<string> queryTokens)
	{
		if (CommercialEvidenceTerms.Any(term => normalizedQuery.Contains(term, StringComparison.Ordinal)))
		{
			return true;
		}

		return queryTokens.Contains("pricing") ||
			queryTokens.Contains("price") ||
			queryTokens.Contains("prices") ||
			queryTokens.Contains("cost") ||
			queryTokens.Contains("quote") ||
			queryTokens.Contains("quotes") ||
			queryTokens.Contains("proposal") ||
			queryTokens.Contains("budget") ||
			queryTokens.Contains("estimate") ||
			queryTokens.Contains("estimates") ||
			queryTokens.Contains("מחיר") ||
			queryTokens.Contains("תמחור") ||
			queryTokens.Contains("אומדן");
	}

	private static bool LooksLikeAutonomousQuery(string normalizedQuery, HashSet<string> queryTokens)
	{
		return SearchConcepts
			.First(concept => concept.Name == "autonomous")
			.Terms.Any(term => normalizedQuery.Contains(NormalizeText(term), StringComparison.Ordinal)) ||
			queryTokens.Contains("autonomous") ||
			queryTokens.Contains("adb") ||
			queryTokens.Contains("lakehouse") ||
			queryTokens.Contains("אוטונומ") ||
			queryTokens.Contains("לקהאוס");
	}

	private static HashSet<string> Tokenize(string value)
	{
		var tokens = TokenRegex
			.Matches(value ?? string.Empty)
			.Select(match => match.Value.Trim().ToLowerInvariant())
			.Where(token => token.Length >= 2 && !StopWords.Contains(token))
			.ToHashSet(StringComparer.OrdinalIgnoreCase);

		return tokens;
	}

	private static string NormalizeText(string value) =>
		CollapseWhitespace(value).ToLowerInvariant();

	private static string StripReplyPrefixes(string value)
	{
		var current = value ?? string.Empty;
		while (current.StartsWith("re:", StringComparison.OrdinalIgnoreCase) ||
			current.StartsWith("fw:", StringComparison.OrdinalIgnoreCase) ||
			current.StartsWith("fwd:", StringComparison.OrdinalIgnoreCase))
		{
			var separator = current.IndexOf(':');
			if (separator < 0 || separator == current.Length - 1)
			{
				break;
			}

			current = current[(separator + 1)..].Trim();
		}

		return current;
	}

	private static SearchQueryProfile BuildQueryProfile(string query)
	{
		var normalizedQuery = NormalizeText(query);
		var tokens = Tokenize(query);
		var concepts = SearchConcepts
			.Select(concept => new
			{
				Concept = concept,
				QueryTerms = concept.Terms
					.Where(term => normalizedQuery.Contains(NormalizeText(term), StringComparison.Ordinal))
					.Distinct(StringComparer.OrdinalIgnoreCase)
					.ToArray()
			})
			.Where(entry => entry.QueryTerms.Length > 0)
			.Select(concept => new QueryConceptMatch(
				concept.Concept.Name,
				concept.Concept.Terms,
				concept.QueryTerms))
			.ToList();

		return new SearchQueryProfile(
			normalizedQuery,
			tokens,
			concepts,
			LooksLikeCommitmentQuery(normalizedQuery, tokens),
			LooksLikeDecisionQuery(normalizedQuery, tokens),
			LooksLikeCommercialQuery(normalizedQuery, tokens),
			LooksLikeAutonomousQuery(normalizedQuery, tokens));
	}

	private static int CountMatches(string value, IReadOnlyCollection<string> terms)
	{
		if (string.IsNullOrWhiteSpace(value) || terms.Count == 0)
		{
			return 0;
		}

		var normalized = NormalizeText(value);
		return terms.Count(term => normalized.Contains(NormalizeText(term), StringComparison.Ordinal));
	}

	private static double GetMissingConceptPenalty(QueryConceptMatch conceptMatch) =>
		conceptMatch.Name switch
		{
			"pricing" => 260,
			"autonomous" => 140,
			_ => conceptMatch.IsHighIntent ? 110 : 12
		};

	private static string NormalizeScope(string? value)
	{
		var normalized = (value ?? string.Empty).Trim().ToLowerInvariant();
		return normalized switch
		{
			"inbox" => "inbox",
			"sent-items" => "sent-items",
			_ => "all-mail"
		};
	}

	private static string NormalizeFocus(string? value)
	{
		var normalized = (value ?? string.Empty).Trim().ToLowerInvariant();
		return normalized switch
		{
			"commitments" => "commitments",
			"decisions" => "decisions",
			_ => "all"
		};
	}

	private static bool MatchesScope(OutlookSearchCorpusEmail email, string scope)
	{
		if (scope == "all-mail")
		{
			return true;
		}

		var normalizedFolder = NormalizeText(email.FolderName);
		return scope switch
		{
			"inbox" => normalizedFolder.Contains("inbox", StringComparison.Ordinal),
			"sent-items" => normalizedFolder.Contains("sent", StringComparison.Ordinal),
			_ => true
		};
	}

	private static bool MatchesSender(OutlookSearchCorpusEmail email, string sender)
	{
		if (string.IsNullOrWhiteSpace(sender))
		{
			return true;
		}

		return string.Equals(GetSenderKey(email), sender, StringComparison.OrdinalIgnoreCase);
	}

	private static bool MatchesDate(OutlookSearchCorpusEmail email, DateTimeOffset? startDate)
	{
		if (startDate is null)
		{
			return true;
		}

		return ParseReceivedUtcOrMin(email) >= startDate.Value;
	}

	private static bool MatchesFocus(OutlookSearchCorpusEmail email, string focus)
	{
		if (focus == "all")
		{
			return true;
		}

		var source = string.IsNullOrWhiteSpace(email.BodyText) ? email.BodyPreview : email.BodyText;
		return focus switch
		{
			"commitments" => CountCommitmentSignals(source) > 0,
			"decisions" => CountDecisionSignals(source) > 0,
			_ => true
		};
	}

	private static DateTimeOffset? ResolveStartDate(string? range)
	{
		var normalized = (range ?? string.Empty).Trim().ToLowerInvariant();
		return normalized switch
		{
			"30d" => DateTimeOffset.UtcNow.AddDays(-30),
			"90d" => DateTimeOffset.UtcNow.AddDays(-90),
			"180d" => DateTimeOffset.UtcNow.AddDays(-180),
			"365d" => DateTimeOffset.UtcNow.AddDays(-365),
			_ => null
		};
	}

	private static string CollapseWhitespace(string value) =>
		string.Join(" ", (value ?? string.Empty)
			.Replace("\r", " ")
			.Replace("\n", " ")
			.Split(' ', StringSplitOptions.RemoveEmptyEntries));

	private static DateTimeOffset ParseReceivedUtcOrMin(OutlookSearchCorpusEmail email)
	{
		if (DateTimeOffset.TryParse(email.ReceivedDateTime, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var parsed))
		{
			return parsed;
		}

		var raw = email.ReceivedDateTime ?? string.Empty;
		var legacyMatch = Regex.Match(raw, @"^/Date\((?<millis>-?\d+)(?:[+-]\d+)?\)/$", RegexOptions.CultureInvariant);
		if (legacyMatch.Success &&
			long.TryParse(legacyMatch.Groups["millis"].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var millis))
		{
			try
			{
				return DateTimeOffset.FromUnixTimeMilliseconds(millis);
			}
			catch
			{
			}
		}

		return DateTimeOffset.MinValue;
	}

	private static string FormatEmailReference(OutlookSearchCorpusEmail email)
	{
		var subject = string.IsNullOrWhiteSpace(email.Subject) ? "(No subject)" : email.Subject.Trim();
		var sender = string.IsNullOrWhiteSpace(email.FromDisplayName) ? email.FromAddress : email.FromDisplayName;
		return $"{subject} from {sender}".Trim();
	}

	private static string GetSenderKey(OutlookSearchCorpusEmail email)
	{
		if (!string.IsNullOrWhiteSpace(email.FromAddress))
		{
			return email.FromAddress.Trim();
		}

		return email.FromDisplayName.Trim();
	}

	private sealed record SearchConcept(string Name, string[] Terms);

	private sealed record QueryConceptMatch(string Name, string[] Terms, string[] QueryTerms)
	{
		public bool IsHighIntent => Name is "pricing" or "autonomous";

		public string SignalLabel => Name switch
		{
			"pricing" => "pricing terms",
			"autonomous" => "autonomous terms",
			"commitments" => "commitment terms",
			"decisions" => "decision terms",
			_ => "query concept"
		};
	}

	private sealed record SearchQueryProfile(
		string NormalizedQuery,
		HashSet<string> Tokens,
		List<QueryConceptMatch> Concepts,
		bool CommitmentQuery,
		bool DecisionQuery,
		bool CommercialIntent,
		bool AutonomousIntent);
}
