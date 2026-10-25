using System.Text.RegularExpressions;
using Maily.Generation.Models;

namespace Maily.Generation.Services;

public static class ReplySanitizer
{
	private static readonly Regex PlaceholderGreetingRegex = new(
		@"^\s*(?:hi|hello|dear)\s+\{[^}]+\}\s*,?\s*$",
		RegexOptions.IgnoreCase | RegexOptions.Compiled);

	private static readonly Regex PlaceholderTokenRegex = new(
		@"\{[^}\r\n]{1,40}\}",
		RegexOptions.Compiled);

	private static readonly Regex GreetingLineRegex = new(
		@"^\s*(?:hi|hello|dear)(?:\s+[\p{L}][\p{L}.'’\-]*(?:\s+[\p{L}][\p{L}.'’\-]*){0,3})?\s*,\s*$",
		RegexOptions.IgnoreCase | RegexOptions.Compiled);

	private static readonly Regex SignatureDetailRegex = new(
		@"(?:@|https?://|www\.|\+?\d[\d\s\-()]{6,}|\b(?:cto|ceo|cfo|coo|vp|svp|director|manager|lead|architect|engineer|consultant|analyst|specialist|assistant|coordinator|president|founder|partner|office|mobile|phone|tel|fax)\b|\|)",
		RegexOptions.IgnoreCase | RegexOptions.Compiled);

	private static readonly Regex SignatureContextRegex = new(
		@"\b(?:team|department|dept|group|company|inc|llc|ltd|corp|sales|success|support|operations|marketing|finance|legal|hr)\b",
		RegexOptions.IgnoreCase | RegexOptions.Compiled);

	private static readonly string[] ClosingLines =
	[
		"thanks",
		"thanks",
		"thank you",
		"thank you",
		"best",
		"best regards",
		"regards",
		"kind regards",
		"sincerely",
		"cheers",
		"warm regards"
	];

	public static string NormalizeForOutlookReply(string text, AddinProfileSettings settings, string? currentDraft = null)
	{
		if (string.IsNullOrWhiteSpace(text))
		{
			return string.Empty;
		}

		var normalized = text.Replace("\r\n", "\n").Trim();
		var lines = normalized.Split('\n').ToList();
		RemoveTemplateArtifacts(lines);
		RemoveUnexpectedGreeting(lines, currentDraft);
		TrimTrailingBlankLines(lines);

		var removed = true;
		while (removed && lines.Count > 0)
		{
			removed = false;
			if (TryRemoveTrailingSignature(lines, settings))
			{
				removed = true;
				TrimTrailingBlankLines(lines);
			}
		}

		var result = string.Join("\n", lines).Trim();
		return Regex.Replace(result, "\n{3,}", "\n\n");
	}

	private static void RemoveUnexpectedGreeting(List<string> lines, string? currentDraft)
	{
		if (lines.Count == 0 || DraftHasGreeting(currentDraft))
		{
			return;
		}

		if (GreetingLineRegex.IsMatch(lines[0]))
		{
			lines.RemoveAt(0);
			while (lines.Count > 0 && string.IsNullOrWhiteSpace(lines[0]))
			{
				lines.RemoveAt(0);
			}
		}
	}

	private static void RemoveTemplateArtifacts(List<string> lines)
	{
		if (lines.Count == 0)
		{
			return;
		}

		for (var index = lines.Count - 1; index >= 0; index--)
		{
			lines[index] = PlaceholderTokenRegex.Replace(lines[index], string.Empty).TrimEnd();
		}

		while (lines.Count > 0 && PlaceholderGreetingRegex.IsMatch(lines[0]))
		{
			lines.RemoveAt(0);
		}

		TrimTrailingBlankLines(lines);
	}

	private static bool TryRemoveTrailingSignature(List<string> lines, AddinProfileSettings settings)
	{
		if (lines.Count == 0)
		{
			return false;
		}

		var signatureStart = FindTrailingSignatureStart(lines, settings);
		if (signatureStart < 0)
		{
			return false;
		}

		lines.RemoveRange(signatureStart, lines.Count - signatureStart);
		return true;
	}

	private static int FindTrailingSignatureStart(List<string> lines, AddinProfileSettings settings)
	{
		var userName = NormalizeLine(settings.UserName);
		var firstName = NormalizeLine((settings.UserName ?? string.Empty).Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? string.Empty);
		bool IsName(string line) => !string.IsNullOrWhiteSpace(userName) &&
			(line == userName || line.StartsWith(userName + " |", StringComparison.Ordinal) ||
			 (!string.IsNullOrWhiteSpace(firstName) && line == firstName));
		bool IsFooterLine(string raw)
		{
			var value = NormalizeLine(raw);
			if (value.Length == 0 || IsName(value) || ClosingLines.Contains(value, StringComparer.OrdinalIgnoreCase)) return true;
			return raw.Split(' ', StringSplitOptions.RemoveEmptyEntries).Length <= 10 &&
				!LooksLikeSentence(raw) && (SignatureDetailRegex.IsMatch(raw) || SignatureContextRegex.IsMatch(raw));
		}

		// Contact/date/title tokens inside a sentence are not a signature. Only
		// remove a trailing block anchored by the sender name or a closing line;
		// never consume substantive lines preceding that anchor.
		for (var index = 0; index < lines.Count; index++)
		{
			var value = NormalizeLine(lines[index]);
			var nameAnchor = IsName(value);
			var closingAnchor = ClosingLines.Contains(value, StringComparer.OrdinalIgnoreCase);
			if (!nameAnchor && !closingAnchor) continue;
			var suffix = lines.Skip(index + 1).ToArray();
			if (!nameAnchor && !suffix.Any(line => IsName(NormalizeLine(line)))) continue;
			if (suffix.All(IsFooterLine)) return index;
		}
		return -1;
	}

	private static string NormalizeLine(string value) =>
		(value ?? string.Empty).Trim().TrimEnd('.', ',', ';', ':').Trim().ToLowerInvariant();

	private static bool DraftHasGreeting(string? currentDraft)
	{
		if (string.IsNullOrWhiteSpace(currentDraft))
		{
			return false;
		}

		var firstLine = currentDraft
			.Replace("\r\n", "\n")
			.Split('\n', StringSplitOptions.RemoveEmptyEntries)
			.FirstOrDefault()?
			.Trim();

		return !string.IsNullOrWhiteSpace(firstLine) && Regex.IsMatch(firstLine, @"^\s*(?:hi|hello|dear)\b", RegexOptions.IgnoreCase);
	}

	private static bool LooksLikeSentence(string value)
	{
		var trimmed = value.Trim();
		if (string.IsNullOrWhiteSpace(trimmed))
		{
			return false;
		}

		return trimmed.EndsWith(".", StringComparison.Ordinal) ||
			trimmed.EndsWith("!", StringComparison.Ordinal) ||
			trimmed.EndsWith("?", StringComparison.Ordinal);
	}

	private static void TrimTrailingBlankLines(List<string> lines)
	{
		while (lines.Count > 0 && string.IsNullOrWhiteSpace(lines[^1]))
		{
			lines.RemoveAt(lines.Count - 1);
		}
	}
}
