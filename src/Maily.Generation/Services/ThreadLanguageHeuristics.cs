using System.Net;
using System.Text.RegularExpressions;
namespace Maily.Generation.Services;
public static class ThreadLanguageHeuristics
{
    private static readonly Regex HebrewRegex = new(@"[\u0590-\u05FF]", RegexOptions.Compiled);
    // Samples are ordered by authority: typed reply, latest body, then subject.
    public static string DescribeDominantLanguage(params string?[] samples)
    {
        foreach (var value in samples)
        {
            if (string.IsNullOrWhiteSpace(value)) continue;
            var text = NormalizeSample(value);
            var letters = text.Count(char.IsLetter);
            if (letters == 0) continue;
            if (HebrewRegex.Matches(text).Count > letters / 2.0) return "Hebrew";
            return "the dominant language of the actual message text (identify it directly; Latin script does not imply English)";
        }
        return "the language already used in the current message or thread";
    }
    public static bool ContainsHebrew(params string?[] samples) =>
        samples.Any(sample => !string.IsNullOrWhiteSpace(sample) && HebrewRegex.IsMatch(sample));
    private static string NormalizeSample(string value)
    {
        var text = Regex.Replace(value, @"<(style|script)\b[^>]*>[\s\S]*?</\1\s*>", " ", RegexOptions.IgnoreCase);
        text = Regex.Replace(text, @"<(br\b[^>]*|/p|/div)>", "\n", RegexOptions.IgnoreCase);
        text = WebUtility.HtmlDecode(Regex.Replace(text, @"<[^>]*>", " ")).Replace("\r\n", "\n");
        var lines = new List<string>();
        foreach (var line in text.Split('\n'))
        {
            if (Regex.IsMatch(line, @"^\s*(?:--\s*$|On .+wrote:|-----Original Message-----)", RegexOptions.IgnoreCase)) break;
            if (Regex.IsMatch(line, @"^\s*(?:From|Sent|Subject|To|Cc):", RegexOptions.IgnoreCase))
            {
                if (lines.Any(static item => !string.IsNullOrWhiteSpace(item))) break;
                continue;
            }
            if (!line.TrimStart().StartsWith('>')) lines.Add(line);
        }
        text = string.Join("\n", lines);
        text = Regex.Replace(text, @"\b(?:mailto:|https?://|www\.)\S*|\b[^\s@]+@[^\s@]+\b", " ", RegexOptions.IgnoreCase);
        return Regex.Replace(text, @"\b(?:re|fw|fwd)\s*:", " ", RegexOptions.IgnoreCase).Trim();
    }
}