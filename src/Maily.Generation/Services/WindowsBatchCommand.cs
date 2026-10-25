using System.Text;

namespace Maily.Generation.Services;

public static class WindowsBatchCommand
{
    // cmd parses its command tail itself, so ProcessStartInfo.ArgumentList's
    // CRT escaping must not be applied to the entire /c command.
    public static string BuildArguments(IEnumerable<string> arguments) =>
        "/d /v:off /s /c \"" + string.Join(" ", arguments.Select(QuoteArgument)) + "\"";

    public static string QuoteArgument(string value)
    {
        if (value.Any(char.IsControl) || value.IndexOfAny(['%', '!', '&', '|', '<', '>', '^']) >= 0)
            throw new ArgumentException("Batch CLI paths and arguments cannot contain shell expansion or control characters. Configure a native executable instead.");
        if (value.Length > 0 && !value.Any(char.IsWhiteSpace) && value.IndexOfAny(['\"', '(', ')']) < 0) return value;
        var result = new StringBuilder("\"");
        var slashes = 0;
        foreach (var c in value)
        {
            if (c == '\\') { slashes++; continue; }
            result.Append('\\', c == '\"' ? slashes * 2 + 1 : slashes);
            result.Append(c);
            slashes = 0;
        }
        result.Append('\\', slashes * 2);
        return result.Append('"').ToString();
    }
}
