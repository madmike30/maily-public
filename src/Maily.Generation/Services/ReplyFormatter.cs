using System.Net;

namespace Maily.Generation.Services;

public static class ReplyFormatter
{
	public static string ToOutlookHtml(string text)
	{
		var normalized = (text ?? string.Empty).Replace("\r\n", "\n");
		var isHebrew = ThreadLanguageHeuristics.ContainsHebrew(normalized);
		var outerDirection = isHebrew ? "rtl" : "ltr";
		var textAlign = isHebrew ? "right" : "left";
		var lines = normalized.Split('\n');
		var htmlLines = lines.Select(line =>
		{
			var content = string.IsNullOrEmpty(line) ? "<br/>" : WebUtility.HtmlEncode(line);
			return $"<div dir='{outerDirection}' style='margin:0; line-height:1.4; text-align:{textAlign};'>{content}</div>";
		});

		return $"<!-- MAILY_REPLY_START --><div dir='{outerDirection}' style='margin:0; font-size:11pt; font-family:Aptos,Calibri,sans-serif; text-align:{textAlign};'>{string.Join(string.Empty, htmlLines)}</div><!-- MAILY_REPLY_END -->";
	}
}
