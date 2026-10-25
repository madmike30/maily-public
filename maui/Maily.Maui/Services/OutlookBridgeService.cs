using System.Diagnostics;
using System.Text.Json;
using Maily.Maui.Models;

namespace Maily.Maui.Services;

public sealed class OutlookBridgeService
{
	private readonly AppPaths appPaths;
	private readonly MailyLogger logger;
	private readonly Lazy<Task<string>> mailboxAddress;

	public OutlookBridgeService(AppPaths appPaths, MailyLogger logger)
	{
		this.appPaths = appPaths;
		this.logger = logger;
		// Keep one explicit account for this service lifetime so cached item IDs
		// cannot silently move to a different account after a settings edit.
		mailboxAddress = new Lazy<Task<string>>(async () =>
		{
			var settings = await new UserSettingsService(appPaths, logger).LoadAsync();
			var address = settings.MailboxAddress?.Trim();
			if (string.IsNullOrWhiteSpace(address))
			{
				throw new InvalidOperationException("Configure mailbox_address in settings/user_settings.json and restart Maily before connecting to Outlook.");
			}
			return address;
		});
	}

	public async Task TestConnectionAsync(CancellationToken cancellationToken = default)
	{
		var timer = Stopwatch.StartNew();
		await InvokeJsonAsync("test", null, cancellationToken);
		timer.Stop();
		logger.Info("outlook_bridge_tested", new { elapsedMs = timer.ElapsedMilliseconds });
	}

	public async Task<IReadOnlyList<OutlookEmail>> ListInboxAsync(int limit, DateTime? startDate = null, DateTime? endDate = null, CancellationToken cancellationToken = default)
	{
		var timer = Stopwatch.StartNew();
		var args = new Dictionary<string, string>
		{
			["Limit"] = limit.ToString()
		};
		if (startDate.HasValue)
		{
			args["StartDate"] = startDate.Value.ToString("yyyy-MM-dd");
		}

		if (endDate.HasValue)
		{
			args["EndDate"] = endDate.Value.ToString("yyyy-MM-dd");
		}

		var payload = await InvokeJsonAsync("list-emails-preview", args, cancellationToken);

		var results = new List<OutlookEmail>();
		foreach (var item in payload.RootElement.EnumerateArray())
		{
			results.Add(MapEmail(item, includeBody: false, Array.Empty<OutlookAttachmentInfo>()));
		}

		timer.Stop();
		logger.Info("outlook_inbox_loaded", new
		{
			count = results.Count,
			limit,
			startDate = startDate?.ToString("yyyy-MM-dd"),
			endDate = endDate?.ToString("yyyy-MM-dd"),
			elapsedMs = timer.ElapsedMilliseconds
		});
		return results;
	}

	public async Task<OutlookEmail> GetEmailBodyAsync(string mailId, CancellationToken cancellationToken = default)
	{
		var timer = Stopwatch.StartNew();
		var payload = await InvokeJsonAsync("get-email", new Dictionary<string, string> { ["MailId"] = mailId }, cancellationToken);
		var email = MapEmail(payload.RootElement, includeBody: true, Array.Empty<OutlookAttachmentInfo>());
		timer.Stop();
		logger.Info("outlook_email_body_loaded", new { id = mailId, elapsedMs = timer.ElapsedMilliseconds });
		return email;
	}

	public async Task<IReadOnlyList<OutlookAttachmentInfo>> GetAttachmentsAsync(string mailId, CancellationToken cancellationToken = default)
	{
		var timer = Stopwatch.StartNew();
		var attachmentsPayload = await InvokeJsonAsync("list-attachments", new Dictionary<string, string> { ["MailId"] = mailId }, cancellationToken);
		var attachments = attachmentsPayload.RootElement.ValueKind == JsonValueKind.Array
			? attachmentsPayload.RootElement.EnumerateArray().Select(MapAttachment).ToArray()
			: Array.Empty<OutlookAttachmentInfo>();
		timer.Stop();
		logger.Info("outlook_email_attachments_loaded", new { id = mailId, attachmentCount = attachments.Length, elapsedMs = timer.ElapsedMilliseconds });
		return attachments;
	}

	public async Task<OutlookEmail> GetEmailAsync(string mailId, CancellationToken cancellationToken = default)
	{
		var bodyTask = GetEmailBodyAsync(mailId, cancellationToken);
		var attachmentsTask = GetAttachmentsAsync(mailId, cancellationToken);
		var body = await bodyTask;
		var attachments = await attachmentsTask;
		return body.WithAttachments(attachments);
	}

	public async Task CreateReplyDraftAsync(string mailId, string htmlBody, bool replyAll, CancellationToken cancellationToken = default)
	{
		var timer = Stopwatch.StartNew();
		await mailboxAddress.Value.WaitAsync(cancellationToken);
		var tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".html");
		await File.WriteAllTextAsync(tempPath, htmlBody, cancellationToken);

		try
		{
			await InvokeJsonAsync(replyAll ? "create-reply-all" : "create-reply", new Dictionary<string, string>
			{
				["MailId"] = mailId,
				["HtmlPath"] = tempPath
			}, cancellationToken);

			timer.Stop();
			logger.Info("outlook_reply_draft_created", new { mailId, replyAll, elapsedMs = timer.ElapsedMilliseconds });
		}
		finally
		{
			File.Delete(tempPath);
		}
	}

	private async Task<JsonDocument> InvokeJsonAsync(string action, IReadOnlyDictionary<string, string>? args, CancellationToken cancellationToken)
	{
		var address = await mailboxAddress.Value.WaitAsync(cancellationToken);
		var outputPath = Path.GetTempFileName();
		try
		{
			var startInfo = new ProcessStartInfo
			{
				FileName = "powershell",
				WorkingDirectory = appPaths.RepoRoot,
				UseShellExecute = false,
				CreateNoWindow = true,
				RedirectStandardError = true
			};

			startInfo.ArgumentList.Add("-NoProfile");
			startInfo.ArgumentList.Add("-NonInteractive");
			startInfo.ArgumentList.Add("-ExecutionPolicy");
			startInfo.ArgumentList.Add("Bypass");
			startInfo.ArgumentList.Add("-File");
			startInfo.ArgumentList.Add(appPaths.OutlookBridgeScript);
			startInfo.ArgumentList.Add("-Action");
			startInfo.ArgumentList.Add(action);
			startInfo.ArgumentList.Add("-MailboxAddress");
			startInfo.ArgumentList.Add(address);
			startInfo.ArgumentList.Add("-OutputPath");
			startInfo.ArgumentList.Add(outputPath);

			if (args is not null)
			{
				foreach (var pair in args)
				{
					startInfo.ArgumentList.Add("-" + pair.Key);
					startInfo.ArgumentList.Add(pair.Value);
				}
			}

			using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Could not start Outlook bridge.");
			await process.WaitForExitAsync(cancellationToken);
			var stdErr = await process.StandardError.ReadToEndAsync(cancellationToken);
			if (process.ExitCode != 0)
			{
				throw new InvalidOperationException(string.IsNullOrWhiteSpace(stdErr) ? $"Outlook bridge action '{action}' failed." : stdErr.Trim());
			}

			var content = await File.ReadAllTextAsync(outputPath, cancellationToken);
			if (string.IsNullOrWhiteSpace(content))
			{
				content = "{}";
			}

			return JsonDocument.Parse(content);
		}
		catch (Exception ex)
		{
			logger.Error("outlook_bridge_failed", ex, new { action });
			throw;
		}
		finally
		{
			File.Delete(outputPath);
		}
	}

	private static OutlookAttachmentInfo MapAttachment(JsonElement element) =>
		new()
		{
			Id = element.GetProperty("id").GetString() ?? string.Empty,
			Name = element.GetProperty("name").GetString() ?? string.Empty,
			Index = element.GetProperty("index").GetInt32(),
			Size = element.TryGetProperty("size", out var size) ? size.GetInt64() : 0
		};

	private static OutlookEmail MapEmail(JsonElement element, bool includeBody, IReadOnlyList<OutlookAttachmentInfo> attachments)
	{
		var from = element.GetProperty("from").GetProperty("emailAddress");
		var senderName = from.TryGetProperty("name", out var name) ? name.GetString() ?? string.Empty : string.Empty;
		var senderAddress = from.TryGetProperty("address", out var address) ? address.GetString() ?? string.Empty : string.Empty;
		var received = element.TryGetProperty("receivedDateTime", out var receivedProp) &&
			receivedProp.ValueKind == JsonValueKind.String &&
			DateTime.TryParse(receivedProp.GetString(), out var parsedReceived)
				? parsedReceived
				: DateTime.MinValue;
		var importance = element.TryGetProperty("importance", out var importanceProp) ? importanceProp.GetString() ?? "normal" : "normal";
		var body = element.GetProperty("body");

		var category = "General";
		if (element.TryGetProperty("categories", out var categories) &&
			categories.ValueKind == JsonValueKind.Array &&
			categories.GetArrayLength() > 0)
		{
			category = categories[0].GetString() ?? "General";
		}

		var previewText = element.TryGetProperty("bodyPreview", out var preview)
			? (preview.GetString() ?? string.Empty)
			: string.Empty;
		if (string.IsNullOrWhiteSpace(previewText))
		{
			previewText = element.TryGetProperty("subject", out var subjectPreview)
				? subjectPreview.GetString() ?? string.Empty
				: string.Empty;
		}

		return new OutlookEmail
		{
			Id = element.GetProperty("id").GetString() ?? string.Empty,
			Subject = element.TryGetProperty("subject", out var subject) ? subject.GetString() ?? "(No subject)" : "(No subject)",
			SenderName = string.IsNullOrWhiteSpace(senderName) ? senderAddress : senderName,
			SenderAddress = senderAddress,
			SenderLine = string.IsNullOrWhiteSpace(senderAddress) ? senderName : $"{senderName} · {senderAddress}",
			Importance = importance,
			ReceivedAgo = received == DateTime.MinValue ? string.Empty : FormatRelative(received),
			ReceivedFull = received == DateTime.MinValue ? string.Empty : received.ToString("f"),
			Preview = previewText,
			BodyText = includeBody && body.TryGetProperty("text", out var text) ? text.GetString() ?? string.Empty : string.Empty,
			BodyHtml = includeBody && body.TryGetProperty("content", out var content) ? content.GetString() ?? string.Empty : string.Empty,
			Category = category,
			AccountName = GetAccountName(senderAddress),
			TopicSummary = string.IsNullOrWhiteSpace(category) ? "Email thread" : $"{category} thread",
			PriorityLabel = importance switch
			{
				"high" => "High priority",
				"low" => "Low priority",
				_ => "Standard priority"
			},
			IsUnread = element.TryGetProperty("isRead", out var isRead) && isRead.ValueKind == JsonValueKind.False,
			HasAttachments = element.TryGetProperty("hasAttachments", out var hasAttachments) && hasAttachments.GetBoolean(),
			Attachments = attachments
		};
	}

	private static string FormatRelative(DateTime timestamp)
	{
		var delta = DateTime.Now - timestamp;
		if (delta.TotalMinutes < 60) return $"{Math.Max(1, (int)delta.TotalMinutes)}m";
		if (delta.TotalHours < 24) return $"{Math.Max(1, (int)delta.TotalHours)}h";
		return $"{Math.Max(1, (int)delta.TotalDays)}d";
	}

	private static string GetAccountName(string senderAddress)
	{
		if (string.IsNullOrWhiteSpace(senderAddress) || !senderAddress.Contains('@'))
		{
			return "Mail";
		}

		var domain = senderAddress.Split('@')[1];
		var label = domain.Split('.')[0];
		return string.IsNullOrWhiteSpace(label) ? "Mail" : char.ToUpperInvariant(label[0]) + label[1..];
	}
}
