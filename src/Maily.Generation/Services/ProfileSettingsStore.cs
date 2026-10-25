using System.Text.Json;
using Maily.Generation.Contracts;
using Maily.Generation.Models;

namespace Maily.Generation.Services;

public sealed class ProfileSettingsStore
{
	private readonly string filePath;
	private readonly object gate = new();
	private static readonly JsonSerializerOptions JsonOptions = new()
	{
		PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
		WriteIndented = true
	};

	public ProfileSettingsStore(PathResolver paths)
	{
		filePath = paths.SettingsFilePath;
	}

	public AddinSettingsResponse Get(string mailboxAddress, string profileId)
	{
		lock (gate)
		{
			var store = LoadUnsafe();
			var key = BuildKey(mailboxAddress, profileId);
			if (!store.Items.TryGetValue(key, out var item))
			{
				return new AddinSettingsResponse
				{
					MailboxAddress = mailboxAddress,
					ProfileId = profileId,
					Settings = new AddinProfileSettings(),
					UpdatedUtc = DateTimeOffset.UtcNow
				};
			}

			return new AddinSettingsResponse
			{
				MailboxAddress = item.MailboxAddress,
				ProfileId = item.ProfileId,
				Settings = item.Settings,
				UpdatedUtc = item.UpdatedUtc
			};
		}
	}

	public AddinSettingsResponse Upsert(string mailboxAddress, string profileId, AddinProfileSettings settings)
	{
		lock (gate)
		{
			var store = LoadUnsafe();
			var key = BuildKey(mailboxAddress, profileId);
			var item = new StoredProfile
			{
				MailboxAddress = mailboxAddress,
				ProfileId = profileId,
				Settings = settings,
				UpdatedUtc = DateTimeOffset.UtcNow
			};
			store.Items[key] = item;
			SaveUnsafe(store);
			return new AddinSettingsResponse
			{
				MailboxAddress = item.MailboxAddress,
				ProfileId = item.ProfileId,
				Settings = item.Settings,
				UpdatedUtc = item.UpdatedUtc
			};
		}
	}

	private SettingsStore LoadUnsafe()
	{
		if (!File.Exists(filePath))
		{
			return new SettingsStore();
		}

		var json = File.ReadAllText(filePath);
		return JsonSerializer.Deserialize<SettingsStore>(json, JsonOptions) ?? new SettingsStore();
	}

	private void SaveUnsafe(SettingsStore store)
	{
		var directory = Path.GetDirectoryName(filePath);
		if (!string.IsNullOrWhiteSpace(directory))
		{
			Directory.CreateDirectory(directory);
		}

		var json = JsonSerializer.Serialize(store, JsonOptions);
		File.WriteAllText(filePath, json);
	}

	private static string BuildKey(string mailboxAddress, string profileId) =>
		$"{mailboxAddress.Trim().ToLowerInvariant()}::{profileId.Trim().ToLowerInvariant()}";

	private sealed class SettingsStore
	{
		public Dictionary<string, StoredProfile> Items { get; set; } = new(StringComparer.OrdinalIgnoreCase);
	}

	private sealed class StoredProfile
	{
		public string MailboxAddress { get; set; } = string.Empty;
		public string ProfileId { get; set; } = "default";
		public AddinProfileSettings Settings { get; set; } = new();
		public DateTimeOffset UpdatedUtc { get; set; } = DateTimeOffset.UtcNow;
	}
}
