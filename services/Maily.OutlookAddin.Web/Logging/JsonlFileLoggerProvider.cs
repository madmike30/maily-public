using System.Collections.Concurrent;
using System.Text.Json;
using Microsoft.Extensions.Logging;

namespace Maily.OutlookAddin.Web.Logging;

public sealed class JsonlFileLoggerProvider : ILoggerProvider
{
	private readonly ConcurrentDictionary<string, JsonlFileLogger> loggers = new(StringComparer.OrdinalIgnoreCase);
	private readonly Func<string> pathFactory;
	private readonly object gate = new();
	private bool disposed;

	public JsonlFileLoggerProvider(Func<string> pathFactory)
	{
		this.pathFactory = pathFactory;
	}

	public ILogger CreateLogger(string categoryName)
	{
		ObjectDisposedException.ThrowIf(disposed, this);
		return loggers.GetOrAdd(categoryName, category => new JsonlFileLogger(category, WriteEntry));
	}

	public void Dispose()
	{
		disposed = true;
		loggers.Clear();
	}

	private void WriteEntry(JsonElement payload)
	{
		var path = pathFactory();
		var directory = Path.GetDirectoryName(path);
		if (!string.IsNullOrWhiteSpace(directory))
		{
			Directory.CreateDirectory(directory);
		}

		lock (gate)
		{
			using var stream = new FileStream(
				path,
				FileMode.Append,
				FileAccess.Write,
				FileShare.ReadWrite);
			using var writer = new StreamWriter(stream);
			writer.WriteLine(payload.GetRawText());
		}
	}

	private sealed class JsonlFileLogger : ILogger
	{
		private readonly string categoryName;
		private readonly Action<JsonElement> writeEntry;

		public JsonlFileLogger(string categoryName, Action<JsonElement> writeEntry)
		{
			this.categoryName = categoryName;
			this.writeEntry = writeEntry;
		}

		public IDisposable BeginScope<TState>(TState state) where TState : notnull => NullScope.Instance;

		public bool IsEnabled(LogLevel logLevel) => logLevel != LogLevel.None;

		public void Log<TState>(
			LogLevel logLevel,
			EventId eventId,
			TState state,
			Exception? exception,
			Func<TState, Exception?, string> formatter)
		{
			if (!IsEnabled(logLevel))
			{
				return;
			}

			var message = formatter(state, exception);
			var properties = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);

			if (state is IEnumerable<KeyValuePair<string, object?>> structuredState)
			{
				foreach (var entry in structuredState)
				{
					if (string.Equals(entry.Key, "{OriginalFormat}", StringComparison.Ordinal))
					{
						continue;
					}

					properties[entry.Key] = SimplifyValue(entry.Value);
				}
			}

			var payload = JsonSerializer.SerializeToElement(new
			{
				timestamp = DateTimeOffset.UtcNow,
				level = logLevel.ToString(),
				category = categoryName,
				eventId = eventId.Id,
				message,
				properties,
				exception = exception?.ToString()
			});

			writeEntry(payload);
		}

		private static object? SimplifyValue(object? value)
		{
			if (value is null)
			{
				return null;
			}

			var type = value.GetType();
			if (type.IsPrimitive ||
				value is string ||
				value is decimal ||
				value is DateTime ||
				value is DateTimeOffset ||
				value is TimeSpan ||
				value is Guid ||
				value is Enum)
			{
				return value;
			}

			try
			{
				return Convert.ToString(value);
			}
			catch
			{
				return type.FullName ?? "unknown";
			}
		}
	}

	private sealed class NullScope : IDisposable
	{
		public static readonly NullScope Instance = new();

		public void Dispose()
		{
		}
	}
}
