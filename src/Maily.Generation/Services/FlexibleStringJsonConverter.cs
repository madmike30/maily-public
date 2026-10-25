using System.Text.Json;
using System.Text.Json.Serialization;

namespace Maily.Generation.Services;

public sealed class FlexibleStringJsonConverter : JsonConverter<string>
{
	public override string Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
	{
		using var document = JsonDocument.ParseValue(ref reader);
		return Normalize(document.RootElement);
	}

	public override void Write(Utf8JsonWriter writer, string value, JsonSerializerOptions options)
	{
		writer.WriteStringValue(value);
	}

	public static string Normalize(JsonElement element)
	{
		return (element.ValueKind switch
		{
			JsonValueKind.Null or JsonValueKind.Undefined => string.Empty,
			JsonValueKind.String => element.GetString() ?? string.Empty,
			JsonValueKind.Number => element.ToString(),
			JsonValueKind.True => "true",
			JsonValueKind.False => "false",
			JsonValueKind.Array => NormalizeArray(element),
			JsonValueKind.Object => NormalizeObject(element),
			_ => element.ToString()
		}).Trim();
	}

	private static string NormalizeArray(JsonElement element)
	{
		var items = element.EnumerateArray()
			.Select(Normalize)
			.Where(static value => !string.IsNullOrWhiteSpace(value))
			.ToList();

		if (items.Count == 0)
		{
			return string.Empty;
		}

		if (items.Count == 1)
		{
			return items[0];
		}

		return string.Join(Environment.NewLine, items.Select(static item => $"- {item}"));
	}

	private static string NormalizeObject(JsonElement element)
	{
		var preferredProperty = element.EnumerateObject()
			.FirstOrDefault(static property =>
				property.NameEquals("text") ||
				property.NameEquals("summary") ||
				property.NameEquals("content") ||
				property.NameEquals("message") ||
				property.NameEquals("value"));

		if (!preferredProperty.Equals(default(JsonProperty)))
		{
			return Normalize(preferredProperty.Value);
		}

		var pairs = element.EnumerateObject()
			.Select(property =>
			{
				var value = Normalize(property.Value);
				return string.IsNullOrWhiteSpace(value)
					? string.Empty
					: $"{ToDisplayLabel(property.Name)}: {value}";
			})
			.Where(static value => !string.IsNullOrWhiteSpace(value))
			.ToList();

		return string.Join(Environment.NewLine, pairs);
	}

	private static string ToDisplayLabel(string value)
	{
		if (string.IsNullOrWhiteSpace(value))
		{
			return string.Empty;
		}

		var chars = new List<char>(value.Length + 4);
		for (var index = 0; index < value.Length; index += 1)
		{
			var current = value[index];
			if (index > 0 && char.IsUpper(current) && char.IsLower(value[index - 1]))
			{
				chars.Add(' ');
			}

			chars.Add(index == 0 ? char.ToUpperInvariant(current) : current);
		}

		return new string(chars.ToArray());
	}
}
