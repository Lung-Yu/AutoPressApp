using System;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace AutoPressApp.Steps
{
    public class StepJsonConverter : JsonConverter<Step>
    {
        public override Step Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            using var doc = JsonDocument.ParseValue(ref reader);
            if (!doc.RootElement.TryGetProperty("$type", out var typeProp))
                throw new JsonException("Missing '$type' discriminator for Step");
            var disc = typeProp.GetString()?.ToLowerInvariant();
            Type concrete = disc switch
            {
                "focuswindow" => typeof(FocusWindowStep),
                "mouseclick" => typeof(MouseClickStep),
                "delay" => typeof(DelayStep),
                "log" => typeof(LogStep),
                "keycombo" => typeof(KeyComboStep),
                "keysequence" => typeof(KeySequenceStep),
                _ => throw new JsonException($"Unknown step type '{disc}'")
            };
            string json = doc.RootElement.GetRawText();
            return (Step)(JsonSerializer.Deserialize(json, concrete, options) ?? throw new JsonException("Deserialize step failed"));
        }

        public override void Write(Utf8JsonWriter writer, Step value, JsonSerializerOptions options)
        {
            // Simple write by adding $type based on runtime type
            string disc = value switch
            {
                FocusWindowStep => "focusWindow",
                MouseClickStep => "mouseClick",
                DelayStep => "delay",
                LogStep => "log",
                KeyComboStep => "keyCombo",
                KeySequenceStep => "keySequence",
                _ => value.GetType().Name
            };
            using var doc = JsonDocument.Parse(JsonSerializer.Serialize(value, value.GetType(), options));
            writer.WriteStartObject();
            writer.WriteString("$type", disc);
            foreach (var prop in doc.RootElement.EnumerateObject())
            {
                prop.WriteTo(writer);
            }
            writer.WriteEndObject();
        }
    }
}
