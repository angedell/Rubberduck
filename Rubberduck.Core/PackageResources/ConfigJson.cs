using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Rubberduck.Settings;

namespace Rubberduck.PackageResources
{
    /// <summary>Checks the text of a <c>rubberduck.json</c> that is being edited.</summary>
    public static class ConfigJson
    {
        /// <summary>
        /// Malformed JSON is a blocking problem. A settings problem (unknown key, bad value) is only a warning: Rubberduck ignores what it cannot read,
        /// and the file may be meant for a newer version. Null if the text is fine.
        /// </summary>
        public static TextProblem Check(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return new TextProblem("The config is empty.");
            }

            try
            {
                if (!(JToken.Parse(json) is JObject))
                {
                    return new TextProblem("The config must be a JSON object.", 1, 1);
                }
            }
            catch (JsonReaderException exception)
            {
                return new TextProblem(exception.Message, exception.LineNumber, exception.LinePosition);
            }
            catch (JsonException exception)
            {
                return new TextProblem(exception.Message);
            }

            var errors = new ProjectConfigReader(null).Parse(json, null).Errors;
            return errors.Any() ? new TextProblem(string.Join(" ", errors), 0, 0, false) : null;
        }

        /// <summary>The JSON indented by two spaces, in the order it was written. Throws <see cref="JsonException"/> if it is not valid JSON or has comments.</summary>
        public static string Format(string json) => Write(json, Formatting.Indented);

        /// <summary>The JSON on a single line. Throws <see cref="JsonException"/> if it is not valid JSON.</summary>
        public static string Linearize(string json) => Write(json, Formatting.None);

        private static string Write(string json, Formatting formatting)
        {
            //The JSON model does not keep comments reliably (those before a property are lost), and a layout must not delete what the user wrote.
            using (var reader = new JsonTextReader(new System.IO.StringReader(json)))
            {
                while (reader.Read())
                {
                    if (reader.TokenType == JsonToken.Comment)
                    {
                        throw new JsonReaderException("The text has comments, which would be lost. Remove them first.", reader.Path, reader.LineNumber, reader.LinePosition, null);
                    }
                }
            }

            var token = JToken.Parse(json, new JsonLoadSettings { LineInfoHandling = LineInfoHandling.Ignore });
            var builder = new System.Text.StringBuilder();
            using (var writer = new JsonTextWriter(new System.IO.StringWriter(builder)) { Formatting = formatting, Indentation = 2, IndentChar = ' ' })
            {
                token.WriteTo(writer);
            }

            return builder.ToString().Replace("\r\n", "\n").Replace("\n", "\r\n");
        }
    }
}
