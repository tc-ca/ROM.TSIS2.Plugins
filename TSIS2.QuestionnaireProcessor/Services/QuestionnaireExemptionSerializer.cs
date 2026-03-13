using System;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TSIS2.Plugins.QuestionnaireProcessor
{
    /// <summary>
    /// Converts the survey exemption payload into the compact JSON stored on ts_questionresponse.
    /// </summary>
    public static class QuestionnaireExemptionSerializer
    {
        public static string SerializeCompact(JToken rawExemptions, ILoggingService logger = null)
        {
            if (rawExemptions == null || rawExemptions.Type == JTokenType.Null)
            {
                return null;
            }

            if (rawExemptions.Type != JTokenType.Array)
            {
                logger?.Warning($"Expected exemption array but received {rawExemptions.Type}. Skipping exemption serialization.");
                return null;
            }

            var compactEntries = new JArray();

            foreach (var item in rawExemptions.Children())
            {
                if (item.Type != JTokenType.Object)
                {
                    continue;
                }

                var exemption = (JObject)item;
                var rawId = exemption.Value<string>("exemptionId") ?? exemption.Value<string>("id");
                if (string.IsNullOrWhiteSpace(rawId))
                {
                    logger?.Warning("Encountered exemption entry without an id. Skipping.");
                    continue;
                }

                compactEntries.Add(new JObject
                {
                    ["id"] = NormalizeGuid(rawId),
                    ["value"] = GetBooleanValue(exemption["exemptionInvoked"] ?? exemption["value"]),
                    ["comment"] = exemption.Value<string>("exemptionComment") ?? exemption.Value<string>("comment") ?? string.Empty
                });
            }

            return compactEntries.ToString(Formatting.None);
        }

        private static string NormalizeGuid(string value)
        {
            var trimmed = (value ?? string.Empty).Trim().Trim('{', '}');
            if (Guid.TryParse(trimmed, out var parsed))
            {
                return parsed.ToString();
            }

            return trimmed;
        }

        private static bool GetBooleanValue(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null)
            {
                return false;
            }

            if (token.Type == JTokenType.Boolean)
            {
                return token.Value<bool>();
            }

            return bool.TryParse(token.ToString(), out var parsed) && parsed;
        }
    }
}
