using Newtonsoft.Json.Linq;
using TSIS2.Plugins.QuestionnaireProcessor;
using Xunit;

namespace ROMTS_GSRST.Plugins.QuestionnaireProcessor
{
    public class QuestionnaireExemptionSerializerTests
    {
        [Fact]
        public void SerializeCompact_SurveyPayload_ReturnsCompactJson()
        {
            var raw = JArray.Parse(@"[
                {
                    ""exemptionId"": ""{3f9c8d4a-3bde-4c7a-9c7a-1d5a6c8b2f10}"",
                    ""exemptionInvoked"": true,
                    ""exemptionComment"": ""comment1"",
                    ""provisionId"": ""abc""
                },
                {
                    ""exemptionId"": ""7b2e1a6c-6c9e-4d0c-8b75-0c2f8a5e9b41"",
                    ""exemptionInvoked"": false,
                    ""exemptionComment"": ""comment2""
                }
            ]");

            var compact = QuestionnaireExemptionSerializer.SerializeCompact(raw);

            Assert.Equal(@"[{""id"":""3f9c8d4a-3bde-4c7a-9c7a-1d5a6c8b2f10"",""value"":true,""comment"":""comment1""},{""id"":""7b2e1a6c-6c9e-4d0c-8b75-0c2f8a5e9b41"",""value"":false,""comment"":""comment2""}]", compact);
        }

        [Fact]
        public void SerializeCompact_AlreadyCompactPayload_IsStable()
        {
            var raw = JArray.Parse(@"[{""id"":""8e3b7a41-2cde-4b60-9c14-6d5f0a2b7c33"",""value"":true,""comment"":""comment23""}]");

            var compact = QuestionnaireExemptionSerializer.SerializeCompact(raw);

            Assert.Equal(@"[{""id"":""8e3b7a41-2cde-4b60-9c14-6d5f0a2b7c33"",""value"":true,""comment"":""comment23""}]", compact);
        }

        [Fact]
        public void SerializeCompact_NullPayload_ReturnsNull()
        {
            Assert.Null(QuestionnaireExemptionSerializer.SerializeCompact(null));
        }
    }
}
