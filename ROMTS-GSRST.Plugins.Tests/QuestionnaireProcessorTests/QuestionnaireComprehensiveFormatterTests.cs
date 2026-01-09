using Newtonsoft.Json.Linq;
using TSIS2.Plugins.QuestionnaireProcessor;
using Xunit;

namespace ROMTS_GSRST.Plugins.QuestionnaireProcessor
{
    public class QuestionnaireComprehensiveFormatterTests
    {
        private readonly QuestionnaireResponseFormatter _formatter;

        public QuestionnaireComprehensiveFormatterTests()
        {
            _formatter = new QuestionnaireResponseFormatter(new LoggerAdapter());
        }

        [Theory]
        [InlineData("&ge;", ">=")]
        [InlineData("&le;", "<=")]
        [InlineData("&ne;", "!=")]
        [InlineData("&eacute;", "é")]
        [InlineData("&agrave;", "à")]
        [InlineData("&ccedil;", "ç")]
        [InlineData("&ldquo;Notice of Inspection&rdquo;", "\u201CNotice of Inspection\u201D")] // Smart quotes preserved
        [InlineData("&laquo; Avis d'inspection &raquo;", "« Avis d'inspection »")]
        [InlineData("&ndash; &mdash;", "- -")]
        [InlineData("&plusmn; &times; &divide;", "+/- * /")]
        [InlineData("Line1<br>Line2", "Line1 Line2")]
        [InlineData("<p>Text</p>", "Text")]
        public void RemoveHtmlTags_HandlesSpecificEntities(string input, string expected)
        {
            var result = _formatter.RemoveHtmlTags(input);
            Assert.Equal(expected, result);
        }

        [Fact]
        public void RemoveHtmlTags_WhitespaceOnly_ReturnsEmptyString()
        {
            var result = _formatter.RemoveHtmlTags("&nbsp;");
            Assert.Equal("", result);
        }

        #region Newline and Tab Handling Tests

        [Fact]
        public void RemoveHtmlTags_WithNewlines_ReplacesWithSpaces()
        {
            var result = _formatter.RemoveHtmlTags("Line one\nLine two\nLine three");
            Assert.Equal("Line one Line two Line three", result);
        }

        [Fact]
        public void RemoveHtmlTags_WithCarriageReturnNewlines_ReplacesWithSpaces()
        {
            var result = _formatter.RemoveHtmlTags("Line one\r\nLine two\r\nLine three");
            Assert.Equal("Line one Line two Line three", result);
        }

        [Fact]
        public void RemoveHtmlTags_WithTabs_ReplacesWithSpaces()
        {
            var result = _formatter.RemoveHtmlTags("col1\tcol2\tcol3");
            Assert.Equal("col1 col2 col3", result);
        }

        [Fact]
        public void RemoveHtmlTags_WithEscapedNewlines_ReplacesWithSpaces()
        {
            var result = _formatter.RemoveHtmlTags("Line one\\nLine two\\nLine three");
            Assert.Equal("Line one Line two Line three", result);
        }

        [Fact]
        public void RemoveHtmlTags_WithEscapedTabs_ReplacesWithSpaces()
        {
            var result = _formatter.RemoveHtmlTags("col1\\tcol2\\tcol3");
            Assert.Equal("col1 col2 col3", result);
        }

        [Fact]
        public void RemoveHtmlTags_WithMixedNewlinesAndTabs_ReplacesWithSpaces()
        {
            var result = _formatter.RemoveHtmlTags("Header\n\tcol1\tcol2\nRow1\n\tval1\tval2");
            Assert.Equal("Header col1 col2 Row1 val1 val2", result);
        }

        [Fact]
        public void RemoveHtmlTags_ComplexTextWithOperatorsAndNewlines_PreservesOperators()
        {
            var input = "OPERATORS:\n> 10\n< 5\n>= 3\n<= 7\n!= 2";
            var result = _formatter.RemoveHtmlTags(input);
            Assert.Equal("OPERATORS: > 10 < 5 >= 3 <= 7 != 2", result);
        }

        [Fact]
        public void RemoveHtmlTags_FrenchTextWithNewlines_PreservesAccents()
        {
            var input = "FRENCH:\né è ê ë à â æ ç ô œ ù û ü ÿ\ngarçon façade naïve rôle élève cœur";
            var result = _formatter.RemoveHtmlTags(input);
            Assert.Equal("FRENCH: é è ê ë à â æ ç ô œ ù û ü ÿ garçon façade naïve rôle élève cœur", result);
        }

        [Fact]
        public void RemoveHtmlTags_EmojiWithNewlines_PreservesEmojis()
        {
            var input = "EMOJI:\n😀 😁 🚀";
            var result = _formatter.RemoveHtmlTags(input);
            Assert.Equal("EMOJI: 😀 😁 🚀", result);
        }

        [Fact]
        public void RemoveHtmlTags_HtmlWithNewlines_RemovesTagsAndNewlines()
        {
            var input = "HTML:\n<p>para</p><br>line2\n<span>span text</span>";
            var result = _formatter.RemoveHtmlTags(input);
            Assert.Equal("HTML: para line2 span text", result);
        }

        [Fact]
        public void RemoveHtmlTags_QuotesWithNewlines_PreservesQuotes()
        {
            // Smart quotes are preserved (not converted to regular quotes) to avoid JSON escaping issues
            var input = "QUOTES:\n\"double\"\n'single'\n\u201Csmart double\u201D\n\u2018smart single\u2019";
            var result = _formatter.RemoveHtmlTags(input);
            // Smart quotes remain as smart quotes, regular quotes remain as regular quotes
            Assert.Equal("QUOTES: \"double\" 'single' \u201Csmart double\u201D \u2018smart single\u2019", result);
        }

        [Fact]
        public void RemoveHtmlTags_SmartQuotes_PreservedAsIs()
        {
            // Smart double quotes " " (U+201C, U+201D) are preserved
            // Smart single quotes ' ' (U+2018, U+2019) are preserved
            // This avoids JSON escaping issues when storing in ts_details
            var input = "\u201CSmart quoted text\u201D and \u2018single\u2019";
            var result = _formatter.RemoveHtmlTags(input);
            Assert.Equal("\u201CSmart quoted text\u201D and \u2018single\u2019", result);
        }

        [Fact]
        public void RemoveHtmlTags_SmartQuotesInJson_NotEscaped()
        {
            // Verify that smart quotes don't get escaped when stored in JSON
            var input = "\u201Csmart quotes\u201D 'single quotes'";
            var cleanedText = _formatter.RemoveHtmlTags(input);
            
            // Create a JObject to simulate how it's stored in ts_details
            var detailObject = new JObject
            {
                ["question"] = "Comment",
                ["answer"] = cleanedText
            };
            var jsonArray = new JArray(detailObject);
            var jsonString = jsonArray.ToString(Newtonsoft.Json.Formatting.None);
            
            // Smart quotes should NOT be escaped in the JSON string
            Assert.DoesNotContain("\\\"", jsonString);
            Assert.Contains("\u201Csmart quotes\u201D", jsonString);
            
            // And when parsed back, they remain as smart quotes
            var parsedArray = JArray.Parse(jsonString);
            var parsedAnswer = parsedArray[0]["answer"].ToString();
            Assert.Equal("\u201Csmart quotes\u201D 'single quotes'", parsedAnswer);
        }

        #endregion

        [Fact]
        public void FormatMatrixResponse_SingleResultRow_ReturnsOnlyValue()
        {
            var definition = JObject.Parse(@"{
                ""type"": ""matrix"",
                ""name"": ""q1"",
                ""rows"": [{ ""value"": ""Result"", ""text"": ""Result"" }],
                ""columns"": [{ ""value"": ""sat"", ""text"": ""Satisfactory"" }]
            }");
            var response = JObject.Parse(@"{ ""Result"": ""sat"" }");

            var result = _formatter.FormatResponse(response, "matrix", definition);

            Assert.Equal("\"Satisfactory\"", result);
        }

        [Fact]
        public void FormatMatrixResponse_MultipleRows_ReturnsKeyValuePairs()
        {
            var definition = JObject.Parse(@"{
                ""type"": ""matrix"",
                ""name"": ""q1"",
                ""rows"": [
                    { ""value"": ""row1"", ""text"": ""Row 1"" },
                    { ""value"": ""row2"", ""text"": ""Row 2"" }
                ],
                ""columns"": [
                    { ""value"": ""c1"", ""text"": ""Col 1"" }
                ]
            }");
            var response = JObject.Parse(@"{ ""row1"": ""c1"", ""row2"": ""c1"" }");

            var result = _formatter.FormatResponse(response, "matrix", definition);

            Assert.Contains("\"Row 1\": \"Col 1\"", result);
            Assert.Contains("\"Row 2\": \"Col 1\"", result);
        }

        #region FormatResponse Text Type Tests

        [Fact]
        public void FormatResponse_TextWithNewlines_ReplacesWithSpaces()
        {
            var responseValue = JToken.FromObject("Line one\nLine two\nLine three");
            var result = _formatter.FormatResponse(responseValue, "text", null);
            Assert.Equal("Line one Line two Line three", result);
        }

        [Fact]
        public void FormatResponse_TextWithTabs_ReplacesWithSpaces()
        {
            var responseValue = JToken.FromObject("col1\tcol2\tcol3");
            var result = _formatter.FormatResponse(responseValue, "text", null);
            Assert.Equal("col1 col2 col3", result);
        }

        [Fact]
        public void FormatResponse_CommentWithNewlines_ReplacesWithSpaces()
        {
            var responseValue = JToken.FromObject("First paragraph.\n\nSecond paragraph.");
            var result = _formatter.FormatResponse(responseValue, "comment", null);
            Assert.Equal("First paragraph. Second paragraph.", result);
        }

        #endregion

        #region Truncation Edge Cases

        [Fact]
        public void RemoveHtmlTags_VeryLongString_HandlesWithoutError()
        {
            var longInput = new string('a', 10000) + "<br>" + new string('b', 10000);
            var result = _formatter.RemoveHtmlTags(longInput);
            
            Assert.NotNull(result);
            Assert.Contains(new string('a', 100), result);
            Assert.Contains(new string('b', 100), result);
            Assert.DoesNotContain("<br>", result);
        }

        [Fact]
        public void RemoveHtmlTags_UnicodeCharacters_PreservesCorrectly()
        {
            var input = "日本語テスト 中文测试 한국어테스트";
            var result = _formatter.RemoveHtmlTags(input);
            Assert.Equal(input, result);
        }

        [Fact]
        public void RemoveHtmlTags_MixedEmojiAndHtml_HandlesCorrectly()
        {
            var input = "<p>Test 🎉 emoji</p> and <br>more 🚀 here";
            var result = _formatter.RemoveHtmlTags(input);
            Assert.Equal("Test 🎉 emoji and more 🚀 here", result);
        }

        #endregion

        #region Zero-Width Character Tests

        [Fact]
        public void RemoveHtmlTags_ZeroWidthSpaces_RemovesCorrectly()
        {
            var input = "text\u200Bwith\u200Bzero\u200Bwidth";
            var result = _formatter.RemoveHtmlTags(input);
            Assert.Equal("textwithzerowidth", result);
        }

        [Fact]
        public void RemoveHtmlTags_ByteOrderMark_RemovesCorrectly()
        {
            var input = "\uFEFFText with BOM";
            var result = _formatter.RemoveHtmlTags(input);
            Assert.Equal("Text with BOM", result);
        }

        #endregion

        #region Dropdown Type Tests

        [Fact]
        public void FormatResponse_Dropdown_ReturnsDisplayText()
        {
            var definition = JObject.Parse(@"{
                ""type"": ""dropdown"",
                ""name"": ""q1"",
                ""choices"": [{ ""value"": ""opt1"", ""text"": { ""default"": ""Option One"" } }]
            }");
            var responseValue = JToken.FromObject("opt1");

            var result = _formatter.FormatResponse(responseValue, "dropdown", definition);

            Assert.Equal("Option One", result);
        }

        #endregion
    }
}
