using System;
using TSIS2.Plugins.QuestionnaireProcessor;
using Xunit.Abstractions;

namespace ROMTS_GSRST.Plugins.Tests.QuestionnaireProcessorTests
{
    /// <summary>
    /// Adapter that implements ILoggingService and writes to xUnit's ITestOutputHelper.
    /// </summary>
    public class TestLoggingAdapter : ILoggingService
    {
        private readonly ITestOutputHelper _output;

        public TestLoggingAdapter(ITestOutputHelper output)
        {
            _output = output;
        }

        public void Trace(string message) => WriteLine($"[TRACE] {message}");

        public void Trace(string format, params object[] args) => WriteLine($"[TRACE] {string.Format(format, args)}");

        public void Error(string message) => WriteLine($"[ERROR] {message}");

        public void Warning(string message) => WriteLine($"[WARNING] {message}");

        public void Info(string message) => WriteLine($"[INFO] {message}");

        public void Processing(string message) => WriteLine($"[PROCESSING] {message}");

        public void Verbose(string message) => WriteLine($"[VERBOSE] {message}");

        public void Debug(string message) => WriteLine($"[DEBUG] {message}");

        public void TraceErrorWithDebugInfo(string basicMessage, string detailedMessage)
        {
            WriteLine($"[ERROR] {basicMessage}");
            WriteLine($"[DEBUG] {detailedMessage}");
        }

        public bool VerboseMode => true;

        private void WriteLine(string message)
        {
            try
            {
                _output.WriteLine(message);
            }
            catch (Exception)
            {
                // Ignore errors if output is closed
            }
        }
    }
}
