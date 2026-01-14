using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using System.IO;

namespace TSIS2.QuestionnaireProcessorConsole.Logging
{
    public class LogFileTracingService : ITracingService, IDisposable
    {
        private readonly StreamWriter _logFileWriter;
        private readonly string _logFilePath;

        public LogFileTracingService()
        {
            // Create a timestamped filename
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string logFileName = $"QuestionnaireOrchestrator_{timestamp}.log";

            // Create logs directory if it doesn't exist
            string logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
            if (!Directory.Exists(logDirectory))
            {
                Directory.CreateDirectory(logDirectory);
            }

            // Setup the full file path
            _logFilePath = Path.Combine(logDirectory, logFileName);

            // Create the file writer
            _logFileWriter = new StreamWriter(_logFilePath, true);

            // Write header to log file
            _logFileWriter.WriteLine($"=== Questionnaire Processor Log - Started at {DateTime.Now} ===");
            _logFileWriter.WriteLine();
            _logFileWriter.Flush();

            Console.WriteLine($"Logging to file: {_logFilePath}");
        }

        public void Trace(string message)
        {
            // Add timestamp to log file entries
            string timestampedMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}";

            // Write to log file
            _logFileWriter.WriteLine(timestampedMessage);
            _logFileWriter.Flush(); // Ensure it's written immediately
        }
        public void TraceErrorWithDebugInfo(string message, string detailedError)
        {
            string timestampedMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}";
            _logFileWriter.WriteLine(timestampedMessage);
            _logFileWriter.WriteLine(detailedError);
            _logFileWriter.Flush();
        }

        public void Trace(string format, params object[] args)
        {
            if (args != null && args.Length > 0)
            {
                Trace(string.Format(format, args));
            }
            else
            {
                Trace(format);
            }
        }

        public void Dispose()
        {
            // Write footer to log file
            _logFileWriter.WriteLine();
            _logFileWriter.WriteLine($"=== Logging Ended at {DateTime.Now} ===");

            // Close and dispose the writer
            _logFileWriter.Close();
            _logFileWriter.Dispose();
        }
    }
}


