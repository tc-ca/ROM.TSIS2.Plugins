using Microsoft.Xrm.Sdk;
using System;

namespace TSIS2.Plugins.QuestionnaireExtractor
{
    /// <summary>
    /// Adapter that wraps ITracingService to implement ILoggingService for Dynamics 365 plugins.
    /// This allows the questionnaire processing logic to work with the standard Dynamics tracing service.
    /// </summary>
    public class TracingServiceAdapter : ILoggingService
    {
        private readonly ITracingService _tracingService;
        private readonly LogLevel _minLogLevel;

        public TracingServiceAdapter(ITracingService tracingService, LogLevel minLogLevel = LogLevel.Info)
        {
            _tracingService = tracingService ?? throw new ArgumentNullException(nameof(tracingService));
            _minLogLevel = minLogLevel;
        }

        public void Trace(string message) => LogIfEnabled(LogLevel.Info, message);

        public void Trace(string format, params object[] args) => LogIfEnabled(LogLevel.Info, string.Format(format, args));

        public void Error(string message) => LogIfEnabled(LogLevel.Error, $"ERROR: {message}");

        public void Warning(string message) => LogIfEnabled(LogLevel.Warning, $"WARNING: {message}");

        public void Info(string message) => LogIfEnabled(LogLevel.Info, $"INFO: {message}");

        public void Processing(string message) => LogIfEnabled(LogLevel.Processing, $"PROCESSING: {message}");

        public void Verbose(string message) => LogIfEnabled(LogLevel.Verbose, $"VERBOSE: {message}");

        public void Debug(string message) => LogIfEnabled(LogLevel.Debug, $"DEBUG: {message}");

        /// <summary>
        /// Logs the message only if the specified level is at or below the minimum log level.
        /// </summary>
        /// <param name="level">The log level of the message.</param>
        /// <param name="message">The message to log.</param>
        private void LogIfEnabled(LogLevel level, string message)
        {
            if (level <= _minLogLevel)
                _tracingService.Trace(message);
        }
    }
} 