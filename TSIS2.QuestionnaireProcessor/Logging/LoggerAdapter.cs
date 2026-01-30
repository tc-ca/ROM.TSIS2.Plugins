using System;

namespace TSIS2.Plugins.QuestionnaireProcessor
{
    /// <summary>
    /// Adapter that implements ILoggingService for standalone applications with console output.
    /// This provides a simple console-based logging implementation.
    /// </summary>
    public class LoggerAdapter : ILoggingService
    {
        private LogLevel _currentLogLevel;
        private bool _simulationMode;

        public LoggerAdapter()
        {
            _currentLogLevel = LogLevel.Info;
            _simulationMode = false;
        }

        public void SetLogLevel(LogLevel level)
        {
            _currentLogLevel = level;
        }

        public bool VerboseMode => _currentLogLevel >= LogLevel.Verbose;

        public void SetSimulationMode(bool simulationMode)
        {
            _simulationMode = simulationMode;
        }

        public void Trace(string message) => LogIfEnabled(LogLevel.Verbose, message);

        public void Trace(string format, params object[] args) => LogIfEnabled(LogLevel.Verbose, string.Format(format, args));

        public void Error(string message) => LogIfEnabled(LogLevel.Error, $"ERROR: {message}");

        public void Warning(string message) => LogIfEnabled(LogLevel.Warning, $"WARNING: {message}");

        public void Info(string message) => LogIfEnabled(LogLevel.Info, message);

        public void Processing(string message) => LogIfEnabled(LogLevel.Processing, $">>> {message}");

        public void Verbose(string message) => LogIfEnabled(LogLevel.Verbose, message);

        public void Debug(string message) => LogIfEnabled(LogLevel.Debug, message);

        public void TraceErrorWithDebugInfo(string basicMessage, string detailedMessage)
        {
            Error(basicMessage);
            Debug(detailedMessage);
        }

        private void LogIfEnabled(LogLevel level, string message)
        {
            if (level <= _currentLogLevel || level == LogLevel.Processing)
            {
                switch (level)
                {
                    case LogLevel.Error:
                        Console.ForegroundColor = _simulationMode ? ConsoleColor.DarkRed : ConsoleColor.Red;
                        break;
                    case LogLevel.Warning:
                        Console.ForegroundColor = _simulationMode ? ConsoleColor.DarkYellow : ConsoleColor.Yellow;
                        break;
                    case LogLevel.Info:
                        Console.ForegroundColor = _simulationMode ? ConsoleColor.Yellow : ConsoleColor.White;
                        break;
                    case LogLevel.Processing:
                        Console.ForegroundColor = _simulationMode ? ConsoleColor.DarkMagenta : ConsoleColor.Magenta;
                        break;
                    case LogLevel.Verbose:
                        Console.ForegroundColor = _simulationMode ? ConsoleColor.DarkGray : ConsoleColor.Gray;
                        break;
                    case LogLevel.Debug:
                        Console.ForegroundColor = _simulationMode ? ConsoleColor.DarkGray : ConsoleColor.DarkGray;
                        break;
                }

                Console.WriteLine(message);
                Console.ResetColor();
            }
        }
    }
}