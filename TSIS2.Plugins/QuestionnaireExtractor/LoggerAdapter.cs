using System;
using System.Configuration;

namespace TSIS2.Plugins.QuestionnaireExtractor
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
            // Read log level from configuration
            _currentLogLevel = GetLogLevelFromConfig();
            _simulationMode = false; // Default simulation mode
        }

        /// <summary>
        /// Gets the log level from configuration.
        /// </summary>
        /// <returns>The configured log level, or Info as default.</returns>
        private LogLevel GetLogLevelFromConfig()
        {
            try
            {
                string configLogLevel = ConfigurationManager.AppSettings["LogLevel"];
                if (!string.IsNullOrEmpty(configLogLevel))
                {
                    if (Enum.TryParse(configLogLevel, true, out LogLevel level))
                    {
                        return level;
                    }
                }
            }
            catch
            {
                // If we can't read from config, use default
            }
            
            return LogLevel.Info; // Default fallback
        }

        public void SetLogLevel(LogLevel level)
        {
            _currentLogLevel = level;
        }

        public void SetSimulationMode(bool simulationMode)
        {
            _simulationMode = simulationMode;
        }

        public void Trace(string message) => LogIfEnabled(LogLevel.Info, message);

        public void Trace(string format, params object[] args) => LogIfEnabled(LogLevel.Info, string.Format(format, args));

        public void Error(string message) => LogIfEnabled(LogLevel.Error, $"ERROR: {message}");

        public void Warning(string message) => LogIfEnabled(LogLevel.Warning, $"WARNING: {message}");

        public void Info(string message) => LogIfEnabled(LogLevel.Info, message);

        public void Processing(string message) => LogIfEnabled(LogLevel.Processing, $">>> {message}");

        public void Verbose(string message) => LogIfEnabled(LogLevel.Verbose, message);

        public void Debug(string message) => LogIfEnabled(LogLevel.Debug, message);

        /// <summary>
        /// Logs the message only if the specified level is at or below the current log level.
        /// </summary>
        /// <param name="level">The log level of the message.</param>
        /// <param name="message">The message to log.</param>
        private void LogIfEnabled(LogLevel level, string message)
        {
            if (level <= _currentLogLevel || level == LogLevel.Processing)
            {
                // Add level-specific formatting
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