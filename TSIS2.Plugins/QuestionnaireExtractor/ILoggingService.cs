using System;

namespace TSIS2.Plugins.QuestionnaireExtractor
{
    /// <summary>
    /// Log levels in ascending order of verbosity.
    /// </summary>
    public enum LogLevel
    {
        Error = 0,      // Only errors and critical issues
        Warning = 1,    // Errors and warnings
        Info = 2,       // General progress information
        Processing = 3, // Highlighted processing messages (always shown)
        Verbose = 4,    // Detailed operation information
        Debug = 5       // Very detailed debugging information
    }

    /// <summary>
    /// Abstraction layer for logging that can work with both standalone applications and Dynamics 365 plugins.
    /// This interface provides a consistent logging API regardless of the underlying logging system.
    /// </summary>
    public interface ILoggingService
    {
        void Trace(string message);

        void Trace(string format, params object[] args);

        void Error(string message);

        void Warning(string message);

        void Info(string message);

        void Processing(string message);

        void Verbose(string message);

        void Debug(string message);
    }
} 