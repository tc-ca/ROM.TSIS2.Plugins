using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace TSIS2.QuestionnaireProcessorConsole
{
    public enum UserAction
    {
        Proceed,
        Simulate,
        Cancel
    }

    public class ConsoleUI
    {
        public void ShowWelcomeMessage()
        {
            Console.WriteLine("Questionnaire Processor Console Application");
            Console.WriteLine("==========================================");
            Console.WriteLine($"Started at {DateTime.Now:yyyy-MM-dd hh:mm:ss tt}");
            Console.WriteLine();
        }

        public void ShowLogLevel(string logLevel)
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine($"Current log level: {logLevel}");
            Console.ResetColor();
        }

        public void ShowConnectionStatus(string url)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("\n===================================================");
            Console.WriteLine($"Connecting to: {url}");
            Console.WriteLine("===================================================\n");
            Console.ResetColor();
        }

        public void ShowDataRetrievalProgress(int pageSize)
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("Data retrieval in progress:");
            Console.WriteLine("- Using FetchXML pagination to handle large datasets");
            Console.WriteLine($"- Page size: {pageSize} records per request");
            Console.WriteLine();
            Console.ResetColor();
        }

        public async Task ShowLoadingIndicator(CancellationToken cancellationToken)
        {
            try
            {
                string[] frames = { "⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏" };
                int frameIndex = 0;
                
                while (!cancellationToken.IsCancellationRequested)
                {
                    Console.Write($"\rFetching records {frames[frameIndex]} ");
                    frameIndex = (frameIndex + 1) % frames.Length;
                    await Task.Delay(100, cancellationToken);
                }
                
                // Clear the loading indicator line
                Console.Write("\r                      \r");
            }
            catch (OperationCanceledException)
            {
                // This is expected when cancellation is requested
                Console.Write("\r                      \r");
            }
        }

        public void ShowDataRetrievalSummary(int recordCount, bool useTargetedWost, int pageSize, int totalPages)
        {
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Green;
            if (useTargetedWost)
            {
                Console.WriteLine($"[TARGET] Found {recordCount} targeted work order service tasks with unprocessed questionnaires");
            }
            else
            {
                Console.WriteLine($"===> Found {recordCount} work order service tasks with unprocessed questionnaires");
            }
            
            // Add details about the pagination that was used
            Console.WriteLine($"     Data retrieved in {totalPages} page(s) of up to {pageSize} records each");
            Console.WriteLine($"     Each WOST will be processed individually to ensure reliability");
            Console.ResetColor();
            Console.WriteLine();
        }

        public void ShowPagingProgress(int pageCount, int recordsInPage, int totalRecords)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"[PAGE] Retrieved Page {pageCount}: {recordsInPage} records (Total: {totalRecords})");
            Console.ResetColor();
        }

        public void ShowCompletionMessage(int totalRecords, int pageCount)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"[OK] Paging complete: {totalRecords} total records retrieved across {pageCount} pages\n");
            Console.ResetColor();
        }

        public UserAction GetUserConfirmation()
        {
            Console.Write("Do you want to proceed with processing these questionnaires? (Y/N/S=Simulate): ");
            var response = Console.ReadLine()?.Trim().ToUpper();
            
            if (response == "S")
            {
                return UserAction.Simulate;
            }
            else if (response == "Y" || response == "YES")
            {
                return UserAction.Proceed;
            }
            else
            {
                return UserAction.Cancel;
            }
        }

        public void DisplayFatalError(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"FATAL ERROR: {message}");
            Console.ResetColor();
        }

        public void DisplayFinalSummary(ProcessingSummary summary)
        {
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine($"=== Processing Summary ===");
            if (summary.IsSimulationMode)
            {
                Console.WriteLine("--- NO DATABASE CHANGES WERE MADE ---");
            }
            Console.WriteLine($"Total tasks found: {summary.TotalTasks}");
            Console.WriteLine($"Successfully processed: {summary.Processed}");
            Console.WriteLine($"Failed: {summary.Failed}");
            Console.WriteLine($"Total question responses created: {summary.TotalResponsesCreated}");
            Console.WriteLine($"Total time: {summary.TotalTime.ToString(@"hh\:mm\:ss")}");
            if (summary.Processed > 0)
            {
                Console.WriteLine($"Average processing time: {summary.TotalTime.TotalSeconds / summary.Processed:F2} seconds per task");
            }
            Console.ResetColor();
        }

        public void WaitForExit()
        {
            // Program will exit automatically without waiting for user input
            // Console.WriteLine("\nPress any key to exit...");
            // Console.ReadKey();
        }

        public void ShowError(string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"ERROR: {message}");
            Console.ResetColor();
        }

        public void ShowWarning(string message)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"WARNING: {message}");
            Console.ResetColor();
        }

        public void ShowSuccess(string message)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(message);
            Console.ResetColor();
        }

        public void ShowInfo(string message)
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine($"INFO: {message}");
            Console.ResetColor();
        }

        public void ShowHelp()
        {
            Console.WriteLine("TSIS2.QuestionnaireProcessorConsole");
            Console.WriteLine("====================================");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("  TSIS2.QuestionnaireProcessorConsole.exe [options]");
            Console.WriteLine();
            Console.WriteLine("Environment Options:");
            Console.WriteLine("  --env <name>    Select environment from secrets.config (e.g., --env dev, --env prod)");
            Console.WriteLine("                  If omitted and multiple environments exist, you'll be prompted to choose.");
            Console.WriteLine();
            Console.WriteLine("Logging Options:");
            Console.WriteLine("  --debug         Set log level to Debug (most verbose)");
            Console.WriteLine("  --verbose       Set log level to Verbose");
            Console.WriteLine("  --quiet         Set log level to Warning (least verbose)");
            Console.WriteLine("                  Default log level is read from settings.config");
            Console.WriteLine("  --logtofile     Enable logging to file (creates timestamped log in logs/ directory)");
            Console.WriteLine("                  By default, logs are only written to the console.");
            Console.WriteLine();
            Console.WriteLine("Operation Options:");
            Console.WriteLine("  --simulate      Run in simulation mode (no database changes will be made)");
            Console.WriteLine("  --backfill-workorder");
            Console.WriteLine("                  Backfill ts_workorder field on existing ts_questionresponse records");
            Console.WriteLine("  --guids <guids> Process only the specified WOST GUIDs (comma-separated)");
            Console.WriteLine("                  Supports GUIDs with or without braces: guid1,{guid2},guid3");
            Console.WriteLine("  --wost-ids <guids>");
            Console.WriteLine("                  Alias for --guids");
            Console.WriteLine("  --from-file     Process WOSTs from wost_ids.txt file before normal processing");
            Console.WriteLine("  --page-size <n>");
            Console.WriteLine("                  Override page size for FetchXML queries (default: from settings.config)");
            Console.WriteLine();
            Console.WriteLine("Date Filter Options:");
            Console.WriteLine("  --since <date>, --created-before <date>");
            Console.WriteLine("                  Only process WOSTs created before the specified date.");
            Console.WriteLine("                  Cannot be used with --guids (applies only to FetchXML queries).");
            Console.WriteLine();
            Console.WriteLine("                  Supported date formats:");
            Console.WriteLine("                    2023-05-25              (date only, midnight)");
            Console.WriteLine("                    2023-05-25T19:36:41     (date and time, local)");
            Console.WriteLine("                    2023-05-25T19:36:41Z    (date and time, UTC)");
            Console.WriteLine("                    05/25/2023              (US format)");
            Console.WriteLine("                    25/05/2023              (European format, depends on system locale)");
            Console.WriteLine();
            Console.WriteLine("Help:");
            Console.WriteLine("  --help, -help, -h, /?");
            Console.WriteLine("                  Display this help message");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("  TSIS2.QuestionnaireProcessorConsole.exe --env dev");
            Console.WriteLine("  TSIS2.QuestionnaireProcessorConsole.exe --env prod --debug");
            Console.WriteLine("  TSIS2.QuestionnaireProcessorConsole.exe --env dev --backfill-workorder");
            Console.WriteLine("  TSIS2.QuestionnaireProcessorConsole.exe --simulate --verbose");
            Console.WriteLine("  TSIS2.QuestionnaireProcessorConsole.exe --env dev --guids f478ef1a-04f5-ed11-8848-000d3af4fb40");
            Console.WriteLine("  TSIS2.QuestionnaireProcessorConsole.exe --env prod --guids=f478ef1a-04f5-ed11-8848-000d3af4fb40,{b56cfc0f-b761-f011-bec1-002248b09c68} --simulate");
            Console.WriteLine("  TSIS2.QuestionnaireProcessorConsole.exe --env dev --from-file");
            Console.WriteLine("  TSIS2.QuestionnaireProcessorConsole.exe --env prod --page-size 1000");
            Console.WriteLine();
            Console.WriteLine("Date Filter Examples:");
            Console.WriteLine("  TSIS2.QuestionnaireProcessorConsole.exe --env dev --since 2023-05-25");
            Console.WriteLine("  TSIS2.QuestionnaireProcessorConsole.exe --env prod --since=2023-05-25T19:36:41Z --simulate");
            Console.WriteLine("  TSIS2.QuestionnaireProcessorConsole.exe --env dev --created-before 2023-05-25");
            Console.WriteLine();
        }
    }
}