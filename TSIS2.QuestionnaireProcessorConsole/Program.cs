using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Xrm.Tooling.Connector;
using TSIS2.Plugins.QuestionnaireProcessor;
using TSIS2.QuestionnaireProcessorConsole.Logging;

namespace TSIS2.QuestionnaireProcessorConsole
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var ui = new ConsoleUI();

            // Check for help flag first
            if (args != null && args.Length > 0)
            {
                foreach (var arg in args)
                {
                    if (arg.Equals("--help", StringComparison.OrdinalIgnoreCase) ||
                        arg.Equals("-help", StringComparison.OrdinalIgnoreCase) ||
                        arg.Equals("-h", StringComparison.OrdinalIgnoreCase) ||
                        arg.Equals("/?", StringComparison.OrdinalIgnoreCase))
                    {
                        ui.ShowHelp();
                        return;
                    }
                }
            }

            ui.ShowWelcomeMessage();

            try
            {
                var config = new ConfigurationService(args);

                // Validate conflicting arguments
                if (config.SpecificWostIds.Count > 0 && config.CreatedBefore.HasValue)
                {
                    ui.DisplayFatalError("Cannot use --since/--created-before with --guids. The date filter only applies to FetchXML queries, not specific GUIDs.");
                    return;
                }

                // Create logger - only log to file if --logtofile is specified
                TSIS2.Plugins.QuestionnaireProcessor.ILoggingService logger;
                LogFileTracingService fileTracingService = null;

                if (config.LogToFile)
                {
                    // Create a composite logger that writes to both console and file
                    fileTracingService = new LogFileTracingService();
                    var fileLogger = new TracingServiceAdapter(fileTracingService);
                    var consoleLogger = new LoggerAdapter();
                    logger = new CompositeLogger(consoleLogger, fileLogger);
                }
                else
                {
                    // Console-only logging (default)
                    logger = new LoggerAdapter();
                }

                if (System.Enum.TryParse(config.LogLevel, true, out LogLevel level))
                {
                    if (logger is CompositeLogger compositeLogger)
                    {
                        compositeLogger.SetLogLevel(level);
                    }
                    else if (logger is LoggerAdapter loggerAdapter)
                    {
                        loggerAdapter.SetLogLevel(level);
                    }
                }

                ui.ShowLogLevel(config.LogLevel);
                
                // Show date filter if active
                if (config.CreatedBefore.HasValue)
                {
                    ui.ShowInfo($"Date filter active: Only processing WOSTs created before {config.CreatedBefore.Value:yyyy-MM-dd HH:mm:ss}");
                }
                
                ui.ShowConnectionStatus(config.Url);

                using (var crmClient = new CrmServiceClient(config.ConnectString))
                {
                    if (!crmClient.IsReady)
                    {
                        ui.DisplayFatalError($"Connection Failed: {crmClient.LastCrmError}");
                        return;
                    }
                    logger.Info("Connected successfully!");

                    // If we're running the one-off backfill, do that and exit.
                    if (config.BackfillWorkOrderRefs)
                    {
                        var backfiller = new QuestionResponseBackfiller(crmClient, logger);
                        var result = backfiller.BackfillWorkOrderReference(config.SimulationMode);
                        ui.ShowSuccess($"Backfill complete. Scanned {result.TotalScanned}, updated {result.Updated}, skipped (no work order) {result.SkippedNoWorkOrder}.");
                        return;
                    }

                    var dataService = new CrmDataService(crmClient, logger, ui, config.PageSize);
                    var processor = new WostProcessor(crmClient, logger, ui);

                    // If specific GUIDs are provided, process only those WOSTs
                    if (config.SpecificWostIds.Count > 0)
                    {
                        logger.Info($"[GUIDS] Processing {config.SpecificWostIds.Count} specific WOST(s)");
                        var wostList = new List<Microsoft.Xrm.Sdk.Entity>();
                        var failedGuids = new List<Guid>();

                        foreach (var wostId in config.SpecificWostIds)
                        {
                            try
                            {
                                var wost = crmClient.Retrieve("msdyn_workorderservicetask", wostId,
                                    new Microsoft.Xrm.Sdk.Query.ColumnSet("msdyn_name", "ovs_questionnaire", "ovs_questionnaireresponse", "ovs_questionnairedefinition"));

                                string wostName = wost.GetAttributeValue<string>("msdyn_name");
                                var questionnaireRef = wost.GetAttributeValue<Microsoft.Xrm.Sdk.EntityReference>("ovs_questionnaire");

                                if (questionnaireRef == null)
                                {
                                    ui.ShowWarning($"WOST {wostName} (ID: {wostId}) has no linked questionnaire reference - skipping.");
                                    failedGuids.Add(wostId);
                                    continue;
                                }

                                wostList.Add(wost);
                                ui.ShowInfo($"Found WOST: {wostName} (ID: {wostId})");
                            }
                            catch (Exception ex)
                            {
                                ui.ShowError($"Error retrieving WOST {wostId}: {ex}");
                                failedGuids.Add(wostId);
                            }
                        }

                        if (wostList.Count == 0)
                        {
                            ui.DisplayFatalError($"No valid WOSTs found to process. {failedGuids.Count} GUID(s) failed.");
                            return;
                        }

                        if (failedGuids.Count > 0)
                        {
                            ui.ShowWarning($"{failedGuids.Count} GUID(s) failed to retrieve, but proceeding with {wostList.Count} valid WOST(s).");
                        }

                        UserAction confirmation = UserAction.Proceed;
                        if (!config.SimulationMode)
                        {
                            confirmation = ui.GetUserConfirmation();
                            if (confirmation == UserAction.Cancel) return;
                        }

                        bool runInSimMode = config.SimulationMode || (confirmation == UserAction.Simulate);
                        processor.IsSimulationMode = runInSimMode;

                        var summary = processor.Process(wostList);
                        ui.DisplayFinalSummary(summary);
                        return;
                    }

                    if (config.ProcessWostsFromFile)
                    {
                        logger.Info($"[FILE] Processing WOSTs from file '{config.WostIdsFileName}' first...");
                        var wostIds = processor.LoadWostIdsFromFile(config.WostIdsFileName);

                        if (wostIds.Count > 0)
                        {
                            UserAction confirmation = UserAction.Proceed;
                            if (!config.SimulationMode)
                            {
                                confirmation = ui.GetUserConfirmation();
                                if (confirmation == UserAction.Cancel) return;
                            }

                            bool runInSimMode = config.SimulationMode || (confirmation == UserAction.Simulate);
                            processor.IsSimulationMode = runInSimMode;

                            var fileSummary = processor.ProcessFromFile(wostIds);
                            ui.ShowSuccess($"[OK] Finished processing {fileSummary.Processed} WOST(s) from file");
                        }
                        else
                        {
                            ui.ShowWarning("[FILE] No WOSTs found in file or file was empty");
                        }
                    }

                    string fetchXml = config.GetActiveFetchXml();
                    if (string.IsNullOrEmpty(fetchXml))
                    {
                        ui.DisplayFatalError("No valid FetchXML query found in configuration");
                        return;
                    }

                    ui.ShowDataRetrievalProgress(config.PageSize);

                    var loadingCancellation = new CancellationTokenSource();
                    var loadingTask = Task.Run(async () => await ui.ShowLoadingIndicator(loadingCancellation.Token));

                    var recordsToProcess = dataService.RetrieveAllPages(fetchXml, loadingCancellation.Token);

                    loadingCancellation.Cancel();
                    if (loadingTask != null) Task.WaitAll(new[] { loadingTask }, 1000);
                    Console.WriteLine();

                    // Show detailed data retrieval summary
                    int totalPages = (recordsToProcess.Count + config.PageSize - 1) / config.PageSize;
                    ui.ShowDataRetrievalSummary(recordsToProcess.Count, false, config.PageSize, totalPages);

                    ui.ShowInfo($"Found {recordsToProcess.Count} work order service tasks with unprocessed questionnaires");

                    if (recordsToProcess.Any())
                    {
                        UserAction confirmation = UserAction.Proceed;
                        if (!config.SimulationMode)
                        {
                            confirmation = ui.GetUserConfirmation();
                            if (confirmation == UserAction.Cancel) return;
                        }

                        bool runInSimMode = config.SimulationMode || (confirmation == UserAction.Simulate);
                        processor.IsSimulationMode = runInSimMode;

                        var mainSummary = processor.Process(recordsToProcess);
                        ui.DisplayFinalSummary(mainSummary);
                    }
                    else
                    {
                        ui.ShowInfo("No records found to process");
                    }
                }
            }
            catch (Exception ex)
            {
                ui.DisplayFatalError($"A fatal error occurred: {ex}");
            }

            ui.WaitForExit();
        }
    }

    /// <summary>
    /// Composite logger that writes to multiple logging services.
    /// </summary>
    public class CompositeLogger : TSIS2.Plugins.QuestionnaireProcessor.ILoggingService
    {
        private readonly TSIS2.Plugins.QuestionnaireProcessor.ILoggingService _consoleLogger;
        private readonly TSIS2.Plugins.QuestionnaireProcessor.ILoggingService _fileLogger;
        private LogLevel _logLevel = LogLevel.Info;

        public CompositeLogger(TSIS2.Plugins.QuestionnaireProcessor.ILoggingService consoleLogger, TSIS2.Plugins.QuestionnaireProcessor.ILoggingService fileLogger)
        {
            _consoleLogger = consoleLogger ?? throw new ArgumentNullException(nameof(consoleLogger));
            _fileLogger = fileLogger ?? throw new ArgumentNullException(nameof(fileLogger));
        }

        public void SetLogLevel(LogLevel level)
        {
            _logLevel = level;
            if (_consoleLogger is LoggerAdapter consoleAdapter)
            {
                consoleAdapter.SetLogLevel(level);
            }
        }

        public bool VerboseMode => _logLevel >= LogLevel.Verbose;

        public void Trace(string message)
        {
            _consoleLogger.Trace(message);
            _fileLogger.Trace(message);
        }

        public void Trace(string format, params object[] args)
        {
            _consoleLogger.Trace(format, args);
            _fileLogger.Trace(format, args);
        }

        public void Error(string message)
        {
            _consoleLogger.Error(message);
            _fileLogger.Error(message);
        }

        public void Warning(string message)
        {
            _consoleLogger.Warning(message);
            _fileLogger.Warning(message);
        }

        public void Info(string message)
        {
            _consoleLogger.Info(message);
            _fileLogger.Info(message);
        }

        public void Processing(string message)
        {
            _consoleLogger.Processing(message);
            _fileLogger.Processing(message);
        }

        public void Verbose(string message)
        {
            _consoleLogger.Verbose(message);
            _fileLogger.Verbose(message);
        }

        public void Debug(string message)
        {
            _consoleLogger.Debug(message);
            _fileLogger.Debug(message);
        }
        public void TraceErrorWithDebugInfo(string basicMessage, string detailedMessage)
        {
            _consoleLogger.TraceErrorWithDebugInfo(basicMessage, detailedMessage);
            _fileLogger.TraceErrorWithDebugInfo(basicMessage, detailedMessage);
        }
    }
}
