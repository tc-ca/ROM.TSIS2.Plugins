using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using TSIS2.Plugins.QuestionnaireProcessor;
using TSIS2.QuestionnaireProcessorConsole;

public class ProcessingSummary
{
    public int Processed { get; set; }
    public int Failed { get; set; }
    public TimeSpan TotalTime { get; set; }
    public int TotalResponsesCreated { get; set; }
    public bool IsSimulationMode { get; set; }
    public int TotalTasks { get; set; }
    public List<(Guid Id, string Name, string ErrorMessage, string DetailedError)> FailedTasks { get; set; } = new List<(Guid Id, string Name, string ErrorMessage, string DetailedError)>();
}

public class WostProcessor
{
    private readonly IOrganizationService _service;
    private readonly TSIS2.Plugins.QuestionnaireProcessor.ILoggingService _logger;
    private readonly ConsoleUI _ui;
    public bool IsSimulationMode { get; set; } = false;

    public WostProcessor(IOrganizationService service, TSIS2.Plugins.QuestionnaireProcessor.ILoggingService logger, ConsoleUI ui)
    {
        _service = service;
        _logger = logger;
        _ui = ui;
    }

    public ProcessingSummary Process(List<Entity> workOrderServiceTasks)
    {
        var summary = new ProcessingSummary
        {
            TotalTasks = workOrderServiceTasks.Count,
            IsSimulationMode = IsSimulationMode
        };
        var failedTasks = new List<(Guid Id, string Name, string ErrorMessage, string DetailedError)>();
        int processed = 0;
        int failed = 0;
        int totalResponsesCreated = 0;
        DateTime startTime = DateTime.Now;

        _logger.Info($"Beginning to process {workOrderServiceTasks.Count} questionnaires...");

        foreach (var wost in workOrderServiceTasks)
        {
            Guid wostId = wost.Id;
            string wostName = wost.GetAttributeValue<string>("msdyn_name");
            _logger.Verbose($"Processing WOST: {wostName}");

            try
            {
                var wostDetail = _service.Retrieve("msdyn_workorderservicetask", wostId, new ColumnSet("ovs_questionnaire"));
                var questionnaireRef = wostDetail.GetAttributeValue<EntityReference>("ovs_questionnaire");

                if (questionnaireRef == null)
                {
                    _logger.Warning($"Skipping WOST {wostName} because it has no linked questionnaire reference.");
                    continue;
                }

                _logger.Processing($"Starting questionnaire processing for WOST: {wostName} (ID: {wostId})");

                var result = QuestionnaireOrchestrator.ProcessQuestionnaire(
                    _service,
                    wostId,
                    questionnaireRef,
                    false,
                    IsSimulationMode,
                    _logger,
                    _logger.VerboseMode
                );

                if (result.TotalCreatedOrUpdatedRecords > 0)
                {
                    totalResponsesCreated += result.CreatedResponseIds.Count;
                    _ui.ShowSuccess($">>> Completed questionnaire processing for WOST: {wostName}. Records Created/Updated/Up-to-date: {result.CreatedResponseIds.Count}/{result.UpdatedRecordsCount}/{result.UpToDateRecordsCount} (Total Records: {result.TotalCrmRecords})");
                }
                else
                {
                    _ui.ShowInfo($">>> Completed questionnaire processing for WOST: {wostName}. Records Up-to-date: {result.UpToDateRecordsCount}/{result.TotalCrmRecords}");
                }

                _ui.ShowInfo($"    Note: {result.VisibleQuestionCount} visible questions were processed, but {result.HiddenMergedCount} were merged into parents.");

                if (_logger.VerboseMode)
                {
                    DisplayQuestionLogTable(wostName, result);
                }

                processed++;
            }
            catch (Exception ex)
            {
                string errorMessage = ex.Message;
                string detailedErrorInfo = $"Failed processing response value: {(ex.Data.Contains("ResponseValue") ? ex.Data["ResponseValue"] : "N/A")}\n" +
                                           $"Error processing question: {(ex.Data.Contains("QuestionName") ? ex.Data["QuestionName"] : "N/A")}\n" +
                                           $"Exception Details: {ex}";

                _ui.ShowError($"Error processing WOST {wostName} - see log file for details");

                _logger.TraceErrorWithDebugInfo(
                $"Failed processing WOST {wostId}: {errorMessage}",
                detailedErrorInfo
                );

                failedTasks.Add((wostId, wostName, errorMessage, detailedErrorInfo));
                failed++;
            }
        }

        summary.Processed = processed;
        summary.Failed = failed;
        summary.TotalTime = DateTime.Now - startTime;
        summary.TotalResponsesCreated = totalResponsesCreated;
        summary.FailedTasks = failedTasks;

        return summary;
    }

    public ProcessingSummary ProcessFromFile(List<Guid> wostIds)
    {
        var summary = new ProcessingSummary();
        var processedWOSTs = new List<Guid>();
        int processed = 0;
        int failed = 0;
        DateTime startTime = DateTime.Now;

        _logger.Info($"[FILE] Processing {wostIds.Count} WOST(s) from file");

        foreach (var wostId in wostIds)
        {
            try
            {
                var wost = _service.Retrieve("msdyn_workorderservicetask", wostId,
                    new ColumnSet("msdyn_name", "ovs_questionnaire", "ovs_questionnaireresponse", "ovs_questionnairedefinition"));

                string wostName = wost.GetAttributeValue<string>("msdyn_name");
                var questionnaireRef = wost.GetAttributeValue<EntityReference>("ovs_questionnaire");

                if (questionnaireRef == null)
                {
                    _logger.Warning($"Skipping WOST {wostName} (ID: {wostId}) - no questionnaire reference");
                    continue;
                }

                _logger.Processing($"Starting questionnaire processing for WOST: {wostName} (ID: {wostId})");

                var result = QuestionnaireOrchestrator.ProcessQuestionnaire(
                    _service,
                    wostId,
                    questionnaireRef,
                    false,
                    IsSimulationMode,
                    _logger,
                    _logger.VerboseMode // Collect inventory only if verbose
                );

                if (result.UpdatedRecordsCount > 0 || result.CreatedResponseIds.Count > 0)
                {
                    _ui.ShowSuccess($">>> Completed WOST from file: {wostName}. Records Created/Updated/Up-to-date: {result.CreatedResponseIds.Count}/{result.UpdatedRecordsCount}/{result.UpToDateRecordsCount} (Total Records: {result.TotalCrmRecords})");
                }
                else
                {
                    _ui.ShowInfo($">>> Completed WOST from file: {wostName}. Records Up-to-date: {result.UpToDateRecordsCount}/{result.TotalCrmRecords}");
                }

                _ui.ShowInfo($"    Note: {result.VisibleQuestionCount} visible questions were processed, but {result.HiddenMergedCount} were merged into parents.");

                if (_logger.VerboseMode)
                {
                    DisplayQuestionLogTable(wostName, result);
                }

                processedWOSTs.Add(wostId);
                processed++;
            }
            catch (Exception ex)
            {
                _logger.Error($"Error processing WOST from file (ID: {wostId}): {ex}");
                _ui.ShowError($"[ERROR] Failed to process WOST from file (ID: {wostId}): {ex.Message}");
                failed++;
            }
        }

        summary.Processed = processed;
        summary.Failed = failed;
        summary.TotalTime = DateTime.Now - startTime;

        return summary;
    }

    public List<Guid> LoadWostIdsFromFile(string fileName = "wost_ids.txt")
    {
        var wostIds = new List<Guid>();
        string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);

        if (!File.Exists(filePath))
        {
            _logger.Warning($"File not found: {filePath}");
            _ui.ShowWarning($"[WARNING] File not found: {fileName}");
            _ui.ShowInfo($"   Expected location: {filePath}");
            return wostIds;
        }

        try
        {
            var lines = File.ReadAllLines(filePath);
            foreach (var line in lines)
            {
                var trimmedLine = line.Trim();
                if (!string.IsNullOrEmpty(trimmedLine) && Guid.TryParse(trimmedLine, out Guid wostId))
                {
                    wostIds.Add(wostId);
                }
                else if (!string.IsNullOrEmpty(trimmedLine))
                {
                    _logger.Warning($"Invalid GUID format in file: {trimmedLine}");
                }
            }

            if (wostIds.Count == 0)
            {
                _logger.Warning("No valid GUIDs found in file");
                return wostIds;
            }

            _logger.Info($"[FILE] Found {wostIds.Count} WOST ID(s) in file");
            _ui.ShowInfo($"[FILE] Found {wostIds.Count} WOST ID(s) in file: {fileName}");
            _ui.ShowInfo($"   File location: {filePath}");

            _ui.ShowInfo("\nSample WOST IDs from file (showing up to 5):");
            foreach (var wostId in wostIds.Take(5))
            {
                _ui.ShowInfo($"  - {wostId}");
            }

            if (wostIds.Count > 5)
            {
                _ui.ShowInfo($"  - ... and {wostIds.Count - 5} more");
            }

            return wostIds;
        }
        catch (Exception ex)
        {
            _logger.Error($"Error reading or processing WOST file: {ex}");
            _ui.ShowError($"[ERROR] Error reading or processing WOST file: {ex.Message}");
            return wostIds;
        }
    }
    private void DisplayQuestionLogTable(string wostName, QuestionnaireProcessResult result)
    {
        _ui.ShowInfo("\n------------------------------------------------------------------------------------------");
        _ui.ShowInfo($"QUESTION LOG REPORT: {wostName}");
        _ui.ShowInfo("------------------------------------------------------------------------------------------");
        _ui.ShowInfo(string.Format("{0,-4} | {1,-30} | {2,-30} | {3}", "#", "Question Name", "Dynamics Record Name", "Result"));
        _ui.ShowInfo("------------------------------------------------------------------------------------------");

        foreach (var item in result.Inventory.OrderBy(i => i.SequenceNumber ?? 999))
        {
            string seq = item.SequenceNumber?.ToString() ?? "--";
            _ui.ShowInfo(string.Format("{0,-4} | {1,-30} | {2,-30} | {3}", seq, Truncate(item.QuestionName, 30), Truncate(item.DynamicsName, 30), item.Status));
        }

        _ui.ShowInfo("------------------------------------------------------------------------------------------");
        _ui.ShowInfo($"TOTALS: {result.TotalCrmRecords} CRM Records | {result.HiddenMergedCount} Merged Questions | {result.VisibleQuestionCount} Total Visible In Questionnaire");
        _ui.ShowInfo("------------------------------------------------------------------------------------------\n");
    }

    private string Truncate(string value, int maxLength)
    {
        if (string.IsNullOrEmpty(value) || value.Length <= maxLength) return value;
        return value.Substring(0, maxLength - 3) + "...";
    }
}