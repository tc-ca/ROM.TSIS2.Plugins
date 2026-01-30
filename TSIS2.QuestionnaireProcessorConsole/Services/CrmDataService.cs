using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using TSIS2.QuestionnaireProcessorConsole;

public class CrmDataService
{
    private readonly IOrganizationService _service;
    private readonly TSIS2.Plugins.QuestionnaireProcessor.ILoggingService _logger;
    private readonly ConsoleUI _ui;
    private readonly int _pageSize;

    public CrmDataService(IOrganizationService service, TSIS2.Plugins.QuestionnaireProcessor.ILoggingService logger, ConsoleUI ui, int pageSize)
    {
        _service = service;
        _logger = logger;
        _ui = ui;
        _pageSize = pageSize;
    }

    public string EnsureFetchXmlHasOrder(string fetchXml)
    {
        try
        {
            XDocument doc = XDocument.Parse(fetchXml);
            var entity = doc.Descendants("entity").FirstOrDefault();
            if (entity != null && !entity.Elements("order").Any())
            {
                _logger.Debug("Adding order element to FetchXML for paging support");
                entity.Add(new XElement("order",
                    new XAttribute("attribute", "msdyn_name"),
                    new XAttribute("descending", "false")));
            }
            return doc.ToString();
        }
        catch (Exception ex)
        {
            _logger.Warning($"Error ensuring order in FetchXML: {ex}. Using original FetchXML.");
            return fetchXml;
        }
    }

    public string UpdateFetchXml(string originalXml, int pageNumber, string pagingCookie)
    {
        var doc = XDocument.Parse(originalXml);
        var fetch = doc.Root;
        fetch.Attribute("top")?.Remove();
        fetch.SetAttributeValue("page", pageNumber.ToString());
        fetch.SetAttributeValue("count", _pageSize.ToString());
        if (!string.IsNullOrEmpty(pagingCookie))
        {
            fetch.SetAttributeValue("paging-cookie", pagingCookie);
        }
        var entity = fetch.Element("entity");
        if (entity != null && !entity.Elements("order").Any())
        {
            _logger.Debug("Adding order element to FetchXML for paging support");
            entity.Add(new XElement("order",
                new XAttribute("attribute", "msdyn_name"),
                new XAttribute("descending", "false")));
        }
        return doc.ToString();
    }

    private int? ExtractTopValue(string fetchXml)
    {
        try
        {
            XDocument doc = XDocument.Parse(fetchXml);
            var fetch = doc.Root;
            var topAttr = fetch?.Attribute("top");
            if (topAttr != null && int.TryParse(topAttr.Value, out int topValue))
            {
                return topValue;
            }
        }
        catch
        {
            // Ignore parsing errors, return null
        }
        return null;
    }

    public List<Entity> RetrieveAllPages(string fetchXml, CancellationToken cancellationToken = default)
    {
        bool hasTopAttribute = fetchXml.Contains("top=");
        if (hasTopAttribute)
        {
            int? topValue = ExtractTopValue(fetchXml);
            if (topValue.HasValue)
            {
                _logger.Info($"Using 'top' attribute to limit records - paging disabled (top={topValue})");
            }
            else
            {
                _logger.Info("Using 'top' attribute to limit records - paging disabled");
            }
            var collection = _service.RetrieveMultiple(new FetchExpression(fetchXml));
            _logger.Info($"Retrieved {collection.Entities.Count} records with 'top' attribute");
            return collection.Entities.ToList();
        }

        var allRecords = new List<Entity>();
        int pageNumber = 1;
        string pagingCookie = null;
        int totalRecordCount = 0;
        int pageCount = 0;

        try
        {
            do
            {
                string updatedFetchXml = UpdateFetchXml(fetchXml, pageNumber, pagingCookie);
                var collection = _service.RetrieveMultiple(new FetchExpression(updatedFetchXml));
                allRecords.AddRange(collection.Entities);
                pagingCookie = collection.PagingCookie;
                pageNumber++;
                pageCount++;
                totalRecordCount += collection.Entities.Count;

                // Show paging progress
                _logger.Debug($"Retrieved page {pageCount}: {collection.Entities.Count} records (total so far: {totalRecordCount})");
                _ui.ShowPagingProgress(pageCount, collection.Entities.Count, totalRecordCount);

                if (collection.Entities.Count == 0 || string.IsNullOrEmpty(pagingCookie))
                {
                    break;
                }
            } while (true);

            // Show completion message
            _ui.ShowCompletionMessage(totalRecordCount, pageCount);

            _logger.Info($"Total records retrieved: {totalRecordCount} in {pageCount} pages");
            return allRecords;
        }
        catch (Exception ex)
        {
            _logger.Error($"Error during paged retrieval: {ex}");
            _logger.Debug($"FetchXML: {fetchXml}");
            throw;
        }
    }
}