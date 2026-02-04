using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

/// <summary>
/// Loads and provides all configuration settings (secrets, fetchXml, app settings, command-line args).
/// </summary>
public class ConfigurationService
{
    public string Url { get; private set; }
    public string ClientId { get; private set; }
    public string ClientSecret { get; private set; }
    public string Authority { get; private set; }
    public string ConnectString { get; private set; }
    public Dictionary<string, string> FetchXmlQueries { get; private set; } = new Dictionary<string, string>();
    public int PageSize { get; private set; } = 5000;
    public bool SimulationMode { get; private set; } = false;
    public string LogLevel { get; private set; } = "Info";
    public bool ProcessWostsFromFile { get; private set; } = false;
    public string WostIdsFileName { get; private set; } = "wost_ids.txt";
    public List<string> WostIdsFromFile { get; private set; } = new List<string>();
    public bool BackfillWorkOrderRefs { get; private set; } = false;
    public List<Guid> SpecificWostIds { get; private set; } = new List<Guid>();
    public bool LogToFile { get; private set; } = false;
    /// <summary>
    /// Optional date filter: only process WOSTs created before this date.
    /// Used with --created-before argument.
    /// </summary>
    public DateTime? CreatedBefore { get; private set; }

    /// <summary>
    /// Optional logical environment name (e.g. "dev", "prod") selected via --env
    /// or interactively when multiple environments are defined in secrets.config.
    /// </summary>
    public string EnvironmentName { get; private set; }

    public ConfigurationService(string[] args)
    {
        DetectEnvironmentFromArgs(args);
        LoadSecretsConfig();
        LoadFetchXmlQueries();
        ParseCommandLineArgs(args);
        BuildConnectionString();
    }

    private void DetectEnvironmentFromArgs(string[] args)
    {
        if (args == null || args.Length == 0)
            return;

        for (int i = 0; i < args.Length; i++)
        {
            var arg = args[i];

            if (arg.StartsWith("--env", StringComparison.OrdinalIgnoreCase))
            {
                string value = null;

                var parts = arg.Split(new[] { '=' }, 2);
                if (parts.Length == 2 && !string.IsNullOrWhiteSpace(parts[1]))
                {
                    value = parts[1].Trim().Trim('"');
                }
                else if (i + 1 < args.Length)
                {
                    value = args[i + 1].Trim().Trim('"');
                }

                if (!string.IsNullOrWhiteSpace(value))
                {
                    EnvironmentName = value;
                    break;
                }
            }
        }
    }

    private void LoadSecretsConfig()
    {
        string secretsPath = FindSecretsConfig();
        if (string.IsNullOrEmpty(secretsPath) || !File.Exists(secretsPath))
        {
            throw new FileNotFoundException(
                "Required config file 'secrets.config' was not found.\n" +
                "Please create it by copying 'secrets.config.template' and filling in your credentials.\n" +
                $"Expected location (checked in order):\n" +
                $"  1. Project root: {GetExpectedSecretsPath()}\n" +
                $"  2. Bin directory: {Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "secrets.config")}\n" +
                $"  3. Solution root: {GetSolutionRootPath()}");
        }

        XDocument doc;
        try
        {
            doc = XDocument.Load(secretsPath);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to read secrets.config: {ex}", ex);
        }

        var root = doc.Root;
        if (root == null)
            throw new InvalidOperationException("secrets.config is empty or invalid.");

        // Optional multi-environment support:
        // <appSettings>
        //   <environment name="Prod">
        //     <add key="Url" ... />
        //     ...
        //   </environment>
        //   <environment name="Dev">
        //     ...
        //   </environment>
        // </appSettings>
        var environmentElements = root.Elements("environment").ToList();

        if (environmentElements.Any())
        {
            XElement selectedEnvironmentElement;

            if (!string.IsNullOrWhiteSpace(EnvironmentName))
            {
                selectedEnvironmentElement = environmentElements
                    .FirstOrDefault(e =>
                        string.Equals(
                            (string)e.Attribute("name"),
                            EnvironmentName,
                            StringComparison.OrdinalIgnoreCase));

                if (selectedEnvironmentElement == null)
                {
                    var available = environmentElements
                        .Select(e => (string)e.Attribute("name") ?? "<unnamed>")
                        .ToArray();
                    throw new InvalidOperationException(
                        $"Environment '{EnvironmentName}' not found in secrets.config. " +
                        $"Available environments: {string.Join(", ", available)}");
                }
            }
            else
            {
                var envNames = environmentElements
                    .Select((e, index) => new
                    {
                        Element = e,
                        Name = (string)e.Attribute("name") ?? $"Environment{index + 1}",
                        Url = e.Elements("add")
                            .FirstOrDefault(add =>
                                string.Equals((string)add.Attribute("key"), "Url", StringComparison.OrdinalIgnoreCase))
                            ?.Attribute("value")?.Value ?? "N/A"
                    })
                    .ToList();

                Console.WriteLine("Multiple environments defined in secrets.config.");
                Console.WriteLine("Select environment to use:");
                for (int i = 0; i < envNames.Count; i++)
                {
                    Console.WriteLine($"  {i + 1}. {envNames[i].Name} ({envNames[i].Url})");
                }
                Console.Write("Enter number or name (default = 1): ");
                string input = Console.ReadLine();

                if (string.IsNullOrWhiteSpace(input))
                {
                    selectedEnvironmentElement = envNames[0].Element;
                    EnvironmentName = envNames[0].Name;
                }
                else if (int.TryParse(input, out int indexChoice) &&
                         indexChoice >= 1 &&
                         indexChoice <= envNames.Count)
                {
                    selectedEnvironmentElement = envNames[indexChoice - 1].Element;
                    EnvironmentName = envNames[indexChoice - 1].Name;
                }
                else
                {
                    var match = envNames.FirstOrDefault(e =>
                        string.Equals(e.Name, input, StringComparison.OrdinalIgnoreCase));
                    if (match == null)
                    {
                        throw new InvalidOperationException(
                            $"Unknown environment selection '{input}'. " +
                            $"Available environments: {string.Join(", ", envNames.Select(e => e.Name))}");
                    }

                    selectedEnvironmentElement = match.Element;
                    EnvironmentName = match.Name;
                }

                Console.WriteLine($"Using environment: {EnvironmentName}");
            }

            LoadSecretsFromAddElements(selectedEnvironmentElement.Elements("add"));
        }
        else
        {
            // Backwards-compatible single-environment format:
            // <appSettings>
            //   <add key="Url" ... />
            //   ...
            // </appSettings>
            LoadSecretsFromAddElements(root.Elements("add"));
        }
    }

    private void LoadSecretsFromAddElements(IEnumerable<XElement> addElements)
    {
        foreach (var addElement in addElements)
        {
            string key = addElement.Attribute("key")?.Value;
            string value = addElement.Attribute("value")?.Value;

            if (!string.IsNullOrEmpty(key) && value != null)
            {
                switch (key.ToLower())
                {
                    case "url":
                        Url = value;
                        break;
                    case "clientid":
                        ClientId = value;
                        break;
                    case "clientsecret":
                        ClientSecret = value;
                        break;
                    case "authority":
                        Authority = value;
                        break;
                }
            }
        }
    }

    private void LoadFetchXmlQueries()
    {
        string exeDir = AppDomain.CurrentDomain.BaseDirectory;
        string settingsPath = Path.Combine(exeDir, "settings.config");
        if (!File.Exists(settingsPath))
            throw new FileNotFoundException($"Required config file 'settings.config' not found at {settingsPath}");
        XDocument doc = XDocument.Load(settingsPath);
        var settingsElement = doc.Root.Element("settings");
        if (settingsElement != null)
        {
            foreach (var addElement in settingsElement.Elements("add"))
            {
                string key = addElement.Attribute("key")?.Value;
                string value = addElement.Attribute("value")?.Value;
                if (!string.IsNullOrEmpty(key) && value != null)
                {
                    switch (key.ToLower())
                    {
                        case "pagesize":
                            if (int.TryParse(value, out int ps) && ps > 0) PageSize = ps;
                            break;
                        case "loglevel":
                            LogLevel = value;
                            break;
                    }
                }
            }
        }
        var queries = doc.Descendants("query");
        foreach (var query in queries)
        {
            string name = query.Attribute("name")?.Value;
            var fetchXmlElem = query.Elements().FirstOrDefault();
            string fetchXml = fetchXmlElem != null ? fetchXmlElem.ToString() : null;
            if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(fetchXml))
                FetchXmlQueries[name] = fetchXml;
        }
    }

    private void ParseCommandLineArgs(string[] args)
    {
        if (args == null) return;
        for (int i = 0; i < args.Length; i++)
        {
            var arg = args[i];

            if (arg.Equals("--debug", StringComparison.OrdinalIgnoreCase))
                LogLevel = "Debug";
            else if (arg.Equals("--verbose", StringComparison.OrdinalIgnoreCase))
                LogLevel = "Verbose";
            else if (arg.Equals("--quiet", StringComparison.OrdinalIgnoreCase))
                LogLevel = "Warning";
            else if (arg.Equals("--simulate", StringComparison.OrdinalIgnoreCase))
                SimulationMode = true;
            else if (arg.Equals("--backfill-workorder", StringComparison.OrdinalIgnoreCase))
                BackfillWorkOrderRefs = true;
            else if (arg.Equals("--from-file", StringComparison.OrdinalIgnoreCase))
                ProcessWostsFromFile = true;
            else if (arg.Equals("--logtofile", StringComparison.OrdinalIgnoreCase))
                LogToFile = true;
            else if (arg.StartsWith("--created-before", StringComparison.OrdinalIgnoreCase) ||
                     arg.StartsWith("--since", StringComparison.OrdinalIgnoreCase))
            {
                string argName = arg.StartsWith("--since", StringComparison.OrdinalIgnoreCase) ? "--since" : "--created-before";
                string value = null;
                var parts = arg.Split(new[] { '=' }, 2);
                if (parts.Length == 2 && !string.IsNullOrWhiteSpace(parts[1]))
                {
                    value = parts[1].Trim().Trim('"');
                }
                else if (i + 1 < args.Length && !args[i + 1].StartsWith("--"))
                {
                    value = args[i + 1].Trim().Trim('"');
                }

                if (!string.IsNullOrWhiteSpace(value))
                {
                    if (DateTime.TryParse(value, out DateTime parsedDate))
                    {
                        CreatedBefore = parsedDate;
                    }
                    else
                    {
                        throw new ArgumentException($"Invalid date format for {argName}: {value}. Use formats like: 2023-05-25, 2023-05-25T19:36:41, or 2023-05-25T19:36:41Z");
                    }
                }
                else
                {
                    throw new ArgumentException($"{argName} requires a date value.");
                }
            }
            else if (arg.StartsWith("--guids", StringComparison.OrdinalIgnoreCase) ||
                     arg.StartsWith("--wost-ids", StringComparison.OrdinalIgnoreCase))
            {
                string value = null;
                var parts = arg.Split(new[] { '=' }, 2);
                if (parts.Length == 2 && !string.IsNullOrWhiteSpace(parts[1]))
                {
                    value = parts[1].Trim().Trim('"');
                }
                else if (i + 1 < args.Length && !args[i + 1].StartsWith("--"))
                {
                    value = args[i + 1].Trim().Trim('"');
                }

                if (!string.IsNullOrWhiteSpace(value))
                {
                    var guids = ParseGuids(value);
                    if (guids.Count > 0)
                    {
                        SpecificWostIds.AddRange(guids);
                    }
                    else
                    {
                        throw new ArgumentException($"No valid GUIDs found in: {value}");
                    }
                }
            }
            else if (arg.StartsWith("--page-size", StringComparison.OrdinalIgnoreCase))
            {
                string value = null;
                var parts = arg.Split(new[] { '=' }, 2);
                if (parts.Length == 2 && !string.IsNullOrWhiteSpace(parts[1]))
                {
                    value = parts[1].Trim().Trim('"');
                }
                else if (i + 1 < args.Length && !args[i + 1].StartsWith("--"))
                {
                    value = args[i + 1].Trim().Trim('"');
                }

                if (!string.IsNullOrWhiteSpace(value) && int.TryParse(value, out int pageSize) && pageSize > 0)
                {
                    PageSize = pageSize;
                }
                else if (!string.IsNullOrWhiteSpace(value))
                {
                    throw new ArgumentException($"Invalid page size: {value}. Must be a positive integer.");
                }
            }
        }
    }

    private void BuildConnectionString()
    {
        if (string.IsNullOrEmpty(Url) || string.IsNullOrEmpty(ClientId) || string.IsNullOrEmpty(ClientSecret))
            throw new InvalidOperationException("Missing required CRM connection information in secrets.config");
        ConnectString = $"AuthType=ClientSecret;url={Url};ClientId={ClientId};ClientSecret={ClientSecret}";
    }

    public string GetActiveFetchXml()
    {
        string baseFetchXml = FetchXmlQueries.ContainsKey("UnprocessedQuestionnaires") ? FetchXmlQueries["UnprocessedQuestionnaires"] : null;
        
        if (baseFetchXml != null && CreatedBefore.HasValue)
        {
            // Inject createdon condition into the FetchXML filter
            baseFetchXml = InjectCreatedBeforeCondition(baseFetchXml, CreatedBefore.Value);
        }
        
        return baseFetchXml;
    }

    /// <summary>
    /// Injects a createdon condition into the FetchXML filter.
    /// </summary>
    private string InjectCreatedBeforeCondition(string fetchXml, DateTime createdBefore)
    {
        try
        {
            XDocument doc = XDocument.Parse(fetchXml);
            var filterElement = doc.Descendants("filter").FirstOrDefault();
            
            if (filterElement != null)
            {
                // Format date in ISO 8601 format for Dynamics 365
                string dateValue = createdBefore.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ");
                
                var createdOnCondition = new XElement("condition",
                    new XAttribute("attribute", "createdon"),
                    new XAttribute("operator", "lt"),
                    new XAttribute("value", dateValue));
                
                filterElement.Add(createdOnCondition);
            }
            
            return doc.ToString();
        }
        catch
        {
            // If XML parsing fails, return the original
            return fetchXml;
        }
    }

    /// <summary>
    /// Parses a comma-separated list of GUIDs, with or without braces.
    /// Examples: "guid1,guid2" or "{guid1},{guid2}" or "guid1,{guid2}"
    /// </summary>
    private static List<Guid> ParseGuids(string value)
    {
        var guids = new List<Guid>();
        if (string.IsNullOrWhiteSpace(value))
            return guids;

        var parts = value.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var part in parts)
        {
            var trimmed = part.Trim().Trim('{', '}');
            if (Guid.TryParse(trimmed, out Guid guid))
            {
                guids.Add(guid);
            }
        }
        return guids;
    }

    /// <summary>
    /// Finds secrets.config using fallback chain:
    /// 1. Project root (where .csproj is)
    /// 2. Bin directory (where exe runs)
    /// 3. Solution root (where .sln is)
    /// </summary>
    private static string FindSecretsConfig()
    {
        string exeDir = AppDomain.CurrentDomain.BaseDirectory;
        DirectoryInfo dir = new DirectoryInfo(exeDir);

        // 1. Check project root (walk up from exe to find .csproj)
        DirectoryInfo projectRoot = FindProjectRoot(dir);
        if (projectRoot != null)
        {
            string secretsPath = Path.Combine(projectRoot.FullName, "secrets.config");
            if (File.Exists(secretsPath))
            {
                return secretsPath;
            }
        }

        // 2. Check bin directory (where exe runs)
        string binSecretsPath = Path.Combine(exeDir, "secrets.config");
        if (File.Exists(binSecretsPath))
        {
            return binSecretsPath;
        }

        // 3. Check solution root (walk up to find .sln)
        DirectoryInfo solutionRoot = FindSolutionRoot(dir);
        if (solutionRoot != null)
        {
            string secretsPath = Path.Combine(solutionRoot.FullName, "secrets.config");
            if (File.Exists(secretsPath))
            {
                return secretsPath;
            }
        }

        return null;
    }

    /// <summary>
    /// Finds the project root by walking up from the exe directory to find a .csproj file.
    /// </summary>
    private static DirectoryInfo FindProjectRoot(DirectoryInfo startDir)
    {
        DirectoryInfo dir = startDir;
        while (dir != null)
        {
            if (dir.GetFiles("*.csproj").Any())
            {
                return dir;
            }
            dir = dir.Parent;
        }
        return null;
    }

    /// <summary>
    /// Finds the solution root by walking up from the exe directory to find a .sln file.
    /// </summary>
    private static DirectoryInfo FindSolutionRoot(DirectoryInfo startDir)
    {
        DirectoryInfo dir = startDir;
        while (dir != null)
        {
            if (dir.GetFiles("*.sln").Any())
            {
                return dir;
            }
            dir = dir.Parent;
        }
        return null;
    }

    /// <summary>
    /// Gets the expected path for secrets.config (for error messages).
    /// Returns the first location in the fallback chain where the file should be placed.
    /// </summary>
    private static string GetExpectedSecretsPath()
    {
        string exeDir = AppDomain.CurrentDomain.BaseDirectory;
        DirectoryInfo dir = new DirectoryInfo(exeDir);

        // Prefer project root
        DirectoryInfo projectRoot = FindProjectRoot(dir);
        if (projectRoot != null)
        {
            return Path.Combine(projectRoot.FullName, "secrets.config");
        }

        // Fallback to solution root
        DirectoryInfo solutionRoot = FindSolutionRoot(dir);
        if (solutionRoot != null)
        {
            return Path.Combine(solutionRoot.FullName, "secrets.config");
        }

        return "secrets.config (at project root)";
    }

    /// <summary>
    /// Gets the solution root path for error messages.
    /// </summary>
    private static string GetSolutionRootPath()
    {
        string exeDir = AppDomain.CurrentDomain.BaseDirectory;
        DirectoryInfo dir = new DirectoryInfo(exeDir);
        DirectoryInfo solutionRoot = FindSolutionRoot(dir);
        if (solutionRoot != null)
        {
            return Path.Combine(solutionRoot.FullName, "secrets.config");
        }
        return "secrets.config (at solution root)";
    }
}