# TSIS2.QuestionnaireProcessorConsole

## Overview

This console application connects to Dynamics 365 to process Work Order Service Tasks (WOST) with completed questionnaires. It retrieves questionnaire responses and transforms them into structured question response records. It supports processing specific records, date filtering, and simulation mode.

## Build Instructions
1. Open the solution in Visual Studio.
2. Ensure the project `TSIS2.QuestionnaireProcessorConsole` is selected.
3. Build the solution (Ctrl+Shift+B).
4. The executable `QuestionnaireProcessorConsole.exe` will be located in the `bin\Debug` or `bin\Release` folder.

## Setup Instructions

### 1. Prerequisites
- .NET Framework 4.7.1 or higher
- Access to the Dynamics 365 instance

### 2. Configuration Files

The application requires two configuration files to run: `secrets.config` and `settings.config`.

#### `secrets.config`
This file stores your Dynamics 365 connection credentials. You must manually create this file, as it is excluded from version control for security.
You can place this file in the **Project Root**, **Bin Directory**, or **Solution Root**.

**Template:**
```xml
<appSettings>
  <!-- Single Environment -->
  <add key="Url" value="https://your-crm-instance.crm3.dynamics.com/" />
  <add key="ClientId" value="your-client-id" />
  <add key="ClientSecret" value="your-client-secret" />
  <add key="Authority" value="https://login.microsoftonline.com/your-tenant-id" />
</appSettings>
```

**Multiple Environments:**
You can also define multiple environments and select one at runtime using the `--env` argument.
```xml
<appSettings>
  <environment name="Dev">
    <add key="Url" value="https://dev-instance.crm3.dynamics.com/" />
    <!-- ... other keys ... -->
  </environment>
  <environment name="Prod">
    <add key="Url" value="https://prod-instance.crm3.dynamics.com/" />
    <!-- ... other keys ... -->
  </environment>
</appSettings>
```

#### `settings.config`
This file controls application behavior and FetchXML queries. It must be present in the application execution directory.

**Example:**
```xml
<?xml version="1.0" encoding="utf-8" ?>
<appSettings>
  <settings>
    <add key="PageSize" value="5000" />
    <add key="LogLevel" value="Info" />
  </settings>

  <query name="UnprocessedQuestionnaires">
    <!-- Your FetchXML query here -->
     <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
      <entity name='msdyn_workorderservicetask'>
        <!-- ... attributes and filters ... -->
      </entity>
    </fetch>
  </query>
</appSettings>
```

## Command-Line Usage

Run the application from the command line with the following options:

### General
| Argument | Description |
|----------|-------------|
| `--help`, `-h`, `/?` | Show help message. |
| `--env <Name>` | Select a specific environment from `secrets.config`. |
| `--simulate` | Run in simulation mode (no records created/updated). |
| `--page-size <Num>` | Override the default page size for queries. |

### Logging
| Argument | Description |
|----------|-------------|
| `--debug` | Set log level to Debug (most verbose). |
| `--verbose` | Set log level to Verbose. |
| `--quiet` | Set log level to Warning (errors only). |
| `--logtofile` | Enable logging to a file in addition to the console. |

### Processing Modes
| Argument | Description |
|----------|-------------|
| `--created-before <Date>` | Only process WOSTs created before this date (e.g., `2023-01-01`). |
| `--guids <GuidList>` | Process specific WOST IDs (comma-separated). <br>Example: `--guids 35505ff6-347e-ed11-81ad-000d3af4f553,GUID2` |
| `--from-file [filename]` | Process WOST IDs listed in a file (one GUID per line). Defaults to `wost_ids.txt` if filename not specified. |
| `--backfill-workorder` | Run a backfill operation to populate the `ts_workorder` lookup on existing Question Response records. |

## File-Based Processing
To process a specific list of Work Order Service Tasks via file:
1. Create a file named `wost_ids.txt` (or custom name) in the application directory.
2. Add one WOST GUID per line.
3. Run the application with `--from-file` (or `--from-file custom_list.txt`).

## Backfill Operation
The application includes a utility to backfill the `ts_workorder` field on existing `ts_questionresponse` records.
- Use the `--backfill-workorder` argument to run this process.

**Why is this needed?**
This operation is useful for correcting data where Question Responses exist but lack a link to the parent Work Order. This scenario might happen if:
- The Questionnaire Processor encountered errors during previous runs.
- The logic to populate the Work Order was disabled, skipped, or not yet implemented.
- Records were created before the `ts_workorder` association was mandatory.

**Logic:**
- Finds Question Responses that have a linked Work Order Service Task (`ts_msdyn_workorderservicetask`) but are missing the Work Order reference (`ts_workorder`).
- It looks up the parent Work Order from the Service Task and populates the field.
- This creates no new records, only updates existing ones.

## Examples

**Run in simulation mode:**
```bash
QuestionnaireProcessorConsole.exe --simulate
```

**Process specific records:**
```bash
QuestionnaireProcessorConsole.exe --guids 35505ff6-347e-ed11-81ad-000d3af4f553,a1b2c3d4-e5f6-7890-1234-567890abcdef
```

**Run against "Prod" environment with file logging:**
```bash
QuestionnaireProcessorConsole.exe --env Prod --logtofile
```

**Process old records only:**
```bash
QuestionnaireProcessorConsole.exe --created-before 2023-01-01
```

**Process via file input (defaults to `wost_ids.txt`):**
```bash
QuestionnaireProcessorConsole.exe --from-file
```

**Process via custom file input:**
```bash
QuestionnaireProcessorConsole.exe --from-file my_list.txt
```

**Backfill Work Order references:**
```bash
QuestionnaireProcessorConsole.exe --backfill-workorder