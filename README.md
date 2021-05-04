# ROM.TSIS2.Plugins

[![Build Status](https://dev.azure.com/transport-canada/Inspection%20Solution%20Architecture%20WG/_apis/build/status/ROMTS-GSRST/Create%20TSIS2.Plugins%20Artifact?branchName=main)](https://dev.azure.com/transport-canada/Inspection%20Solution%20Architecture%20WG/_build/latest?definitionId=484&branchName=main)

## Developer Instructions

### Installation
1. Use Visual Studio for the best experience.
2. After cloning the repository, open the solution file in Visual Studio.
3. Build the solution to restore the NuGet packages.

### Generating Xrm Types
To generate Type Declarations based on our Dynamics 365 solution, we use [XrmContext](https://github.com/delegateas/XrmContext) for the Plugins project and [XrmMockup](https://github.com/delegateas/XrmMockup) for the Tests project.

1. Create a `MetadataGenerator365.exe.config` file in ROMTS-GSRST.Plugins.Tests\Metadata

```xml
<!-- MetadataGenerator365.exe.config -->
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <appSettings>
    <add key="url" value="https://<Your Instance>.crm3.dynamics.com/XRMServices/2011/Organization.svc" />
    <add key="method" value="ClientSecret" />
    <add key="mfaAppId" value="<Your App ID>" />
    <add key="mfaClientSecret" value="<Your Client Secret>" />
    <add key="solutions" value="<Your solution>" />
    <add key="entities" value="<Comma delimited list of entities>" />
    <add key="servicecontextname" value="Xrm" />
    <add key="fetchFromAssemblies" value="false" />
  </appSettings>
  <startup>
    <supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.6.2" />
  </startup>
</configuration>
```

2. Create a `XrmContext.exe.config` file in TSIS2.Plugins\XrmContext

```xml
<!-- XrmContext.exe.config -->
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <appSettings>
    <add key="url" value="https://romts-gsrst-dev-tcd365.crm3.dynamics.com/XRMServices/2011/Organization.svc" />
    <add key="method" value="ClientSecret" />
    <add key="mfaAppId" value="<Your App ID>" />
    <add key="mfaClientSecret" value="<Your Client Secret>" />
    <add key="solutions" value="<Your solution>" />
    <add key="entities" value="<Comma delimited list of entities>" />
    <add key="servicecontextname" value="Xrm" />
  </appSettings>
</configuration>
```

3. After setting up the configuration files, you can run the `MetadataGenerator365.exe` or `XrmContext.exe.config` executables to generate the Types.  