# Export and Review GPOs
The script will generate 2 different ouput in csv format:  
1 containing the Disabled GPO *(DisabledGPOSummary.csv)*  
1 containing a Summary for all Enabled GPO *(AllGPOSummary.csv)*  

For each GPO an export in XML format, a Backup, and a PolicyRules definition will be generated  

You can later on import all PolicyRules definition in the MSFT PolicyAnalyzer Tool to check for discrepancies/conflicts

### Components:

- Main PowerShell script
- MSFT PolicyAnalyzer Tool *(<a href="https://www.microsoft.com/en-us/download/details.aspx?id=55319" target="_blank">Download Link Latest Version</a>)*

### Execution:

```ruby
.\ExportGPOs.ps1 -domainame 'contoso.com' -location 'C:\GPOReports'
```

### Limitation:
The current version of Policy Analyzer covers most areas of Group Policy, but does not yet include support for analysis of
Group Policy Preferences, nor of startup, shutdown, logon, or logoff scripts
