# Export and Review GPOs

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
