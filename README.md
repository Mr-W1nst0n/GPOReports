# Export and Review GPOs

### Components:

- Main PowerShell script
- MSFT PolicyAnalyzer Tool

### Execution:

```ruby
.\ExportGPOs.ps1 -domainame 'contoso.com' -location 'C:\GPOReports'
```

### Limitation:
The current version of Policy Analyzer covers most areas of Group Policy, but does not yet include support for analysis of  
Group Policy Preferences, nor of startup, shutdown, logon, or logoff scripts
