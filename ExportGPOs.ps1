param (
    [Parameter(Mandatory = $true)]
    [System.String]
    $domainame,

    [Parameter(Mandatory = $true)]
    [System.String]
    $location
)

Clear-Host
Set-Location -Path $PSScriptRoot

#WhatIf
$whatifpreference = $false

#Log
$date = Get-Date -Format "yyyy-MM-dd"
$t=$MyInvocation.MyCommand.Name.Replace('.ps1','')
$LogFilePath = $location + "\Logs" + "\${t}_${date}.log"

Start-Transcript -Path "${LogFilePath}" | Out-Null

#Create Folder Structure
Write-Host 'Create Folders Structure' -ForegroundColor White
New-Item ($location + "\GPOResults\GPOBackup") -ItemType Directory -Force | Out-Null
New-Item ($location + "\GPOResults\GPODetailed") -ItemType Directory -Force | Out-Null
New-Item ($location + "\GPOResults\GPOPolicyRules") -ItemType Directory -Force | Out-Null
New-Item ($location + "\GPOResults\GPOsOutdated") -ItemType Directory -Force | Out-Null
New-Item ($location + "\GPOResults\GPOSummary") -ItemType Directory -Force | Out-Null

#Get all GPOs Disabled or Partially Disabled (User or Computer object level)
Write-Host 'Create Report for Disabled or Partially Disabled GPOs' -ForegroundColor Red
$DisabledGPOs = Get-GPO -All -Domain $domainame | Where-Object {($_.GpoStatus -match 'disabled') -or ($_.GpoStatus -match 'AllSettingsDisabled')}
$DisabledGPOs | Export-Csv -LiteralPath ($location + "\GPOResults\GPOsOutdated\" + "DisabledGPOSummary.csv")

#Get all GPOs in the domain
$AllGPOs = Get-GPO -All -Domain $domainame | Where-Object {($_.GpoStatus -Notmatch 'disabled') -and ($_.GpoStatus -Notmatch 'AllSettingsDisabled')} 

#Do the Magic :)
foreach ($g in $AllGpos)
{
    #Export GPO in XML format (not necessary but helps to pinpoint issue)
    Write-Host "Exporting $($g.DisplayName) in XML" -ForegroundColor Yellow
    Get-GPOReport -ReportType Xml -Guid $g.Id -Path ($location + "\GPOResults\GPODetailed\" + $g.DisplayName + ".xml")
    
    #Backup GPOs
    Write-Host "Backup $($g.DisplayName)" -ForegroundColor Cyan
    New-Item ($location + "\GPOResults\GPOBackup\" + $g.DisplayName) -ItemType Directory -Force | Out-Null
    Backup-GPO -Name $g.DisplayName -Domain $domainame -Path ($location + "\GPOResults\GPOBackup\" + $g.DisplayName) | Out-Null

    #Automate PolicyRule Definition
    Write-Host "Generate PolicyRule Definition for $($g.DisplayName)" -ForegroundColor Magenta
    [string]$pathToGPO2PolicyRules = ($location + "\Tool\PolicyAnalyzer\GPO2PolicyRules.exe")
    [Array]$arguments = ($location + "\GPOResults\GPOBackup\" + $g.DisplayName), ($location + "\GPOResults\GPOPolicyRules\" + $g.DisplayName + ".PolicyRules")
    & $pathToGPO2PolicyRules $arguments 2>&1 | ForEach-Object{ "$_" } | Out-Null

    [xml]$Gpo = Get-GPOReport -ReportType Xml -Guid $g.Id
    
        #Check if the extensiondata contains value
        if ((-Not $Gpo.gpo.user.extensiondata) -and (-Not $Gpo.gpo.computer.extensiondata))
        {
            Write-Warning "$($Gpo.GPO.Name) is Empty and Should be excluded"
        }

        else
        {
            Write-Host "$($Gpo.GPO.Name) saved to CSV file"
            $ExportCSV = New-Object PSObject
            $ExportCSV | Add-Member -MemberType NoteProperty -name 'GPO Name' -value $Gpo.GPO.Name
            $ExportCSV | Add-Member -MemberType NoteProperty -name 'Comp-Ad' -value $Gpo.GPO.Computer.VersionDirectory
            $ExportCSV | Add-Member -MemberType NoteProperty -name 'Comp-Sys' -value $Gpo.GPO.Computer.VersionSysvol
            $ExportCSV | Add-Member -MemberType NoteProperty -name 'Comp Enabled' -value $Gpo.GPO.Computer.Enabled
            $ExportCSV | Add-Member -MemberType NoteProperty -name 'User-Ad' -value $Gpo.GPO.User.VersionDirectory
            $ExportCSV | Add-Member -MemberType NoteProperty -name 'User-Sys' -value $Gpo.GPO.User.VersionSysvol
            $ExportCSV | Add-Member -MemberType NoteProperty -name 'User Enabled' -value $Gpo.GPO.User.Enabled
            $ExportCSV | Add-Member -MemberType NoteProperty -name 'Link Name' -value $Gpo.GPO.LinksTo.SOMName
            $ExportCSV | Add-Member -MemberType NoteProperty -name 'Link Path' -value $Gpo.GPO.LinksTo.SOMPath
            $ExportCSV | Add-Member -MemberType NoteProperty -name 'Link Enabled' -value $Gpo.GPO.LinksTo.Enabled
            $ExportCSV | Add-Member -MemberType NoteProperty -name 'Link Override' -value $Gpo.GPO.LinksTo.NoOverride
            $ExportCSV | Export-CSV ($location+"\GPOResults\GPOSummary\"+"AllGPOSummary.csv") -Encoding UTF8 -Delimiter ";" -Force -NoTypeInformation -Append
        }
}

Stop-Transcript | Out-Null
