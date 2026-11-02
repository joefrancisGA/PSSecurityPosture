# PSSecurityPosture

PowerShell report that counts Microsoft Defender for Cloud recommendations by subscription so the totals can be compared with the Defender recommendations view.

`Export-DefenderRemReport.ps1` writes `JayREMWorkbook_RC2.xlsx`.

**MedQuest tenant:** `5c4694b8-0490-46f0-bd5a-4eb70dd3c0b8` (State of Hawaii MedQuest Division / hawaiianalyticsplatform.onmicrosoft.com). `Connect-PSSecurityPostureAzure.ps1` signs into this tenant by default.

```powershell
.\Connect-PSSecurityPostureAzure.ps1
.\Export-DefenderRemReport.ps1 -UseAzureCli

# Optional: log timestamps to a file (useful when a subscription page runs 10+ minutes)
.\Export-DefenderRemReport.ps1 -UseAzureCli -LogPath .\rem-collection.log
```

Without `-UseAzureCli`, the script uses `Search-AzGraph` after `Connect-AzAccount` (same WAM picker issue on many PCs).

Counts are distinct recommendation titles (`AssessmentName`) with status Unhealthy. Preview recommendations are excluded. Misconfigurations, Secrets, and Vulnerabilities match membership in `securityCategories`. An empty category is not counted as a misconfiguration.
