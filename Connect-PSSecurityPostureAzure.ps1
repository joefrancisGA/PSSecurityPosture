#Requires -Version 7.0
<#
.SYNOPSIS
    Signs in for Export-DefenderRemReport.ps1 without the Windows "false account locked" popup.

.DESCRIPTION
    Azure PowerShell and Azure CLI 2.61+ on Windows default to Web Account Manager (WAM).
    WAM often shows "Your account has been locked" while portal.azure.com still works.

    This script disables the Windows broker and uses device-code sign-in in your browser only.

        .\Export-DefenderRemReport.ps1 -UseAzureCli

.EXAMPLE
    .\Connect-PSSecurityPostureAzure.ps1
#>
[CmdletBinding()]
param(
    [string] $TenantId = '5c4694b8-0490-46f0-bd5a-4eb70dd3c0b8'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($null -eq (Get-Command -Name az -ErrorAction SilentlyContinue))
{
    throw 'Azure CLI is not installed. Install from https://aka.ms/installazurecliwindows then rerun this script.'
}

Write-Host ''
Write-Host 'MedQuest sign-in (WAM disabled — browser device code only)' -ForegroundColor Cyan
Write-Host '-----------------------------------------------------------' -ForegroundColor Cyan
Write-Host 'If a small "account locked" popup appears, close it immediately.' -ForegroundColor Yellow
Write-Host 'Use ONLY the browser tab for https://login.microsoft.com/device (or aka.ms/devicelogin).' -ForegroundColor Yellow
Write-Host 'Sign in with the account that works at https://portal.azure.com (MedQuest / Hmd*).' -ForegroundColor White
Write-Host ''

& az account clear --only-show-errors 2>$null

# Azure CLI 2.61+ defaults to WAM on Windows — same false lock dialog as Connect-AzAccount.
& az config set core.enable_broker_on_windows=false --only-show-errors
if ($LASTEXITCODE -ne 0)
{
    Write-Warning 'Could not set core.enable_broker_on_windows=false (older Azure CLI?). Continuing with --use-device-code.'
}

Write-Host 'Starting device-code login (no WAM broker)...' -ForegroundColor Cyan

& az login --tenant $TenantId --use-device-code --only-show-errors
if ($LASTEXITCODE -ne 0)
{
    throw @"
az login failed (exit $LASTEXITCODE).

Manual fix (run in order):
  az account clear
  az config set core.enable_broker_on_windows=false
  az login --tenant $TenantId --use-device-code

Complete sign-in only in the browser device page, not the Windows popup.
"@
}

Write-Host ''
Write-Host 'Subscriptions visible to this login (Hmd*):' -ForegroundColor Green
& az account list `
    --query "[?starts_with(name, 'Hmd') || starts_with(name, 'HMD')].{Name:name, Id:id, State:state}" `
    --only-show-errors `
    -o table

Write-Host ''
Write-Host 'Next:' -ForegroundColor Green
Write-Host '  cd C:\PSSecurityPosture' -ForegroundColor White
Write-Host '  .\Export-DefenderRemReport.ps1 -UseAzureCli' -ForegroundColor White
