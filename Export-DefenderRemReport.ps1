#Requires -Version 7.0
[CmdletBinding()]
param(
    [switch] $UseAzureCli,

    [string] $LogPath = ''
)

Import-Module ImportExcel -ErrorAction Stop

if (-not $UseAzureCli)
{
    Import-Module Az.ResourceGraph -ErrorAction Stop
}

# ============================================================
# Configuration
# ============================================================

$ExcelFile = ".\JayREMWorkbook_RC2.xlsx"

$TenantLookup = @{
    "5c4694b8-0490-46f0-bd5a-4eb70dd3c0b8" =
        "State of Hawaii MedQuest Division"
}

# ============================================================
# Helper Functions
# ============================================================

function Write-PSSecurityPostureProgress
{
    param(
        [Parameter(Mandatory)]
        [string] $Message,

        [ConsoleColor] $Color = 'Gray'
    )

    [string]$line = '[{0}] {1}' -f (Get-Date -Format 'HH:mm:ss'), $Message
    Write-Host $line -ForegroundColor $Color

    if (-not [string]::IsNullOrWhiteSpace($LogPath))
    {
        Add-Content -LiteralPath $LogPath -Value $line -Encoding utf8
    }
}

function Get-RecommendationCount
{
    param($Collection)

    (
        $Collection |
        Select-Object -ExpandProperty AssessmentName -Unique
    ).Count
}

function Test-Category
{
    param(
        [string]$CategoryValue,
        [string]$CategoryName
    )

    if ([string]::IsNullOrWhiteSpace($CategoryValue))
    {
        return $false
    }

    try
    {
        $Parsed = $CategoryValue | ConvertFrom-Json

        if ($null -eq $Parsed)
        {
            return $false
        }

        return @($Parsed) -contains $CategoryName
    }
    catch
    {
        return $false
    }
}

function Get-SeverityCounts
{
    param($Collection)

    [PSCustomObject]@{
        Critical =
            ($Collection |
                Where-Object Severity -eq 'Critical' |
                Select-Object -ExpandProperty AssessmentName -Unique).Count

        High =
            ($Collection |
                Where-Object Severity -eq 'High' |
                Select-Object -ExpandProperty AssessmentName -Unique).Count

        Medium =
            ($Collection |
                Where-Object Severity -eq 'Medium' |
                Select-Object -ExpandProperty AssessmentName -Unique).Count

        Low =
            ($Collection |
                Where-Object Severity -eq 'Low' |
                Select-Object -ExpandProperty AssessmentName -Unique).Count
    }
}

function Get-HmdSubscriptionRecords
{
    param(
        [switch] $UseAzureCli
    )

    if ($UseAzureCli)
    {
        if ($null -eq (Get-Command -Name az -ErrorAction SilentlyContinue))
        {
            throw 'Azure CLI is not available. Run .\Connect-PSSecurityPostureAzure.ps1 or install Azure CLI.'
        }

        [object[]]$accounts =
            @(& az account list --only-show-errors | ConvertFrom-Json)

        return @(
            $accounts |
                Where-Object {
                    $_.name -like 'Hmd*' -or
                    $_.name -like 'HMD*'
                } |
                ForEach-Object {
                    [PSCustomObject]@{
                        Name = [string]$_.name
                        Id   = [string]$_.id
                    }
                }
        )
    }

    return @(
        Get-AzSubscription |
            Where-Object {
                $_.Name -like 'Hmd*' -or
                $_.Name -like 'HMD*'
            } |
            ForEach-Object {
                [PSCustomObject]@{
                    Name = [string]$_.Name
                    Id   = [string]$_.Id
                }
            }
    )
}

function Get-GraphDataViaAzureCli
{
    param(
        [Parameter(Mandatory)]
        [string] $Query,

        [Parameter(Mandatory)]
        [string] $SubscriptionId,

        [Parameter(Mandatory)]
        [string] $SubscriptionName
    )

    [System.Collections.ArrayList]$rows = [System.Collections.ArrayList]::new()
    [string]$skipToken = ''
    [int]$page = 0
    [int]$totalRows = 0
    [System.Diagnostics.Stopwatch]$subscriptionWatch = [System.Diagnostics.Stopwatch]::StartNew()

    do
    {
        $page++
        [System.Diagnostics.Stopwatch]$pageWatch = [System.Diagnostics.Stopwatch]::StartNew()

        Write-PSSecurityPostureProgress `
            "$SubscriptionName — Resource Graph page $page (waiting on Azure; large subs can take several minutes)..." `
            -Color DarkCyan

        [System.Collections.ArrayList]$azArgs = [System.Collections.ArrayList]::new()
        [void]$azArgs.Add('graph')
        [void]$azArgs.Add('query')
        [void]$azArgs.Add('-q')
        [void]$azArgs.Add($Query)
        [void]$azArgs.Add('--first')
        [void]$azArgs.Add('1000')
        [void]$azArgs.Add('--only-show-errors')
        [void]$azArgs.Add('--subscriptions')
        [void]$azArgs.Add($SubscriptionId)

        if (-not [string]::IsNullOrWhiteSpace($skipToken))
        {
            [void]$azArgs.Add('--skip-token')
            [void]$azArgs.Add($skipToken)
        }

        [string]$responseJson = (& az @azArgs 2>&1 | Out-String).Trim()

        if ($LASTEXITCODE -ne 0)
        {
            throw "az graph query failed for $SubscriptionName (page $page): $responseJson"
        }

        [object]$response = $responseJson | ConvertFrom-Json -ErrorAction Stop
        [int]$pageRows = 0

        foreach ($item in @($response.data))
        {
            if ($null -eq $item) { continue }

            [void]$rows.Add([PSCustomObject]@{
                tenantId           = [string]$item.tenantId
                SubscriptionName   = [string]$item.SubscriptionName
                SecurityCategory   = [string]$item.SecurityCategory
                Severity           = [string]$item.Severity
                ResourceType       = [string]$item.ResourceType
                AssessmentName     = [string]$item.AssessmentName
            })
            $pageRows++
        }

        $totalRows += $pageRows
        $pageWatch.Stop()

        Write-PSSecurityPostureProgress `
            ("$SubscriptionName — page {0}: +{1} row(s), {2} total ({3:0.0}s this page)" -f $page, $pageRows, $totalRows, $pageWatch.Elapsed.TotalSeconds) `
            -Color Green

        $skipToken = ''
        if ($null -ne $response.'$skipToken')
        {
            $skipToken = [string]$response.'$skipToken'
        }
    }
    while (-not [string]::IsNullOrWhiteSpace($skipToken))

    $subscriptionWatch.Stop()
    Write-PSSecurityPostureProgress `
        ("Finished $SubscriptionName — $totalRows row(s) in $page page(s), {0:0.0}s total" -f $subscriptionWatch.Elapsed.TotalSeconds) `
        -Color Cyan

    return @($rows.ToArray())
}

function Get-GraphData
{
    param(
        [Parameter(Mandatory)]
        [string]$Query,

        [switch] $UseAzureCli
    )

    $Results = @()

    $script:CollectionWarnings = @()

    $Subscriptions = Get-HmdSubscriptionRecords -UseAzureCli:$UseAzureCli

    if (@($Subscriptions).Count -eq 0)
    {
        if (-not $UseAzureCli -and (Get-Command -Name az -ErrorAction SilentlyContinue))
        {
            [object[]]$azHmd =
                @(& az account list --only-show-errors 2>$null | ConvertFrom-Json) |
                    Where-Object { $_.name -like 'Hmd*' -or $_.name -like 'HMD*' }

            if (@($azHmd).Count -gt 0)
            {
                throw @'
No Hmd* subscriptions in the Az PowerShell session, but Azure CLI is signed in.

Run the script file with -UseAzureCli (do not paste the script into the console):

    .\Export-DefenderRemReport.ps1 -UseAzureCli

Or sign in with Connect-AzAccount for tenant 5c4694b8-0490-46f0-bd5a-4eb70dd3c0b8 before running without -UseAzureCli.
'@
            }
        }

        throw 'No Hmd* subscriptions found. Sign in to tenant 5c4694b8-0490-46f0-bd5a-4eb70dd3c0b8 (Connect-PSSecurityPostureAzure.ps1 or az login --tenant ... --use-device-code), then run .\Export-DefenderRemReport.ps1 -UseAzureCli'
    }

    if ($UseAzureCli)
    {
        Write-Host ''
        Write-PSSecurityPostureProgress `
            "Collecting Defender assessments via Azure CLI for $(@($Subscriptions).Count) subscription(s)..." `
            -Color Cyan

        foreach ($Sub in @($Subscriptions))
        {
            [object[]]$batch =
                Get-GraphDataViaAzureCli `
                    -Query $Query `
                    -SubscriptionId $Sub.Id `
                    -SubscriptionName $Sub.Name

            if ($batch.Count -gt 0)
            {
                $Results += $batch
            }

            $script:CollectionWarnings +=
                [PSCustomObject]@{
                    Subscription    = $Sub.Name
                    RecordsReturned = $batch.Count
                    Status          = 'Azure CLI Resource Graph (paged)'
                }
        }

        return $Results
    }

    foreach ($Sub in $Subscriptions)
    {
        Write-Host ""
        Write-PSSecurityPostureProgress "Collecting $($Sub.Name) (Az.ResourceGraph)..." -Color Cyan

        $Skip = 0
        $TotalForSubscription = 0
        [int]$page = 0

        do
        {
            $page++
            Write-PSSecurityPostureProgress "$($Sub.Name) — page $page (Search-AzGraph)..." -Color DarkCyan

            [System.Diagnostics.Stopwatch]$pageWatch = [System.Diagnostics.Stopwatch]::StartNew()

            try
            {
                if ($Skip -eq 0)
                {
                    $Batch =
                        Search-AzGraph `
                            -Subscription $Sub.Id `
                            -Query $Query `
                            -First 1000
                }
                else
                {
                    $Batch =
                        Search-AzGraph `
                            -Subscription $Sub.Id `
                            -Query $Query `
                            -First 1000 `
                            -Skip $Skip
                }
            }
            catch
            {
                Write-Warning $_.Exception.Message
                break
            }

            $BatchCount =
                if ($Batch)
                {
                    $Batch.Count
                }
                else
                {
                    0
                }

            if ($BatchCount -gt 0)
            {
                $Results += $Batch
                $TotalForSubscription += $BatchCount
            }

            $pageWatch.Stop()
            Write-PSSecurityPostureProgress `
                ("$($Sub.Name) — page {0}: +{1} row(s), {2} total ({3:0.0}s)" -f $page, $BatchCount, $TotalForSubscription, $pageWatch.Elapsed.TotalSeconds) `
                -Color Green

            $Skip += 1000
        }
        while ($BatchCount -eq 1000)

        $script:CollectionWarnings +=
            [PSCustomObject]@{
                Subscription    = $Sub.Name
                RecordsReturned = $TotalForSubscription
                Status =
                    if ($TotalForSubscription -gt 1000)
                    {
                        'Paged Collection'
                    }
                    else
                    {
                        'Single Page'
                    }
            }
    }

    return $Results
}

# ============================================================
# ARG Query
# ============================================================

$Query = @"
securityresources
| where type =~ 'microsoft.security/assessments'
| where tostring(properties.status.code) == 'Unhealthy'
| where properties.metadata.preview != true
| extend
    SecurityCategory=tostring(properties.metadata.securityCategories),
    Severity=tostring(properties.metadata.severity),
    ResourceType=tolower(tostring(properties.resourceDetails.ResourceType)),
    AssessmentName=tostring(properties.displayName)
| join kind=leftouter (
    resourcecontainers
    | where type == 'microsoft.resources/subscriptions'
    | project
        subscriptionId,
        SubscriptionName=name,
        tenantId
) on subscriptionId
| project
    tenantId,
    SubscriptionName,
    SecurityCategory,
    Severity,
    ResourceType,
    AssessmentName
"@

# ============================================================
# Collect Data
# ============================================================

[System.Diagnostics.Stopwatch]$scriptWatch = [System.Diagnostics.Stopwatch]::StartNew()

Write-PSSecurityPostureProgress `
    "Starting collection (UseAzureCli=$UseAzureCli). Unhealthy assessment rows are counted per subscription; Hmd_HI_HAP_Non_Prod is usually the slowest." `
    -Color Cyan

$Data = Get-GraphData -Query $Query -UseAzureCli:$UseAzureCli

if (($null -eq $Data) -or ($Data.Count -eq 0))
{
    throw "No Defender data returned."
}

Write-PSSecurityPostureProgress `
    "ARG complete — $($Data.Count) row(s). Building REM report sections..." `
    -Color Cyan

Write-Host ""
Write-Host "Validation Snapshot"
Write-Host "==================="

$Data |
    Group-Object SubscriptionName |
    Sort-Object Count -Descending |
    Format-Table Count,Name -AutoSize

# ============================================================
# Build Report
# ============================================================

$OutputRows = @()
$ValidationRows = @()

$Subscriptions =
    $Data |
    Select-Object -ExpandProperty SubscriptionName -Unique

foreach ($Subscription in $Subscriptions)
{
    $SubData =
        $Data |
        Where-Object SubscriptionName -eq $Subscription

    $TenantId =
        ($SubData | Select-Object -First 1).tenantId

    $Tenant =
        if ($TenantLookup.ContainsKey($TenantId))
        {
            $TenantLookup[$TenantId]
        }
        else
        {
            $TenantId
        }

    $SecurityPosture = $SubData

    $Misconfigurations =
        $SubData |
        Where-Object {
            Test-Category -CategoryValue $_.SecurityCategory -CategoryName 'Misconfigurations'
        }

    $Secrets =
        $SubData |
        Where-Object {
            Test-Category -CategoryValue $_.SecurityCategory -CategoryName 'Secrets'
        }

    $Vulnerabilities =
        $SubData |
        Where-Object {
            Test-Category -CategoryValue $_.SecurityCategory -CategoryName 'Vulnerabilities'
        }

    $Machines =
        $Vulnerabilities |
        Where-Object {
            $_.ResourceType -eq 'microsoft.compute/virtualmachines'
        }

    $Containers =
        $Vulnerabilities |
        Where-Object {
            $_.ResourceType -match 'container'
        }

    $SQL =
        $Vulnerabilities |
        Where-Object {
            $_.ResourceType -match 'sql'
        }

    $Other =
        $Vulnerabilities |
        Where-Object {
            $_.ResourceType -ne 'microsoft.compute/virtualmachines' -and
            $_.ResourceType -notmatch 'container' -and
            $_.ResourceType -notmatch 'sql'
        }

    $Sections = @(
        @{ Type='Security Posture'; Data=$SecurityPosture }
        @{ Type='(1) Misconfigurations (CSPM)'; Data=$Misconfigurations }
        @{ Type='(2) Secrets (CSPM+CWPP)'; Data=$Secrets }
        @{ Type='(3) Vulnerabilities (CWPP)'; Data=$Vulnerabilities }
        @{ Type='(3.1) Machines (Servers)'; Data=$Machines }
        @{ Type='(3.2) Containers (Containers)'; Data=$Containers }
        @{ Type='(3.3) SQL (Databases)'; Data=$SQL }
        @{ Type='(3.4) Other'; Data=$Other }
    )

    foreach ($Section in $Sections)
    {
        $SeverityCounts =
            Get-SeverityCounts $Section.Data

        $RecommendationCount =
            Get-RecommendationCount $Section.Data

        $OutputRows += [PSCustomObject]@{
            Tenant         = $Tenant
            Subscription   = $Subscription
            Env            = ''
            Type           = $Section.Type
            Counts         = $RecommendationCount
            Critical       = $SeverityCounts.Critical
            High           = $SeverityCounts.High
            Medium         = $SeverityCounts.Medium
            Low            = $SeverityCounts.Low
            'Owning Team'  = ''
            'Owner POC(s)' = ''
            'Rem Plan'     = ''
        }
    }

    $SecurityPostureCount =
        Get-RecommendationCount $SecurityPosture

    $MisconfigCount =
        Get-RecommendationCount $Misconfigurations

    $SecretsCount =
        Get-RecommendationCount $Secrets

    $VulnCount =
        Get-RecommendationCount $Vulnerabilities

    $ReconciledTotal =
        $MisconfigCount +
        $SecretsCount +
        $VulnCount

    $ValidationRows +=
        [PSCustomObject]@{
            Subscription      = $Subscription
            SecurityPosture   = $SecurityPostureCount
            Misconfigurations = $MisconfigCount
            Secrets           = $SecretsCount
            Vulnerabilities   = $VulnCount
            ReconciledTotal   = $ReconciledTotal
            Difference        = $SecurityPostureCount - $ReconciledTotal
        }
}

# ============================================================
# Export Workbook
# ============================================================

Write-PSSecurityPostureProgress "Writing Excel workbook: $ExcelFile" -Color Cyan

$OutputRows |
    Export-Excel `
        -Path $ExcelFile `
        -WorksheetName 'REM Report' `
        -AutoSize `
        -AutoFilter `
        -BoldTopRow `
        -ClearSheet

$Data |
    Export-Excel `
        -Path $ExcelFile `
        -WorksheetName 'Raw ARG Data' `
        -AutoSize `
        -AutoFilter `
        -BoldTopRow

$ValidationRows |
    Export-Excel `
        -Path $ExcelFile `
        -WorksheetName 'Validation' `
        -AutoSize `
        -AutoFilter `
        -BoldTopRow

$script:CollectionWarnings |
    Export-Excel `
        -Path $ExcelFile `
        -WorksheetName 'Collection Warnings' `
        -AutoSize `
        -AutoFilter `
        -BoldTopRow

$Metadata = @(
    [PSCustomObject]@{
        Property='Generated'
        Value=(Get-Date)
    }
    [PSCustomObject]@{
        Property='Collection Method'
        Value=$(if ($UseAzureCli) { 'Azure CLI az graph query (browser sign-in via Connect-PSSecurityPostureAzure.ps1)' } else { 'Az.ResourceGraph Search-AzGraph paging' })
    }
    [PSCustomObject]@{
        Property='Counting Model'
        Value='Unique Defender Recommendations (AssessmentName), preview excluded'
    }
    [PSCustomObject]@{
        Property='RC2 Enhancement'
        Value='Queries continue until fewer than 1000 rows are returned'
    }
    [PSCustomObject]@{
        Property='Category Mapping'
        Value='Membership in securityCategories: Misconfigurations, Vulnerabilities, Secrets. Empty categories are not misconfigurations.'
    }
)

$Metadata |
    Export-Excel `
        -Path $ExcelFile `
        -WorksheetName 'Metadata' `
        -AutoSize `
        -BoldTopRow

Write-Host ""
Write-Host "Workbook created:"
Write-Host $ExcelFile

$scriptWatch.Stop()
Write-PSSecurityPostureProgress `
    ("Done — total elapsed {0:0.0} minutes" -f ($scriptWatch.Elapsed.TotalMinutes)) `
    -Color Cyan
