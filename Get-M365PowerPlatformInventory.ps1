[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputCsvPath = (Join-Path -Path $PWD -ChildPath ("M365_PowerPlatform_Inventory_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))),

    [Parameter(Mandatory = $false)]
    [switch]$IncludeConnectionReferences,

    [Parameter(Mandatory = $false)]
    [switch]$SkipModuleInstall
        ,
        [Parameter(Mandatory = $false)]
    [string]$SpecifyEnvironment,

    [Parameter(Mandatory = $false)]
    [switch]$SkipAgentDataverseQuery
)

<#
Script: Get-M365PowerPlatformInventory.ps1

Purpose:
- Collects Power Platform inventory across environments (Flows, Apps, and Agents)
- Exports a consolidated CSV with owner, creation date, and optional connection/tool metadata

How to run:
1. Default output in current directory:
    .\Get-M365PowerPlatformInventory.ps1

2. Include connection references/tools:
    .\Get-M365PowerPlatformInventory.ps1 -IncludeConnectionReferences

3. Write to a specific CSV path:
    .\Get-M365PowerPlatformInventory.ps1 -OutputCsvPath .\inventory.csv

4. Run without auto-installing missing modules:
    .\Get-M365PowerPlatformInventory.ps1 -SkipModuleInstall

5. Skip Dataverse agent discovery and use PAC directly:
    .\Get-M365PowerPlatformInventory.ps1 -SkipAgentDataverseQuery

Parameters:
- OutputCsvPath (string)
  Path to the CSV output file. Defaults to M365_PowerPlatform_Inventory_yyyyMMdd_HHmmss.csv in the current folder.

- IncludeConnectionReferences (switch)
  When set, attempts to include Flow/App connection references and Agent tool/action summaries.

- SkipModuleInstall (switch)
  When set, missing modules are not installed automatically and the script errors if required modules are absent.

- SkipAgentDataverseQuery (switch)
    When set (and PAC fallback is active), bypasses all Dataverse-based agent queries and uses `pac copilot list --environment` directly.

PowerShell cmdlets and commands used:
- Module/bootstrap: Get-Module, Install-Module, Import-Module, Get-Command
- Power Platform auth/admin: Add-PowerAppsAccount, Get-AdminPowerAppEnvironment, Get-AdminEnvironment,
  Get-AdminFlow, Get-Flow, Get-AdminPowerApp, Get-AdminPowerAppConnectionReferences,
  Get-AdminPowerVirtualAgent, Get-AdminPowerVirtualAgents, Get-AdminCopilotStudioAgent, Get-AdminPowerAppCopilot
- Microsoft Graph (owner resolution): Get-MgContext, Connect-MgGraph, Get-MgUser
- Azure token acquisition for Dataverse: Get-AzContext, Connect-AzAccount, Get-AzAccessToken
- Dataverse API + export: Invoke-RestMethod, Export-Csv
- Background execution/timeouts for PAC calls: Start-Job, Wait-Job, Receive-Job, Stop-Job, Remove-Job
- External CLI fallback: pac (Power Platform CLI)
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:entraLookupInitialized = $false
$script:entraLookupEnabled = $false
$script:entraLookupFailureWarningShown = $false
$script:entraUserEmailCache = @{}

$script:dataverseTokenCache = @{}
$script:dataverseSystemUserEmailCache = @{}
$script:azAccountsAvailable = $false

# Ensures a required module is present and imported, with optional on-demand install.
function Ensure-Module {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [switch]$SkipInstall
    )

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        if ($SkipInstall) {
            throw "Required module '$Name' is not installed. Install it first or run without -SkipModuleInstall."
        }

        Write-Host "Installing module '$Name' for current user..." -ForegroundColor Yellow
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module -Name $Name -ErrorAction Stop -WarningAction SilentlyContinue
}

# Returns the first non-empty value found by traversing one of the provided dotted paths.
function Get-ValueByPath {
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string[]]$Paths
    )

    foreach ($path in $Paths) {
        $segments = $path -split "\."
        $current = $InputObject
        $found = $true

        foreach ($segment in $segments) {
            if ($null -eq $current) {
                $found = $false
                break
            }

            if ($current -is [System.Collections.IDictionary] -and $current.Contains($segment)) {
                $current = $current[$segment]
                continue
            }

            $prop = $current.PSObject.Properties[$segment]
            if ($null -eq $prop) {
                $found = $false
                break
            }

            $current = $prop.Value
        }

        if ($found -and $null -ne $current -and -not [string]::IsNullOrWhiteSpace([string]$current)) {
            return $current
        }
    }

    return $null
}

# Normalizes raw connection-reference payloads into a count and semicolon-delimited label list.
function Convert-ToConnectionReferenceSummary {
    param(
        [Parameter(Mandatory = $false)]
        [object]$Entries
    )

    $labels = [System.Collections.Generic.List[string]]::new()

    $addLabel = {
        param([string]$Candidate)
        if ([string]::IsNullOrWhiteSpace($Candidate)) {
            return
        }

        if (-not $labels.Contains($Candidate)) {
            $labels.Add($Candidate) | Out-Null
        }
    }

    $buildLabel = {
        param(
            [object]$Entry,
            [string]$Fallback
        )

        $display = [string](Get-ValueByPath -InputObject $Entry -Paths @(
                "displayName",
                "connectionReferenceDisplayName",
                "properties.displayName",
                "Name",
                "name"
            ))

        $logicalName = [string](Get-ValueByPath -InputObject $Entry -Paths @(
                "connectionReferenceLogicalName",
                "logicalName",
                "ConnectionReferenceLogicalName",
                "id",
                "Id"
            ))

        $connector = [string](Get-ValueByPath -InputObject $Entry -Paths @(
                "api.name",
                "connectorName",
                "properties.apiName"
            ))

        if ([string]::IsNullOrWhiteSpace($display)) {
            if (-not [string]::IsNullOrWhiteSpace($logicalName)) {
                $display = $logicalName
            }
            elseif (-not [string]::IsNullOrWhiteSpace($Fallback)) {
                $display = $Fallback
            }
        }

        if ([string]::IsNullOrWhiteSpace($display)) {
            return $null
        }

        if (-not [string]::IsNullOrWhiteSpace($connector) -and ($display -notlike "*$connector*")) {
            return "$display ($connector)"
        }

        return $display
    }

    if ($null -eq $Entries) {
        return [PSCustomObject]@{
            Count = 0
            List  = $null
        }
    }

    if ($Entries -is [System.Collections.IEnumerable] -and -not ($Entries -is [string]) -and -not ($Entries -is [System.Collections.IDictionary])) {
        foreach ($entry in $Entries) {
            $label = & $buildLabel $entry $null
            & $addLabel $label
        }

        $listText = $null
        if ($labels.Count -gt 0) {
            $listText = [string]::Join("; ", $labels)
        }

        return [PSCustomObject]@{
            Count = $labels.Count
            List  = $listText
        }
    }

    $entryProperties = @($Entries.PSObject.Properties | Where-Object { $_.MemberType -in @("NoteProperty", "Property") })
    if ($entryProperties.Count -gt 0) {
        $singleEntryFields = @("displayName", "name", "id", "apiName", "connectionReferenceLogicalName")
        $isSingleEntry = $false
        foreach ($fieldName in $singleEntryFields) {
            if ($Entries.PSObject.Properties[$fieldName]) {
                $isSingleEntry = $true
                break
            }
        }

        if (-not $isSingleEntry) {
            foreach ($prop in $entryProperties) {
                if ($null -eq $prop.Value) {
                    continue
                }

                $label = & $buildLabel $prop.Value ([string]$prop.Name)
                & $addLabel $label
            }
        }
    }

    if ($Entries -is [System.Collections.IDictionary]) {
        foreach ($key in $Entries.Keys) {
            $entry = $Entries[$key]
            $label = & $buildLabel $entry ([string]$key)
            & $addLabel $label
        }
    }
    else {
        $label = & $buildLabel $Entries $null
        & $addLabel $label
    }

    $listText = $null
    if ($labels.Count -gt 0) {
        $listText = [string]::Join("; ", $labels)
    }

    return [PSCustomObject]@{
        Count = $labels.Count
        List  = $listText
    }
}

# Extracts connection references directly from a Flow/App/Agent object when present.
function Get-ConnectionReferencesFromItem {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item
    )

    $entries = Get-ValueByPath -InputObject $Item -Paths @(
        "ConnectionReferences",
        "connectionReferences",
        "properties.connectionReferences",
        "Internal.properties.connectionReferences"
    )

    return (Convert-ToConnectionReferenceSummary -Entries $entries)
}

# Fetches app connection references via cmdlet first, then falls back to object-level extraction.
function Get-AppConnectionReferences {
    param(
        [Parameter(Mandatory = $true)]
        [string]$EnvironmentId,

        [Parameter(Mandatory = $true)]
        [string]$AppId,

        [Parameter(Mandatory = $true)]
        [object]$AppItem
    )

    try {
        $references = Get-AdminPowerAppConnectionReferences -EnvironmentName $EnvironmentId -AppName $AppId
        $summary = Convert-ToConnectionReferenceSummary -Entries $references
        if ($summary.Count -gt 0) {
            return $summary
        }
    }
    catch {
    }

    return (Get-ConnectionReferencesFromItem -Item $AppItem)
}

# Fetches flow connection references via flow details first, then falls back to object-level extraction.
function Get-FlowConnectionReferences {
    param(
        [Parameter(Mandatory = $true)]
        [string]$EnvironmentId,

        [Parameter(Mandatory = $true)]
        [string]$FlowId,

        [Parameter(Mandatory = $true)]
        [object]$FlowItem
    )

    try {
        $flowDetails = Get-Flow -EnvironmentName $EnvironmentId -FlowName $FlowId
        if ($flowDetails -is [System.Collections.IEnumerable] -and -not ($flowDetails -is [string])) {
            $flowDetails = @($flowDetails) | Select-Object -First 1
        }

        if ($null -ne $flowDetails) {
            $summary = Convert-ToConnectionReferenceSummary -Entries (Get-ValueByPath -InputObject $flowDetails -Paths @(
                    "Internal.properties.connectionReferences",
                    "Internal.properties.definition.parameters.`$connections.defaultValue",
                    "ConnectionReferences",
                    "connectionReferences"
                ))

            if ($summary.Count -gt 0) {
                return $summary
            }
        }
    }
    catch {
    }

    return (Get-ConnectionReferencesFromItem -Item $FlowItem)
}

# Builds an agent tool/action summary from Dataverse botcomponent payloads.
function Get-AgentToolSummary {
    param(
        [Parameter(Mandatory = $false)]
        [string]$EnvironmentUrl,

        [Parameter(Mandatory = $false)]
        [string]$AgentId
    )

    if ([string]::IsNullOrWhiteSpace($EnvironmentUrl) -or [string]::IsNullOrWhiteSpace($AgentId)) {
        return [PSCustomObject]@{
            Count = 0
            List  = $null
        }
    }

    $normalizedUrl = $EnvironmentUrl.TrimEnd('/')
    $token = Get-DataverseAccessToken -ResourceUrl $normalizedUrl
    if ([string]::IsNullOrWhiteSpace($token)) {
        return [PSCustomObject]@{
            Count = 0
            List  = $null
        }
    }

    try {
        $headers = @{
            "Authorization"    = "Bearer $token"
            "Accept"           = "application/json"
            "OData-MaxVersion" = "4.0"
            "OData-Version"    = "4.0"
        }

        $toolLabels = [System.Collections.Generic.List[string]]::new()

        $addTool = {
            param([string]$Value)
            if ([string]::IsNullOrWhiteSpace($Value)) {
                return
            }

            if (-not $toolLabels.Contains($Value)) {
                $toolLabels.Add($Value) | Out-Null
            }
        }

        $componentUri = "$normalizedUrl/api/data/v9.2/botcomponents?`$select=content,data,_parentbotid_value&`$filter=_parentbotid_value eq $AgentId"
        $components = @((Invoke-RestMethod -Uri $componentUri -Headers $headers -Method GET -ErrorAction Stop).value)

        foreach ($component in $components) {
            $contentBlob = ([string]$component.content + "`n" + [string]$component.data)
            if ([string]::IsNullOrWhiteSpace($contentBlob)) {
                continue
            }

            $kindMatches = [regex]::Matches($contentBlob, "(?i)\b(?:SearchAndSummarizeContent|Invoke[A-Za-z0-9]+(?:Action|TaskAction|MCP)|HttpRequestAction|GotoAction|PCFControlAction)\b")
            foreach ($kindMatch in $kindMatches) {
                $kind = [string]$kindMatch.Value
                if ([string]::IsNullOrWhiteSpace($kind)) {
                    continue
                }

                & $addTool ("action:{0}" -f $kind)
            }
        }

        $listText = $null
        if ($toolLabels.Count -gt 0) {
            $listText = [string]::Join("; ", $toolLabels)
        }

        return [PSCustomObject]@{
            Count = $toolLabels.Count
            List  = $listText
        }
    }
    catch {
        return [PSCustomObject]@{
            Count = 0
            List  = $null
        }
    }
}

# Converts a date-like value into ISO-8601 text while preserving non-date values as strings.
function Convert-ToIsoDate {
    param(
        [Parameter(Mandatory = $false)]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    try {
        return ([datetime]$Value).ToString("o")
    }
    catch {
        return [string]$Value
    }
}

# Resolves owner email/UPN candidates across known schema variants.
function Resolve-OwnerEmail {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item
    )

    $candidate = Get-ValueByPath -InputObject $Item -Paths @(
        "Owner.email",
        "Owner.userPrincipalName",
        "Owner.upn",
        "properties.owner.email",
        "properties.owner.userPrincipalName",
        "properties.creator.email",
        "properties.creator.userPrincipalName",
        "Creator.email",
        "Creator.userPrincipalName",
        "CreatedBy.email",
        "CreatedBy.userPrincipalName",
        "UserDetails.email",
        "UserDetails.userPrincipalName"
    )

    return $candidate
}

# Resolves owner display name candidates across known schema variants.
function Resolve-OwnerDisplayName {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item
    )

    return (Get-ValueByPath -InputObject $Item -Paths @(
            "Owner.displayName",
            "Owner.name",
            "properties.owner.displayName",
            "properties.creator.displayName",
            "Creator.displayName",
            "CreatedBy.displayName",
            "UserDetails.displayName"
        ))
}

# Extracts owner object IDs from multiple common payload shapes for flows/apps/agents.
function Get-OwnerObjectIdFromItem {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item
    )

    return [string](Get-ValueByPath -InputObject $Item -Paths @(
            "Owner.id",
            "Owner.objectId",
            "Owner.userId",
            "Owner.aadObjectId",
            "owner.id",
            "owner.objectId",
            "owner.userId",
            "owner.aadObjectId",
            "CreatedBy.id",
            "CreatedBy.objectId",
            "CreatedBy.userId",
            "CreatedBy.aadObjectId",
            "createdBy.id",
            "createdBy.objectId",
            "createdBy.userId",
            "createdBy.aadObjectId",
            "Creator.id",
            "Creator.objectId",
            "Creator.userId",
            "creator.id",
            "creator.objectId",
            "creator.userId",
            "properties.owner.id",
            "properties.owner.objectId",
            "properties.owner.userId",
            "properties.creator.id",
            "properties.creator.objectId",
            "properties.creator.userId",
            "properties.createdBy.id",
            "properties.createdBy.objectId",
            "properties.createdBy.userId",
            "UserDetails.id",
            "userDetails.id",
            "_ownerid_value",
            "ownerid"
        ))
}

# Formats owner output as "Display <Email>" when both are available.
function Resolve-OwnerField {
    param(
        [Parameter(Mandatory = $false)]
        [string]$OwnerDisplayName,

        [Parameter(Mandatory = $false)]
        [string]$OwnerEmail
    )

    if (-not [string]::IsNullOrWhiteSpace($OwnerDisplayName) -and -not [string]::IsNullOrWhiteSpace($OwnerEmail) -and $OwnerDisplayName -ne $OwnerEmail) {
        return "$OwnerDisplayName <$OwnerEmail>"
    }

    if (-not [string]::IsNullOrWhiteSpace($OwnerEmail)) {
        return $OwnerEmail
    }

    return $OwnerDisplayName
}

# Pulls a GUID out of an input value, including free-form strings that contain GUID text.
function Get-GuidFromValue {
    param(
        [Parameter(Mandatory = $false)]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [guid]) {
        return ([guid]$Value).ToString()
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    if ($text -match "(?i)[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}") {
        return $matches[0].ToLowerInvariant()
    }

    return $null
}

# Extracts an agent owner GUID from known owner/creator fields and fallback object blobs.
function Get-AgentOwnerGuidFromItem {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item
    )

    $ownerCandidate = Get-ValueByPath -InputObject $Item -Paths @(
        "Owner.id",
        "Owner.objectId",
        "Owner.aadObjectId",
        "Owner.userId",
        "Owner.principalId",
        "owner.id",
        "owner.objectId",
        "owner.aadObjectId",
        "owner.userId",
        "owner.principalId",
        "ownerObjectId",
        "OwnerObjectId",
        "OwnerId",
        "ownerId",
        "_ownerid_value",
        "properties.owner.id",
        "properties.owner.objectId",
        "properties.owner.aadObjectId",
        "properties.owner.userId",
        "properties.owner.principalId",
        "properties.ownerObjectId",
        "properties.OwnerObjectId",
        "properties.OwnerId",
        "properties.ownerId",
        "properties._ownerid_value",
        "CreatedBy.id",
        "CreatedBy.objectId",
        "CreatedBy.aadObjectId",
        "createdBy.id",
        "createdBy.objectId",
        "createdBy.aadObjectId",
        "properties.createdBy.id",
        "properties.createdBy.objectId",
        "properties.creator.id",
        "properties.creator.objectId",
        "Creator.id",
        "Creator.objectId",
        "creator.id",
        "creator.objectId",
        "UserDetails.id",
        "userDetails.id"
    )

    $ownerGuid = Get-GuidFromValue -Value $ownerCandidate
    if (-not [string]::IsNullOrWhiteSpace($ownerGuid)) {
        return $ownerGuid
    }

    # Some cmdlets expose owner metadata as free-form strings (for example, resource IDs).
    foreach ($path in @("Owner", "owner", "CreatedBy", "createdBy", "properties.owner", "properties.createdBy")) {
        $candidateValue = Get-ValueByPath -InputObject $Item -Paths @($path)
        $parsedGuid = Get-GuidFromValue -Value $candidateValue
        if (-not [string]::IsNullOrWhiteSpace($parsedGuid)) {
            return $parsedGuid
        }
    }

    return $null
}

# Initializes Microsoft Graph modules/session used for Entra user lookups.
function Initialize-EntraLookup {
    if ($script:entraLookupInitialized) {
        return
    }

    $script:entraLookupInitialized = $true

    try {
        Ensure-Module -Name "Microsoft.Graph.Authentication" -SkipInstall:$SkipModuleInstall
        Ensure-Module -Name "Microsoft.Graph.Users" -SkipInstall:$SkipModuleInstall

        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($null -eq $context) {
            Connect-MgGraph -Scopes "User.Read.All" -NoWelcome | Out-Null
        }

        $script:entraLookupEnabled = $true
        Write-Host "Entra owner lookup enabled for agents (Microsoft Graph)." -ForegroundColor DarkCyan
    }
    catch {
        $script:entraLookupEnabled = $false
        if (-not $script:entraLookupFailureWarningShown) {
            Write-Warning "Could not initialize Entra owner lookup. Agent owner emails will remain unresolved when only GUIDs are available. $($_.Exception.Message)"
            $script:entraLookupFailureWarningShown = $true
        }
    }
}

# Resolves an Entra email/UPN from an object ID with in-memory caching.
function Resolve-EntraEmailFromObjectId {
    param(
        [Parameter(Mandatory = $false)]
        [string]$ObjectId
    )

    if ([string]::IsNullOrWhiteSpace($ObjectId)) {
        return $null
    }

    if ($script:entraUserEmailCache.ContainsKey($ObjectId)) {
        return $script:entraUserEmailCache[$ObjectId]
    }

    Initialize-EntraLookup
    if (-not $script:entraLookupEnabled) {
        $script:entraUserEmailCache[$ObjectId] = $null
        return $null
    }

    try {
        $user = Get-MgUser -UserId $ObjectId -Property Id,Mail,UserPrincipalName -ErrorAction Stop
        $resolvedEmail = [string]$user.Mail
        if ([string]::IsNullOrWhiteSpace($resolvedEmail)) {
            $resolvedEmail = [string]$user.UserPrincipalName
        }

        $script:entraUserEmailCache[$ObjectId] = $resolvedEmail
        return $resolvedEmail
    }
    catch {
        $script:entraUserEmailCache[$ObjectId] = $null
        return $null
    }
}

# Safely converts secure tokens/secrets into plaintext for HTTP authorization headers.
function Convert-SecretToPlainText {
    param(
        [Parameter(Mandatory = $false)]
        [object]$SecretValue
    )

    if ($null -eq $SecretValue) {
        return $null
    }

    if ($SecretValue -is [Security.SecureString]) {
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecretValue)
        try {
            return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
        }
        finally {
            if ($bstr -ne [IntPtr]::Zero) {
                [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
            }
        }
    }

    return [string]$SecretValue
}

# Gets and caches an access token for Dataverse API calls scoped to an instance URL.
function Get-DataverseAccessToken {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ResourceUrl
    )

    $normalizedUrl = $ResourceUrl.TrimEnd('/')

    if ($script:dataverseTokenCache.ContainsKey($normalizedUrl)) {
        return $script:dataverseTokenCache[$normalizedUrl]
    }

    if ($script:azAccountsAvailable) {
        try {
            $result = Get-AzAccessToken -ResourceUrl $normalizedUrl -ErrorAction Stop
            $token = Convert-SecretToPlainText -SecretValue $result.Token
            if ([string]::IsNullOrWhiteSpace($token) -and $result.PSObject.Properties["AccessToken"]) {
                $token = Convert-SecretToPlainText -SecretValue $result.AccessToken
            }
            if (-not [string]::IsNullOrWhiteSpace($token)) {
                $script:dataverseTokenCache[$normalizedUrl] = $token
                return $token
            }
        }
        catch { }
    }

    $script:dataverseTokenCache[$normalizedUrl] = $null
    return $null
}

# Queries Dataverse bots and returns a map of bot ID to owner/date metadata.
function Get-DataverseBotOwnerMap {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InstanceUrl,

        [Parameter(Mandatory = $false)]
        [string]$EnvironmentDisplayName
    )

    $map = @{}
    $normalizedUrl = $InstanceUrl.TrimEnd('/')
    $token = Get-DataverseAccessToken -ResourceUrl $normalizedUrl

    if ([string]::IsNullOrWhiteSpace($token)) {
        return $map
    }

    try {
        $headers = @{
            "Authorization"    = "Bearer $token"
            "Accept"           = "application/json"
            "OData-MaxVersion" = "4.0"
            "OData-Version"    = "4.0"
            "Prefer"           = "odata.include-annotations=`"OData.Community.Display.V1.FormattedValue`""
        }

        $uri = "$normalizedUrl/api/data/v9.2/bots?`$select=botid,name,_ownerid_value,createdon&`$top=5000"
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method GET -ErrorAction Stop

        foreach ($bot in $response.value) {
            $botId = ([string]$bot.botid).ToLowerInvariant()
            if (-not [string]::IsNullOrWhiteSpace($botId)) {
                $ownerDisplayValue = $bot.PSObject.Properties["_ownerid_value@OData.Community.Display.V1.FormattedValue"]
                $map[$botId] = @{
                    OwnerSystemUserId = [string]$bot."_ownerid_value"
                    OwnerDisplayName  = if ($null -ne $ownerDisplayValue) { [string]$ownerDisplayValue.Value } else { $null }
                    CreatedOn         = [string]$bot.createdon
                }
            }
        }
    }
    catch {
        $errorMessage = $_.Exception.Message
        if ([string]::IsNullOrWhiteSpace($errorMessage)) {
            $errorMessage = "Dataverse bot owner map query failed."
        }

        if (-not [string]::IsNullOrWhiteSpace($EnvironmentDisplayName)) {
            Add-EnvironmentAccessIssue -Environment $EnvironmentDisplayName -Stage "DataverseOwnerMap" -Message $errorMessage
        }
    }

    return $map
}

# Queries Dataverse bots and returns agent inventory records for an environment.
function Get-DataverseBots {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InstanceUrl,

        [Parameter(Mandatory = $false)]
        [string]$EnvironmentDisplayName
    )

    $bots = [System.Collections.Generic.List[object]]::new()
    $normalizedUrl = $InstanceUrl.TrimEnd('/')
    $token = Get-DataverseAccessToken -ResourceUrl $normalizedUrl

    if ([string]::IsNullOrWhiteSpace($token)) {
        return $bots
    }

    try {
        $headers = @{
            "Authorization"    = "Bearer $token"
            "Accept"           = "application/json"
            "OData-MaxVersion" = "4.0"
            "OData-Version"    = "4.0"
            "Prefer"           = "odata.include-annotations=`"OData.Community.Display.V1.FormattedValue`""
        }

        $nextUri = "$normalizedUrl/api/data/v9.2/bots?`$select=botid,name,_ownerid_value,createdon&`$top=5000"
        while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
            $response = Invoke-RestMethod -Uri $nextUri -Headers $headers -Method GET -ErrorAction Stop

            foreach ($bot in $response.value) {
                $ownerDisplayValue = $bot.PSObject.Properties["_ownerid_value@OData.Community.Display.V1.FormattedValue"]
                $bots.Add([PSCustomObject]@{
                        BotId             = [string]$bot.botid
                        Name              = [string]$bot.name
                        OwnerSystemUserId = [string]$bot."_ownerid_value"
                        OwnerDisplayName  = if ($null -ne $ownerDisplayValue) { [string]$ownerDisplayValue.Value } else { $null }
                        CreatedOn         = [string]$bot.createdon
                    }) | Out-Null
            }

            $nextUri = $null
            if ($response.PSObject.Properties['@odata.nextLink']) {
                $nextUri = [string]$response.'@odata.nextLink'
            }
        }
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Warning "Dataverse agent discovery failed for '$normalizedUrl'. $errorMessage"
        if (-not [string]::IsNullOrWhiteSpace($EnvironmentDisplayName)) {
            Add-EnvironmentAccessIssue -Environment $EnvironmentDisplayName -Stage "DataverseAgentDiscovery" -Message $errorMessage
        }
    }

    return $bots
}

# Records environment access/connectivity failures for the end-of-run summary.
function Add-EnvironmentAccessIssue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Environment,

        [Parameter(Mandatory = $true)]
        [string]$Stage,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if ([string]::IsNullOrWhiteSpace($Environment) -or [string]::IsNullOrWhiteSpace($Message)) {
        return
    }

    $normalizedMessage = $Message.Trim()
    if ([string]::IsNullOrWhiteSpace($normalizedMessage)) {
        return
    }

    $issueKey = "{0}|{1}|{2}" -f $Environment.ToLowerInvariant(), $Stage.ToLowerInvariant(), $normalizedMessage.ToLowerInvariant()
    if (-not $script:environmentAccessIssueKeys.Contains($issueKey)) {
        $script:environmentAccessIssueKeys.Add($issueKey) | Out-Null
        $script:environmentAccessIssues.Add([PSCustomObject]@{
                Environment = $Environment
                Stage       = $Stage
                Message     = $normalizedMessage
            }) | Out-Null
    }
}

# Resolves a Dataverse system user email and caches results by environment+user key.
function Resolve-DataverseSystemUserEmail {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InstanceUrl,

        [Parameter(Mandatory = $true)]
        [string]$SystemUserId
    )

    if ([string]::IsNullOrWhiteSpace($SystemUserId)) {
        return $null
    }

    $cacheKey = "$InstanceUrl|$SystemUserId"
    if ($script:dataverseSystemUserEmailCache.ContainsKey($cacheKey)) {
        return $script:dataverseSystemUserEmailCache[$cacheKey]
    }

    $normalizedUrl = $InstanceUrl.TrimEnd('/')
    $token = Get-DataverseAccessToken -ResourceUrl $normalizedUrl

    if ([string]::IsNullOrWhiteSpace($token)) {
        $script:dataverseSystemUserEmailCache[$cacheKey] = $null
        return $null
    }

    try {
        $headers = @{
            "Authorization"    = "Bearer $token"
            "Accept"           = "application/json"
            "OData-MaxVersion" = "4.0"
            "OData-Version"    = "4.0"
        }

        $uri = "$normalizedUrl/api/data/v9.2/systemusers($SystemUserId)?`$select=internalemailaddress,fullname"
        $user = Invoke-RestMethod -Uri $uri -Headers $headers -Method GET -ErrorAction Stop
        $email = [string]$user.internalemailaddress
        $script:dataverseSystemUserEmailCache[$cacheKey] = $email
        return $email
    }
    catch {
        $script:dataverseSystemUserEmailCache[$cacheKey] = $null
        return $null
    }
}

# Returns the PAC CLI command if available in PATH.
function Get-PacCommand {
    return (Get-Command -Name "pac" -ErrorAction SilentlyContinue)
}

# Runs a PAC command in a background job and enforces a timeout to avoid hangs.
function Invoke-PacCommandSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments,

        [Parameter(Mandatory = $false)]
        [int]$TimeoutSeconds = 45
    )

    $job = Start-Job -ScriptBlock {
        param([string[]]$InnerArguments)
        pac @InnerArguments 2>&1 | Out-String
    } -ArgumentList (, $Arguments)

    try {
        $completed = Wait-Job -Job $job -Timeout $TimeoutSeconds
        if ($null -eq $completed) {
            Stop-Job -Job $job -Force | Out-Null
            return $null
        }

        return [string](Receive-Job -Job $job)
    }
    finally {
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    }
}

# Parses tabular `pac copilot list` output into structured name/id records.
function Parse-PacCopilotListOutput {
    param(
        [Parameter(Mandatory = $false)]
        [string]$OutputText
    )

    if ([string]::IsNullOrWhiteSpace($OutputText)) {
        return @()
    }

    $records = [System.Collections.Generic.List[object]]::new()
    $lines = $OutputText -split "`r?`n"
    $guidPattern = "[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"

    foreach ($line in $lines) {
        $trimmed = $line.TrimEnd()
        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            continue
        }

        # Never parse PAC errors/warnings as agent rows.
        if ($trimmed -match "^(?i)(error:|warning:|invalid value|failed|exception|the value passed to '--environment')") {
            continue
        }

        if ($trimmed -match "^(Connected as|Connected to\.\.\.|Microsoft PowerPlatform CLI|Version:|Online documentation:|Feedback, Suggestions, Issues:|Name\s+Copilot ID|---+)") {
            continue
        }

        $idMatch = [regex]::Match($trimmed, $guidPattern)
        if (-not $idMatch.Success) {
            continue
        }

        $copilotId = $idMatch.Value
        $name = $trimmed.Substring(0, $idMatch.Index).Trim()

        $records.Add([PSCustomObject]@{
                Name = $name
                CopilotId = $copilotId
            }) | Out-Null
    }

    return $records
}

# Gets PAC copilot rows for an environment using several argument formats.
function Get-PacCopilotAgentsForEnvironment {
    param(
        [Parameter(Mandatory = $true)]
        [string]$EnvironmentId,

        [Parameter(Mandatory = $false)]
        [string]$InstanceUrl,

        [Parameter(Mandatory = $false)]
        [string]$EnvironmentDisplayName
    )

    $candidateEnvironmentValues = [System.Collections.Generic.List[string]]::new()

    if (-not [string]::IsNullOrWhiteSpace($InstanceUrl)) {
        $candidateEnvironmentValues.Add($InstanceUrl.TrimEnd('/')) | Out-Null
    }

    if (-not [string]::IsNullOrWhiteSpace($EnvironmentId)) {
        $candidateEnvironmentValues.Add($EnvironmentId) | Out-Null

        if ($EnvironmentId -match "(?i)^Default-([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})$") {
            $candidateEnvironmentValues.Add($matches[1]) | Out-Null
        }
    }

    # De-duplicate while preserving order.
    $seen = @{}
    $orderedCandidates = @()
    foreach ($candidate in $candidateEnvironmentValues) {
        $key = $candidate.ToLowerInvariant()
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $orderedCandidates += $candidate
        }
    }

    foreach ($candidate in $orderedCandidates) {
        $output = Invoke-PacCommandSafe -Arguments @("copilot", "list", "--environment", $candidate)
        $records = @(Parse-PacCopilotListOutput -OutputText $output)
        if ($records.Count -gt 0) {
            return $records
        }

        if (-not [string]::IsNullOrWhiteSpace($output) -and $output -match "(?i)(error:|the value passed to '--environment' is invalid|no dataverse organization was found)") {
            Write-Warning "PAC agent lookup failed for environment '$EnvironmentDisplayName' using '--environment $candidate'."

            $pacErrorLine = ($output -split "`r?`n" | Where-Object { $_ -match "(?i)error:|no dataverse organization was found|the value passed to '--environment' is invalid" } | Select-Object -First 1)
            if ([string]::IsNullOrWhiteSpace($pacErrorLine)) {
                $pacErrorLine = "PAC agent lookup failed for '--environment $candidate'."
            }

            if (-not [string]::IsNullOrWhiteSpace($EnvironmentDisplayName)) {
                Add-EnvironmentAccessIssue -Environment $EnvironmentDisplayName -Stage "PacCopilotList" -Message $pacErrorLine
            }
        }
    }

    return @()
}

# Finds the first supported PowerShell cmdlet for listing environments.
function Get-EnvironmentCommand {
    $candidateCommands = @(
        "Get-AdminPowerAppEnvironment",
        "Get-AdminEnvironment"
    )

    foreach ($commandName in $candidateCommands) {
        $command = Get-Command -Name $commandName -ErrorAction SilentlyContinue
        if ($null -ne $command) {
            return $command
        }
    }

    return $null
}

# Adds a normalized artifact row into the in-memory inventory list.
function Add-InventoryRecord {
    param(
        [Parameter(Mandatory = $false)]
        [ValidateNotNull()]
        [System.Collections.Generic.List[object]]$Buffer,

        [Parameter(Mandatory = $true)]
        [string]$ArtifactType,

        [Parameter(Mandatory = $true)]
        [string]$Environment,

        [Parameter(Mandatory = $true)]
        [string]$EnvironmentId,

        [Parameter(Mandatory = $false)]
        [string]$Guid,

        [Parameter(Mandatory = $false)]
        [string]$ArtifactName,

        [Parameter(Mandatory = $false)]
        [string]$Owner,

        [Parameter(Mandatory = $false)]
        [string]$OwnerEmail,

        [Parameter(Mandatory = $false)]
        [string]$OwnerObjectId,

        [Parameter(Mandatory = $false)]
        [string]$DateCreated,

        [Parameter(Mandatory = $false)]
        [string]$ConnectionReferences,

        [Parameter(Mandatory = $false)]
        [int]$ConnectionReferenceCount = 0
    )

    if ($null -eq $Buffer) {
        throw "Buffer cannot be null."
    }

    $ownerToWrite = $Owner
    if ($ArtifactType -eq "Agent" -and -not [string]::IsNullOrWhiteSpace($OwnerObjectId)) {
        $ownerToWrite = $OwnerObjectId
    }

    $Buffer.Add([PSCustomObject]@{
            ArtifactType = $ArtifactType
            Environment  = $Environment
            EnvironmentId = $EnvironmentId
            GUID         = $Guid
            ArtifactName = $ArtifactName
            Owner        = $ownerToWrite
            OwnerEmail   = $OwnerEmail
            OwnerObjectId = $OwnerObjectId
            DateCreated  = $DateCreated
            ConnectionReferenceCount = $ConnectionReferenceCount
            ConnectionReferences = $ConnectionReferences
        }) | Out-Null
}

# Load required Power Platform modules before any data collection.
Ensure-Module -Name "Microsoft.PowerApps.Administration.PowerShell" -SkipInstall:$SkipModuleInstall
Ensure-Module -Name "Microsoft.PowerApps.PowerShell" -SkipInstall:$SkipModuleInstall

# Prepare Az context for Dataverse token-based enrichment (PAC fallback path).
try {
    Ensure-Module -Name "Az.Accounts" -SkipInstall:$SkipModuleInstall
    $azContext = Get-AzContext -ErrorAction SilentlyContinue
    if ($null -eq $azContext) {
        Write-Host "Signing in to Azure (Az.Accounts) for Dataverse token acquisition..." -ForegroundColor Cyan
        Write-Host "A device code will be displayed. Visit https://microsoft.com/devicelogin in your browser and enter the code." -ForegroundColor Yellow
        Connect-AzAccount -UseDeviceAuthentication -ErrorAction Stop | Out-Null
    }
    $script:azAccountsAvailable = $true
    Write-Host "Az.Accounts ready — Dataverse owner/date enrichment enabled for PAC fallback agents." -ForegroundColor DarkCyan
}
catch {
    Write-Warning "Az.Accounts module could not be loaded or authenticated. PAC fallback agents will have blank Owner/OwnerEmail/DateCreated. ($($_.Exception.Message))"
    $script:azAccountsAvailable = $false
}

Write-Host "Signing in to Power Platform..." -ForegroundColor Cyan
Add-PowerAppsAccount | Out-Null

# Discover tenant environments and cache environment URL lookups.
Write-Host "Collecting environments..." -ForegroundColor Cyan
$environmentCommand = Get-EnvironmentCommand
if ($null -eq $environmentCommand) {
    throw "No supported environment command was found. Expected one of: Get-AdminPowerAppEnvironment, Get-AdminEnvironment."
}

Write-Host "Using environment command: $($environmentCommand.Name)" -ForegroundColor DarkCyan
$environments = & $environmentCommand.Name

if (-not $environments) {
    throw "No environments found or you do not have sufficient permissions."
}

# If specifyEnvironment is set, filter environments to only that GUID
if ($SpecifyEnvironment) {
    $specEnvGuid = $SpecifyEnvironment.ToLowerInvariant()
    $environments = @($environments | Where-Object {
        $envId = [string](Get-ValueByPath -InputObject $_ -Paths @("EnvironmentName", "Name", "Id"))
        $envId -and ($envId.ToLowerInvariant() -eq $specEnvGuid)
    })
    if (-not $environments -or $environments.Count -eq 0) {
        throw "No environment found matching GUID: $SpecifyEnvironment"
    }
}

$totalEnvironmentCount = @($environments).Count
$processedEnvironmentCount = 0

$inventory = [System.Collections.Generic.List[object]]::new()
$flowCount = 0
$appCount = 0
$agentCount = 0
$environmentUrlById = @{}
$script:environmentAccessIssues = [System.Collections.Generic.List[object]]::new()
$script:environmentAccessIssueKeys = [System.Collections.Generic.HashSet[string]]::new()

foreach ($environment in $environments) {
    $mappedEnvironmentId = [string](Get-ValueByPath -InputObject $environment -Paths @("EnvironmentName", "Name", "Id"))
    $mappedInstanceUrl = [string](Get-ValueByPath -InputObject $environment -Paths @("Internal.properties.linkedEnvironmentMetadata.instanceUrl"))
    if (-not [string]::IsNullOrWhiteSpace($mappedEnvironmentId)) {
        $environmentUrlById[$mappedEnvironmentId] = $mappedInstanceUrl
    }
}

$pacCommand = Get-PacCommand
$usePacFallbackForAgents = $false
$pacFallbackOwnerDateMissingWarningShown = $false

# Agent inventory is collected through PAC/Dataverse path.
if ($null -ne $pacCommand) {
    $usePacFallbackForAgents = $true
    Write-Host "Using PAC-based agent discovery path (PAC first, Dataverse enrichment unless -SkipAgentDataverseQuery is set)." -ForegroundColor DarkCyan
}
else {
    Write-Warning "PAC CLI was not found in PATH. Agent inventory will be skipped."
}

# Iterate environments and collect Flow/App/Agent artifacts.
foreach ($environment in $environments) {
    $processedEnvironmentCount++
    $environmentId = [string](Get-ValueByPath -InputObject $environment -Paths @("EnvironmentName", "Name", "Id"))
    $environmentDisplayName = [string](Get-ValueByPath -InputObject $environment -Paths @("DisplayName", "EnvironmentDisplayName"))
    $environmentInstanceUrl = [string](Get-ValueByPath -InputObject $environment -Paths @("Internal.properties.linkedEnvironmentMetadata.instanceUrl"))

    if ([string]::IsNullOrWhiteSpace($environmentDisplayName)) {
        $environmentDisplayName = $environmentId
    }

    $percentComplete = [int](($processedEnvironmentCount / $totalEnvironmentCount) * 100)
    $environmentStatus = "Currently processing environment $processedEnvironmentCount of ${totalEnvironmentCount}: $environmentDisplayName"

    Write-Progress -Activity "Collecting Power Platform inventory" -Status $environmentStatus -PercentComplete $percentComplete
    Write-Host $environmentStatus -ForegroundColor Gray

    try {
        $flows = Get-AdminFlow -EnvironmentName $environmentId

        foreach ($flow in $flows) {
            $ownerEmail = [string](Resolve-OwnerEmail -Item $flow)
            $ownerDisplay = [string](Resolve-OwnerDisplayName -Item $flow)
            $ownerObjectId = [string](Get-OwnerObjectIdFromItem -Item $flow)

            if ([string]::IsNullOrWhiteSpace($ownerEmail) -and -not [string]::IsNullOrWhiteSpace($ownerObjectId)) {
                $ownerEmail = [string](Resolve-EntraEmailFromObjectId -ObjectId $ownerObjectId)

                if ([string]::IsNullOrWhiteSpace($ownerEmail) -and -not [string]::IsNullOrWhiteSpace($environmentInstanceUrl)) {
                    $ownerEmail = [string](Resolve-DataverseSystemUserEmail -InstanceUrl $environmentInstanceUrl -SystemUserId $ownerObjectId)
                }
            }

            $ownerValue = Resolve-OwnerField -OwnerDisplayName $ownerDisplay -OwnerEmail $ownerEmail
            if ([string]::IsNullOrWhiteSpace($ownerValue) -and -not [string]::IsNullOrWhiteSpace($ownerObjectId)) {
                $ownerValue = $ownerObjectId
            }

            $guid = [string](Get-ValueByPath -InputObject $flow -Paths @("FlowName", "Name", "Id", "properties.workflowEntityId"))
            $artifactName = [string](Get-ValueByPath -InputObject $flow -Paths @("DisplayName", "FlowDisplayName", "properties.displayName", "Name"))
            $created = Convert-ToIsoDate (Get-ValueByPath -InputObject $flow -Paths @("CreatedTime", "CreatedDateTime", "properties.createdTime", "properties.createdDateTime"))
            $connectionReferences = $null
            $connectionReferenceCount = 0

            if ($IncludeConnectionReferences) {
                $flowRefSummary = Get-FlowConnectionReferences -EnvironmentId $environmentId -FlowId $guid -FlowItem $flow
                $connectionReferences = [string]$flowRefSummary.List
                $connectionReferenceCount = [int]$flowRefSummary.Count
            }

            Add-InventoryRecord -Buffer $inventory -ArtifactType "Flow" -Environment $environmentDisplayName -EnvironmentId $environmentId -Guid $guid -ArtifactName $artifactName -Owner $ownerValue -OwnerEmail $ownerEmail -OwnerObjectId $ownerObjectId -DateCreated $created -ConnectionReferences $connectionReferences -ConnectionReferenceCount $connectionReferenceCount
            $flowCount++
        }
    }
    catch {
        Write-Warning "Failed to collect flows for environment '$environmentDisplayName'. $($_.Exception.Message)"
        Add-EnvironmentAccessIssue -Environment $environmentDisplayName -Stage "Flows" -Message $_.Exception.Message
    }

    try {
        $apps = Get-AdminPowerApp -EnvironmentName $environmentId

        foreach ($app in $apps) {
            $ownerEmail = [string](Resolve-OwnerEmail -Item $app)
            $ownerDisplay = [string](Resolve-OwnerDisplayName -Item $app)
            $ownerObjectId = [string](Get-OwnerObjectIdFromItem -Item $app)

            if ([string]::IsNullOrWhiteSpace($ownerEmail) -and -not [string]::IsNullOrWhiteSpace($ownerObjectId)) {
                $ownerEmail = [string](Resolve-EntraEmailFromObjectId -ObjectId $ownerObjectId)

                if ([string]::IsNullOrWhiteSpace($ownerEmail) -and -not [string]::IsNullOrWhiteSpace($environmentInstanceUrl)) {
                    $ownerEmail = [string](Resolve-DataverseSystemUserEmail -InstanceUrl $environmentInstanceUrl -SystemUserId $ownerObjectId)
                }
            }

            $ownerValue = Resolve-OwnerField -OwnerDisplayName $ownerDisplay -OwnerEmail $ownerEmail
            if ([string]::IsNullOrWhiteSpace($ownerValue) -and -not [string]::IsNullOrWhiteSpace($ownerObjectId)) {
                $ownerValue = $ownerObjectId
            }

            $guid = [string](Get-ValueByPath -InputObject $app -Paths @("AppName", "Name", "Id"))
            $artifactName = [string](Get-ValueByPath -InputObject $app -Paths @("DisplayName", "AppDisplayName", "properties.displayName", "Name"))
            $created = Convert-ToIsoDate (Get-ValueByPath -InputObject $app -Paths @("CreatedTime", "CreatedDateTime", "properties.createdTime", "properties.createdDateTime"))
            $connectionReferences = $null
            $connectionReferenceCount = 0

            if ($IncludeConnectionReferences) {
                $appRefSummary = Get-AppConnectionReferences -EnvironmentId $environmentId -AppId $guid -AppItem $app
                $connectionReferences = [string]$appRefSummary.List
                $connectionReferenceCount = [int]$appRefSummary.Count
            }

            Add-InventoryRecord -Buffer $inventory -ArtifactType "App" -Environment $environmentDisplayName -EnvironmentId $environmentId -Guid $guid -ArtifactName $artifactName -Owner $ownerValue -OwnerEmail $ownerEmail -OwnerObjectId $ownerObjectId -DateCreated $created -ConnectionReferences $connectionReferences -ConnectionReferenceCount $connectionReferenceCount
            $appCount++
        }
    }
    catch {
        Write-Warning "Failed to collect apps for environment '$environmentDisplayName'. $($_.Exception.Message)"
        Add-EnvironmentAccessIssue -Environment $environmentDisplayName -Stage "Apps" -Message $_.Exception.Message
    }

    if ($usePacFallbackForAgents) {
        try {
            $instanceUrl = [string](Get-ValueByPath -InputObject $environment -Paths @("Internal.properties.linkedEnvironmentMetadata.instanceUrl"))

                if ($SkipAgentDataverseQuery) {
                    Write-Host "Strict PAC mode for environment '$environmentDisplayName': skipping all Dataverse-based agent queries." -ForegroundColor DarkYellow

                    $pacAgents = @(Get-PacCopilotAgentsForEnvironment -EnvironmentId $environmentId -InstanceUrl $instanceUrl -EnvironmentDisplayName $environmentDisplayName)

                    foreach ($pacAgent in $pacAgents) {
                        $agentGuid = [string]$pacAgent.CopilotId
                        $agentName = [string]$pacAgent.Name

                        # Strict mode: do not call Dataverse for tools/owner/date enrichment.
                        $connectionReferences = $null
                        $connectionReferenceCount = 0

                        Add-InventoryRecord -Buffer $inventory -ArtifactType "Agent" -Environment $environmentDisplayName -EnvironmentId $environmentId -Guid $agentGuid -ArtifactName $agentName -Owner $null -OwnerEmail $null -OwnerObjectId $null -DateCreated $null -ConnectionReferences $connectionReferences -ConnectionReferenceCount $connectionReferenceCount
                        $agentCount++
                    }

                    if (@($pacAgents).Count -eq 0 -and -not $pacFallbackOwnerDateMissingWarningShown) {
                        Write-Warning "PAC fallback returned zero agents for environment '$environmentDisplayName'."
                        $pacFallbackOwnerDateMissingWarningShown = $true
                    }

                    continue
                }

            $pacAgents = @(Get-PacCopilotAgentsForEnvironment -EnvironmentId $environmentId -InstanceUrl $instanceUrl -EnvironmentDisplayName $environmentDisplayName)

            $botOwnerMap = @{}
            $dataverseBotsById = @{}
            $dataverseQuerySucceeded = $false

            if ([string]::IsNullOrWhiteSpace($instanceUrl)) {
                Write-Warning "Dataverse instance URL not available for environment '$environmentDisplayName'. PAC agents will be returned without Dataverse enrichment."
            }
            else {
                try {
                    $botOwnerMap = Get-DataverseBotOwnerMap -InstanceUrl $instanceUrl -EnvironmentDisplayName $environmentDisplayName
                    $dataverseBots = @(Get-DataverseBots -InstanceUrl $instanceUrl -EnvironmentDisplayName $environmentDisplayName)
                    $dataverseQuerySucceeded = $true

                    foreach ($dataverseBot in $dataverseBots) {
                        $dataverseBotId = [string]$dataverseBot.BotId
                        if ([string]::IsNullOrWhiteSpace($dataverseBotId)) {
                            continue
                        }

                        $dataverseBotsById[$dataverseBotId.ToLowerInvariant()] = $dataverseBot
                    }
                }
                catch {
                    Write-Warning "Dataverse enrichment failed for environment '$environmentDisplayName'. Continuing with PAC-only agent details. $($_.Exception.Message)"
                }
            }

            foreach ($pacAgent in $pacAgents) {
                $agentGuid = [string]$pacAgent.CopilotId
                $agentName = [string]$pacAgent.Name
                $agentMapKey = if ([string]::IsNullOrWhiteSpace($agentGuid)) { $null } else { $agentGuid.ToLowerInvariant() }

                $ownerObjectId = $null
                $ownerDisplayName = $null
                $created = $null

                if (-not [string]::IsNullOrWhiteSpace($agentMapKey) -and $dataverseBotsById.ContainsKey($agentMapKey)) {
                    $matchedDataverseBot = $dataverseBotsById[$agentMapKey]
                    $ownerObjectId = [string]$matchedDataverseBot.OwnerSystemUserId
                    $ownerDisplayName = [string]$matchedDataverseBot.OwnerDisplayName
                    $created = Convert-ToIsoDate $matchedDataverseBot.CreatedOn
                }
                elseif (-not [string]::IsNullOrWhiteSpace($agentMapKey) -and $botOwnerMap.ContainsKey($agentMapKey)) {
                    $ownerObjectId = [string]$botOwnerMap[$agentMapKey].OwnerSystemUserId
                    $ownerDisplayName = [string]$botOwnerMap[$agentMapKey].OwnerDisplayName
                    $created = Convert-ToIsoDate $botOwnerMap[$agentMapKey].CreatedOn
                }

                $ownerEmail = $null
                if (-not [string]::IsNullOrWhiteSpace($ownerObjectId) -and -not [string]::IsNullOrWhiteSpace($instanceUrl)) {
                    $ownerEmail = Resolve-DataverseSystemUserEmail -InstanceUrl $instanceUrl -SystemUserId $ownerObjectId
                }

                $ownerValue = Resolve-OwnerField -OwnerDisplayName $ownerDisplayName -OwnerEmail $ownerEmail
                if ([string]::IsNullOrWhiteSpace($ownerValue) -and -not [string]::IsNullOrWhiteSpace($ownerObjectId)) {
                    $ownerValue = $ownerObjectId
                }

                $connectionReferences = $null
                $connectionReferenceCount = 0
                if ($IncludeConnectionReferences -and -not [string]::IsNullOrWhiteSpace($instanceUrl) -and -not [string]::IsNullOrWhiteSpace($agentGuid)) {
                    $agentToolSummary = Get-AgentToolSummary -EnvironmentUrl $instanceUrl -AgentId $agentGuid
                    $connectionReferences = [string]$agentToolSummary.List
                    $connectionReferenceCount = [int]$agentToolSummary.Count
                }

                Add-InventoryRecord -Buffer $inventory -ArtifactType "Agent" -Environment $environmentDisplayName -EnvironmentId $environmentId -Guid $agentGuid -ArtifactName $agentName -Owner $ownerValue -OwnerEmail $ownerEmail -OwnerObjectId $ownerObjectId -DateCreated $created -ConnectionReferences $connectionReferences -ConnectionReferenceCount $connectionReferenceCount
                $agentCount++
            }

            if (@($pacAgents).Count -eq 0 -and $dataverseQuerySucceeded) {
                Write-Warning "PAC returned zero agents for environment '$environmentDisplayName'. Falling back to Dataverse-only agent inventory."

                $dataverseBots = @($dataverseBotsById.Values)
                foreach ($agent in $dataverseBots) {
                    $agentGuid = [string]$agent.BotId
                    $ownerEmail = $null
                    $ownerDisplayName = [string]$agent.OwnerDisplayName
                    $ownerObjectId = [string]$agent.OwnerSystemUserId
                    $created = Convert-ToIsoDate $agent.CreatedOn

                    if (-not [string]::IsNullOrWhiteSpace($ownerObjectId)) {
                        $ownerEmail = Resolve-DataverseSystemUserEmail -InstanceUrl $instanceUrl -SystemUserId $ownerObjectId
                    }

                    $ownerValue = Resolve-OwnerField -OwnerDisplayName $ownerDisplayName -OwnerEmail $ownerEmail
                    if ([string]::IsNullOrWhiteSpace($ownerValue) -and -not [string]::IsNullOrWhiteSpace($ownerObjectId)) {
                        $ownerValue = $ownerObjectId
                    }

                    $connectionReferences = $null
                    $connectionReferenceCount = 0

                    if ($IncludeConnectionReferences) {
                        $agentToolSummary = Get-AgentToolSummary -EnvironmentUrl $instanceUrl -AgentId $agentGuid
                        $connectionReferences = [string]$agentToolSummary.List
                        $connectionReferenceCount = [int]$agentToolSummary.Count
                    }

                    Add-InventoryRecord -Buffer $inventory -ArtifactType "Agent" -Environment $environmentDisplayName -EnvironmentId $environmentId -Guid $agentGuid -ArtifactName ([string]$agent.Name) -Owner $ownerValue -OwnerEmail $ownerEmail -OwnerObjectId $ownerObjectId -DateCreated $created -ConnectionReferences $connectionReferences -ConnectionReferenceCount $connectionReferenceCount
                    $agentCount++
                }
            }

            if (@($pacAgents).Count -eq 0 -and -not $pacFallbackOwnerDateMissingWarningShown) {
                $zeroAgentWarningMessage = $null

                if ($dataverseQuerySucceeded -and @($dataverseBotsById.Values).Count -eq 0) {
                    $zeroAgentWarningMessage = "PAC and Dataverse agent discovery both returned zero agents for environment '$environmentDisplayName'."
                }
                elseif (-not $dataverseQuerySucceeded) {
                    $zeroAgentWarningMessage = "PAC returned zero agents for environment '$environmentDisplayName', and Dataverse enrichment was unavailable or failed."
                }

                if (-not [string]::IsNullOrWhiteSpace($zeroAgentWarningMessage)) {
                    Write-Warning $zeroAgentWarningMessage
                    $pacFallbackOwnerDateMissingWarningShown = $true
                }
            }

        }
        catch {
            Write-Warning "Failed to collect agents through PAC/Dataverse agent path for environment '$environmentDisplayName'. $($_.Exception.Message)"
            Add-EnvironmentAccessIssue -Environment $environmentDisplayName -Stage "Agents" -Message $_.Exception.Message
        }
    }
}

Write-Progress -Activity "Collecting Power Platform inventory" -Completed

# Ensure output directory exists, then export a stable, sorted CSV.
$outputDirectory = Split-Path -Path $OutputCsvPath -Parent
if (-not [string]::IsNullOrWhiteSpace($outputDirectory) -and -not (Test-Path -Path $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
}

$sortedInventory = $inventory | Sort-Object -Property ArtifactType, Environment, GUID
$maxConnectionReferenceCount = @($sortedInventory | ForEach-Object { [int]$_.ConnectionReferenceCount } | Measure-Object -Maximum).Maximum

$exportRows = foreach ($item in $sortedInventory) {
    $exportRecord = [ordered]@{
        ArtifactType = $item.ArtifactType
        Environment = $item.Environment
        EnvironmentId = $item.EnvironmentId
        GUID = $item.GUID
        ArtifactName = $item.ArtifactName
        Owner = $item.Owner
        OwnerEmail = $item.OwnerEmail
        OwnerObjectId = $item.OwnerObjectId
        DateCreated = $item.DateCreated
        ConnectionReferenceCount = $item.ConnectionReferenceCount
    }

    $referenceValues = @()
    if (-not [string]::IsNullOrWhiteSpace([string]$item.ConnectionReferences)) {
        $referenceValues = @([string]$item.ConnectionReferences -split '\s*;\s*' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }

    for ($index = 1; $index -le $maxConnectionReferenceCount; $index++) {
        $columnName = "ConnectionReference{0}" -f $index
        $exportRecord[$columnName] = if ($index -le $referenceValues.Count) { $referenceValues[$index - 1] } else { $null }
    }

    [PSCustomObject]$exportRecord
}

$exportRows | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8

Write-Host "Inventory complete." -ForegroundColor Green
Write-Host "Flows:  $flowCount" -ForegroundColor Green
Write-Host "Apps:   $appCount" -ForegroundColor Green
Write-Host "Agents: $agentCount" -ForegroundColor Green
Write-Host "Total:  $($inventory.Count)" -ForegroundColor Green
Write-Host "CSV:    $OutputCsvPath" -ForegroundColor Green

if ($script:environmentAccessIssues.Count -gt 0) {
    Write-Host "" 
    Write-Host "Environment access issues:" -ForegroundColor Yellow
    foreach ($issue in ($script:environmentAccessIssues | Sort-Object -Property Environment, Stage, Message)) {
        Write-Host ("- {0} [{1}]: {2}" -f $issue.Environment, $issue.Stage, $issue.Message) -ForegroundColor Yellow
    }
}