##############################################
# Description: Azure Automation version of GetSitePermissionDataSet.ps1
#              Gets site access data from SharePoint and group permissions.
#              Creates a CSV for each owner/member/visitor of a SPO or M365 Group
#              and uploads to SharePoint via Graph API.
#
#              Designed for Azure Automation Account with Managed Identity.
#
# Modules Required (pin versions in Automation Account):
#   - PnP.PowerShell (3.1.0)
#   - Microsoft.Graph.Authentication (1.28.0)
#   - Microsoft.Graph.Sites (1.28.0)
#   - Microsoft.Graph.Groups (1.28.0)
#   - Microsoft.Graph.Reports (1.28.0)
#
# Alex Grover - alexgrover@microsoft.com
#
##############################################
# Parameters
##############################################

param (
    [string]$DriveId = "",                  # SharePoint Drive ID for output upload
    [string]$OutputFolderPath = "",         # Folder path in the drive (e.g. "Reports/Permissions")
    [string]$AdminSiteUrl = "",             # SPO Admin URL (e.g. https://contoso-admin.sharepoint.com)

    # PnP Auth - Managed Identity doesn't work with PnP, so we need app reg for PnP
    [string]$ClientId = "",
    [string]$TenantId = "",
    [string]$CertificateThumbprint = "",

    # Site filter
    [bool]$AllSites = $true,
    [string[]]$SiteList = @(),              # Specific sites to process (ignored if AllSites = $true)

    # Config
    [bool]$GetMembers = $false,
    [bool]$GetVisitors = $false,

    # Groups to skip
    [string[]]$GroupsToSkip = @()
)

##############################################
# Dependencies
##############################################

foreach ($moduleName in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Sites', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Reports')) {
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Write-Error "Module '$moduleName' not found. Please add it to the Automation Account."
        exit 1
    }
    Write-Output "Importing module: $moduleName..."
    Import-Module -Name $moduleName -Force
}

# PnP must be imported after Graph modules to avoid assembly conflicts
if (-not (Get-Module -ListAvailable -Name 'PnP.PowerShell')) {
    Write-Error "Module 'PnP.PowerShell' not found. Please add it to the Automation Account."
    exit 1
}
Write-Output "Importing module: PnP.PowerShell..."
Import-Module -Name PnP.PowerShell -Force

##############################################
# Variables
##############################################

$timeStamp = Get-Date -Format "yyyyMMddHHmmss"
$outputFileName = "SitePermissionRecord-$timeStamp.csv"
$script:csvBuffer = New-Object System.IO.MemoryStream

enum MemberTypes {
    Owner
    Member
    Visitor
    Admin
}

##############################################
# Functions
##############################################

function ConnectToMSGraph {
    try {
        Connect-MgGraph -Identity -NoWelcome
        Write-Output "Connected to Microsoft Graph via Managed Identity."
    }
    catch {
        Write-Error "Error connecting to MS Graph - $($_.Exception.Message)"
        exit 1
    }
}

function ConnectToPnP ($siteUrl) {
    try {
        Connect-PnPOnline -Url $siteUrl -ClientId $ClientId -Tenant $TenantId -Thumbprint $CertificateThumbprint
    }
    catch {
        Write-Output "ERROR: Connecting to PnP ($siteUrl) - $($_.Exception.Message)"
    }
}

function Get-Sites {
    try {
        if (!$AllSites) {
            if ($SiteList.Count -eq 0) {
                Write-Error "SiteList parameter is empty but AllSites is false."
                exit 1
            }
            $sites = Get-MgSite -Property "siteCollection,webUrl,id" -All | Where-Object { !($_.WebUrl.Contains("my.sharepoint.com")) } | Where-Object { $SiteList -contains $_.WebUrl } -ErrorAction Stop
            return $sites
        }

        $sites = Get-MgSite -Property "siteCollection,webUrl,id" -All | Where-Object { !($_.WebUrl.Contains("my.sharepoint.com")) } -ErrorAction Stop
        return $sites
    }
    catch {
        Write-Error "Error getting sites - $($_.Exception.Message)"
        exit 1
    }
}

function Write-CsvLine {
    param([PSObject]$LogEntry)
    
    # Write header if first line
    if ($script:csvBuffer.Length -eq 0) {
        $headers = $LogEntry.PSObject.Properties.Name
        $headerLine = ($headers | ForEach-Object { '"{0}"' -f $_.Replace('"', '""') }) -join ','
        $bytes = [System.Text.Encoding]::UTF8.GetBytes("$headerLine`n")
        $script:csvBuffer.Write($bytes, 0, $bytes.Length)
    }

    # Write data row
    $values = @()
    foreach ($prop in $LogEntry.PSObject.Properties) {
        $value = $prop.Value
        if ($null -eq $value) { $values += '""' }
        elseif ($value -is [string]) { $values += '"{0}"' -f $value.Replace('"', '""') }
        elseif ($value -is [bool]) { $values += $value.ToString() }
        elseif ($value -is [DateTime]) { $values += '"{0:O}"' -f $value }
        else { $values += '"{0}"' -f ($value | ConvertTo-Json -Compress -Depth 5).Replace('"', '""') }
    }
    $csvLine = $values -join ','
    $bytes = [System.Text.Encoding]::UTF8.GetBytes("$csvLine`n")
    $script:csvBuffer.Write($bytes, 0, $bytes.Length)
}

function Write-LogEntry($siteUrl, $siteName, $siteUsage, $user, [MemberTypes]$type, $message, $sharingCapability, $domainList, $lockStatus) {
    $rootWebTemplate = $siteUsage.'Root Web Template'
    $siteTemplate = $rootWebTemplate

    if ($message -ne "Unable to get Site details") {
        if ($rootWebTemplate -eq "Group") { $siteTemplate = "Team site" }
        if ($rootWebTemplate -eq "Site Page Publishing") { $siteTemplate = "Communication site" }
        if ($rootWebTemplate -eq "Team Site") {
            ConnectToPnP $AdminSiteUrl
            $site = Get-PnPTenantSite -Identity $siteUrl
            if ($site.Template -eq "STS#0") { $siteTemplate = "Team site (classic experience)" }
            if ($site.Template -eq "STS#3") { $siteTemplate = "Team site (no Microsoft 365 group)" }
        }
    }

    $logLine = New-Object -TypeName PSObject -Property @{
        UserType = $type
        LogTime = Get-Date
        SiteUrl = $siteUrl
        SiteName = $siteName
        LastContentModifiedDate = $siteUsage.'Last Activity Date'
        Notes = $message
        Email = $user.Item1
        SiteTemplate = $siteTemplate
        SharingCapability = $sharingCapability
        DomainList = $domainList
        FileCount = $siteUsage.'File Count'
        LockStatus = $lockStatus
        PermissionSource = $user.Item2
        GroupId = $user.Item3
        GroupName = $user.Item4
    }

    Write-CsvLine -LogEntry $logLine
}

function GetSiteUsers([MemberTypes]$type, $graphObj, $pnpObj, $group) {
    $userEmails = @()
    ConnectToPnP -siteUrl $graphObj.WebUrl
    $domainObjs = @()

    try {
        if ($type -eq [MemberTypes]::Owner) {
            Write-Output "    Getting Site Owners"
            $domainObjs = Get-PnPGroup -AssociatedOwnerGroup | Get-PnPGroupMember
        }
        if ($type -eq [MemberTypes]::Member) {
            Write-Output "    Getting Site Members"
            $domainObjs = Get-PnPGroup -AssociatedMemberGroup | Get-PnPGroupMember
        }
        if ($type -eq [MemberTypes]::Visitor) {
            Write-Output "    Getting Site Visitors"
            $domainObjs = Get-PnPGroup -AssociatedVisitorGroup | Get-PnPGroupMember
        }
    }
    catch {
        Write-Output "    WARNING: Error getting site $type groups - likely bespoke permissions"
        return "ErrorGettingSPOGroups"
    }

    if ($type -eq [MemberTypes]::Admin) {
        Write-Output "    Getting Site Admins"
        $domainObjs = Get-PnPSiteCollectionAdmin
    }

    foreach ($domainObj in $domainObjs) {
        # User
        if ($domainObj.LoginName.Contains("|membership|")) {
            Write-Output "      Found: $($domainObj.Email) as $type"
            $userEmails += [Tuple]::Create($domainObj.Email, "User", "N/A")
            continue
        }

        # Role Manager
        if ($domainObj.LoginName.Contains("|rolemanager|")) {
            Write-Output "      Found: $($domainObj.Title) role as $type"
            $userEmails += [Tuple]::Create($domainObj.Title, "RoleManager", "N/A")
            continue
        }

        # M365 Group
        if ($domainObj.LoginName.Contains("|federateddirectoryclaimprovider|")) {
            Write-Output "      Found: $($domainObj.Email) (Group) as $type"
            $groupId = $domainObj.LoginName.Split("|")[2]

            if ($GroupsToSkip -contains $groupId.Substring(0, [Math]::Min($groupId.Length, 36))) {
                Write-Output "      Skipping: $($domainObj.Title) - in ignore list"
                $userEmails += [Tuple]::Create($domainObj.Title, "IgnoredGroup", "N/A")
                continue
            }

            $getOwners = $false
            if ($group -and ([MemberTypes]::Owner -eq $type -or [MemberTypes]::Admin -eq $type)) {
                $drive = Get-MgSiteDefaultDrive -SiteId $graphObj.Id
                if ($drive.Owner.AdditionalProperties.group.id -eq $groupId.Substring(0, [Math]::Min($groupId.Length, 36))) {
                    $getOwners = $true
                }
            }

            if ($getOwners) {
                Write-Output "      Getting Group Owners"
                $members = Get-MgGroupOwner -GroupId $groupId.Substring(0, [Math]::Min($groupId.Length, 36)) -Property "userPrincipalName" -All
            }
            else {
                Write-Output "      Getting Group Members"
                $members = Get-MgGroupMember -GroupId $groupId.Substring(0, [Math]::Min($groupId.Length, 36)) -Property "userPrincipalName" -All
            }

            foreach ($member in $members) {
                Write-Output "        Found: $($member.AdditionalProperties.userPrincipalName) as inferred $type"
                $userEmails += [Tuple]::Create($member.AdditionalProperties.userPrincipalName, "Group", $groupId.Substring(0, [Math]::Min($groupId.Length, 36)), $domainObj.Email)
            }
            continue
        }

        # Security Group
        if ($domainObj.LoginName.Contains("|tenant|")) {
            if ($domainObj.LoginName.Equals("c:0t.c|tenant|b71daa58-3cb3-4b97-a6e8-eae7f2a30f20") -or $domainObj.LoginName.Equals("c:0t.c|tenant|e8d578f9-c761-4097-b616-a1111909a468")) {
                continue
            }

            if ($domainObj.Title.Equals("Global Administrator")) {
                continue
            }

            Write-Output "      Found: $($domainObj.LoginName) (SecGroup) as $type"
            $groupId = $domainObj.LoginName.Split("|")[2]

            if ($GroupsToSkip -contains $groupId.Substring(0, [Math]::Min($groupId.Length, 36))) {
                Write-Output "      Skipping: $($domainObj.Title) - in ignore list"
                $userEmails += [Tuple]::Create($domainObj.Title, "IgnoredGroup", "N/A")
                continue
            }

            $groupEmail = $domainObj.Email
            if ($null -eq $groupEmail) { $groupEmail = $domainObj.Title }

            Write-Output "        Getting SecGroup Members"

            try {
                $members = Get-MgGroupMember -GroupId $groupId -Property "userPrincipalName,id,securityEnabled" -All -ErrorAction Stop
            }
            catch {
                Write-Output "        ERROR: getting members for $($domainObj.Title) - $($_.Exception.Message)"
                continue
            }

            foreach ($member in $members) {
                if ($null -ne $member.AdditionalProperties.userPrincipalName) {
                    Write-Output "        Found: $($member.AdditionalProperties.userPrincipalName) as inferred $type"
                    $userEmails += [Tuple]::Create($member.AdditionalProperties.userPrincipalName, "SecGroup", $groupId, $groupEmail)
                    continue
                }
                else {
                    # Skip DLs
                    if (!$member.AdditionalProperties.securityEnabled) {
                        continue
                    }

                    $users = @()
                    $userEmailsFromGroups = GetSecGroupMembers -groupId $member.Id -users $users -space "          "

                    foreach ($userFromGroup in $userEmailsFromGroups) {
                        $userEmails += [Tuple]::Create($userFromGroup, "SecGroup", $groupId, $groupEmail)
                    }
                }
            }
            continue
        }
    }

    return $userEmails | Get-Unique
}

function ProcessSite($site, $usage) {
    Write-Output "  Processing: $($site.WebUrl)"

    try {
        $siteObject = GetSite -siteUrl $site.WebUrl
    }
    catch {
        Write-Output "  WARNING: Unable to get Site details"
        Write-LogEntry -siteUrl $site.WebUrl -siteName $null -lockStatus $null -sharingCapability $null -domainList $null -user $null -siteUsage $usage -message "Unable to get Site details"
        return
    }

    $lockStatus = $siteObject.LockState
    $sharingCapability = $siteObject.SharingCapability
    $domainList = $siteObject.SharingAllowedDomainList
    $siteName = $siteObject.Title

    if ($lockStatus -eq "NoAccess") {
        Write-Output "  WARNING: Site is locked with no access"
        Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $null -siteUsage $usage -message "Site is locked with no access"
        return
    }

    $group = ($usage.'Root Web Template' -eq "Group")

    # Owners
    try {
        $owners = GetSiteUsers -type ([MemberTypes]::Owner) -graphObj $site -pnpObj $siteObject -group $group
        if ($owners -eq "ErrorGettingSPOGroups") {
            Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $null -siteUsage $usage -message "Site has bespoke owner permissions"
        }
        else {
            foreach ($owner in $owners) {
                Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $owner -siteUsage $usage -type ([MemberTypes]::Owner)
            }
        }
    }
    catch {
        Write-Output "  ERROR: getting owners - $($_.Exception.Message)"
        Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $null -siteUsage $usage -message "Error getting owners - $($_.Exception.Message)"
    }

    # Members
    if ($GetMembers) {
        try {
            $members = GetSiteUsers -type ([MemberTypes]::Member) -graphObj $site -pnpObj $siteObject -group $group
            if ($members -eq "ErrorGettingSPOGroups") {
                Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $null -siteUsage $usage -message "Site has bespoke member permissions"
            }
            else {
                foreach ($member in $members) {
                    Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $member -siteUsage $usage -type ([MemberTypes]::Member)
                }
            }
        }
        catch {
            Write-Output "  ERROR: getting members - $($_.Exception.Message)"
            Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $null -siteUsage $usage -message "Error getting members - $($_.Exception.Message)"
        }
    }

    # Visitors
    if ($GetVisitors) {
        try {
            $visitors = GetSiteUsers -type ([MemberTypes]::Visitor) -graphObj $site -pnpObj $siteObject -group $group
            if ($visitors -eq "ErrorGettingSPOGroups") {
                Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $null -siteUsage $usage -message "Site has bespoke visitor permissions"
            }
            else {
                foreach ($visitor in $visitors) {
                    Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $visitor -siteUsage $usage -type ([MemberTypes]::Visitor)
                }
            }
        }
        catch {
            Write-Output "  ERROR: getting visitors - $($_.Exception.Message)"
            Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $null -siteUsage $usage -message "Error getting visitors - $($_.Exception.Message)"
        }
    }

    # Admins
    try {
        $admins = GetSiteUsers -type ([MemberTypes]::Admin) -graphObj $site -pnpObj $siteObject -group $group
        foreach ($admin in $admins) {
            Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $admin -siteUsage $usage -type ([MemberTypes]::Admin)
        }
    }
    catch {
        Write-Output "  ERROR: getting admins - $($_.Exception.Message)"
        Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $null -siteUsage $usage -message "Error getting admins - $($_.Exception.Message)"
    }
}

## Recursive function to get members (users) of a security group
function GetSecGroupMembers($groupId, $users, $space) {
    Write-Output "$($space)Getting Nested Group Members"
    
    $members = Get-MgGroupMember -GroupId $groupId -Property "userPrincipalName,id" -All

    foreach ($member in $members) {
        if ($null -ne $member.AdditionalProperties.userPrincipalName) {
            Write-Output "$($space)  Found: $($member.AdditionalProperties.userPrincipalName)"
            $users += $member.AdditionalProperties.userPrincipalName
        }
        else {
            $users += GetSecGroupMembers -groupId $member.Id -users $users -space "$($space)  "
        }
    }
    return $users
}

function GetSiteUsageReport($sites) {
    try {
        $tempPath = [System.IO.Path]::GetTempPath()
        $reportFile = Join-Path $tempPath "SPOUsageReport.csv"
        Get-MgReportSharePointSiteUsageDetail -Period D7 -OutFile $reportFile
        $siteUsage = Import-Csv $reportFile
        Remove-Item $reportFile -Force
        return $siteUsage | Where-Object { $sites -contains $_.'Site URL' }
    }
    catch {
        Write-Error "Error getting site usage report - $($_.Exception.Message)"
        exit 1
    }
}

function GetSite($siteUrl) {
    ConnectToPnP -siteUrl $AdminSiteUrl
    $site = Get-PnPTenantSite -Identity $siteUrl
    return $site
}

function Upload-CsvToSharePoint {
    param (
        [string]$DriveId,
        [string]$FileName,
        [System.IO.MemoryStream]$Content
    )

    $filePath = if ([string]::IsNullOrWhiteSpace($OutputFolderPath)) {
        "/$FileName"
    } else {
        "/$OutputFolderPath/$FileName"
    }

    $fileSize = $Content.Length
    Write-Output "Uploading CSV to SharePoint ($([Math]::Round($fileSize / 1KB, 2)) KB)..."

    # For files under 4MB, use simple upload
    if ($fileSize -lt 4MB) {
        try {
            $Content.Position = 0
            $bytes = $Content.ToArray()
            Invoke-MgGraphRequest -Method PUT -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId/root:$($filePath):/content" -Body $bytes -ContentType "application/octet-stream" | Out-Null
            Write-Output "Upload complete: $filePath"
        }
        catch {
            Write-Error "Failed to upload file: $($_.Exception.Message)"
            throw
        }
    }
    else {
        # Chunked upload for larger files
        $UploadMultipleSize = 327680  # 320 KiB
        $chunkSize = [Math]::Ceiling((4MB) / $UploadMultipleSize) * $UploadMultipleSize

        try {
            $bodyJson = '{ "item": { "@microsoft.graph.conflictBehavior": "replace" } }'
            $session = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId/root:$($filePath):/createUploadSession" -Body $bodyJson -ContentType "application/json"
            
            if (-not $session.uploadUrl) {
                throw "Failed to create upload session: no uploadUrl returned"
            }

            $uploadUrl = $session.uploadUrl
            $Content.Position = 0
            $position = 0

            while ($position -lt $fileSize) {
                $bytesToRead = [Math]::Min($chunkSize, $fileSize - $position)
                $chunk = New-Object byte[] $bytesToRead
                $Content.Read($chunk, 0, $bytesToRead) | Out-Null

                $end = $position + $bytesToRead - 1
                $range = "bytes $position-$end/$fileSize"

                Invoke-MgGraphRequest -Method PUT -Uri $uploadUrl -Headers @{ "Content-Range" = $range } -Body $chunk -SkipHeaderValidation | Out-Null
                
                $position += $bytesToRead
                Write-Output "  Uploaded: $([Math]::Round($position / 1KB, 2)) KB / $([Math]::Round($fileSize / 1KB, 2)) KB"
            }

            Write-Output "Upload complete: $filePath"
        }
        catch {
            Write-Error "Failed during chunked upload: $($_.Exception.Message)"
            throw
        }
    }
}

##############################################
# Main
##############################################

Write-Output "=== Site Permission Record - Automation ==="
Write-Output "Start time: $(Get-Date)"

## Connect to Graph (Managed Identity)
ConnectToMSGraph

## Get all sites
$sites = Get-Sites
Write-Output "Found $($sites.Count) sites to process."

## Get the site usage report
$siteUsage = GetSiteUsageReport $sites.WebUrl
Write-Output "Site usage report retrieved."

## Process each site
$currentItem = 0
foreach ($site in $sites) {
    $currentItem++
    Write-Output "[$currentItem/$($sites.Count)] $($site.WebUrl)"

    $siteUsageEntry = $siteUsage | Where-Object { $_.'Site URL' -eq $site.WebUrl }
    ProcessSite -site $site -usage $siteUsageEntry
}

## Upload CSV to SharePoint
if ($script:csvBuffer.Length -gt 0) {
    Upload-CsvToSharePoint -DriveId $DriveId -FileName $outputFileName -Content $script:csvBuffer
}
else {
    Write-Output "WARNING: No data collected. CSV not uploaded."
}

## Cleanup
$script:csvBuffer.Dispose()

Write-Output "=== Complete ==="
Write-Output "End time: $(Get-Date)"
