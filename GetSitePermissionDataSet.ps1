#############################################
# Description: This script will get site acess data from the underlying SharePoint
#              and group permissions.
#              The script will create a CSV line for each owner/member/visitors of 
#              a SPO or M365 Group detaling the site access and permissions.
#              The first version will not look at broken inheritance.
#              
#
# Todo :
#           Bespoke permissions
#           Broken inheritance (Sub sites, Libraries, Items?)
#           Sharing links
#
# Alex Grover - alexgrover@microsoft.com
#
#
##############################################
# Dependencies
##############################################
## Requires the following modules:
try {
    Import-Module Microsoft.Graph.Sites
    Import-Module Microsoft.Graph.Groups
    Import-Module PnP.PowerShell
}
catch {
    Write-Error "Error importing modules required modules - $($Error[0].Exception.Message))"
    Exit
}

# Graph Permissions
# Sites.Read.All
# GroupMember.Read.All
# User.Read.All
# Reports.Read.All

# SPO Permissions
# Sites.Read.All
# Site.FullControll.All


##############################################
# Variables
##############################################



$clientId = "6c0f4f31-bf65-4b74-8b1c-9c038ea5c102"
$tenantId = "M365CPI77517573.onmicrosoft.com"
$adminSiteUrl = "https://M365CPI77517573-admin.sharepoint.com"


$thumbprint = "6ADC063641A24BB0BD68786AB71F07315CED9076"




# Process all sites or only sites in the input file
$allSites = $true

# List of Sites to check (ignore if $allSites = $true)
$inputSitesCSV = "./SiteCollectionsList.txt"

# Log file location (timestamped with script start time)
$timeStamp = Get-Date -Format "yyyyMMddHHmmss"
$logFileLocation = "Output\SitePermissionRecord-$timeStamp.csv"

# Groups to skip (that you do not want to expand membership of)
$groupsToSkip = @(
    "EXT-Guest-Users",
    "FirstSource-Guest-Users",
    "Webhelp-Guest-Users",
    "60k-Guest-Users",
    "6c7ef884-1924-4fd8-96f4-c9641fc4187c"
)

# Config (Do we want to get visistors or members? Owner and Admins are always returned)
$getMembers = $false
$getVisitors = $false

enum MemberTypes {
    Owner
    Member
    Visitor
    Admin
}

##############################################
# Functions
##############################################

function ConnectToMSGraph 
{  
    try{
        Connect-MgGraph -ClientId $clientId -TenantId $tenantId -CertificateThumbprint $thumbprint
    }
    catch{
        Write-Host "Error connecting to MS Graph - $($Error[0].Exception.Message)" -ForegroundColor Red
        Exit
    }
}

function ConnectToPnP ($siteUrl){
    try{
        Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant $tenantId -Thumbprint $thumbprint
    }
    catch{
        Write-Host "Error connecting to PnP - $($Error[0].Exception.Message)" -ForegroundColor Red
    }
}

function ReadSitesFromTxtFile($siteListCSVFile) {
    try {
        $siteList = Get-Content $siteListCSVFile
        return $siteList
    }
    catch {
        Write-Host "Error reading site list from file: $siteListCSVFile" -ForegroundColor Red
        Write-Host $_.Exception.Message
        Write-Host "Exiting..."
        exit
    }
}

function Get-Sites
{
    try {
        
        if (!$allSites) {
            $siteList = ReadSitesFromTxtFile($inputSitesCSV)
            $sites = Get-MgSite -Property "siteCollection,webUrl,id" -All | Where-Object { !($_.WebUrl.Contains("my.sharepoint.com"))} | where { $siteList -contains $_.WebUrl } -ErrorAction Stop
            return $sites 
        }

        # Get all sites, filter out OneDrive sites
        $sites = Get-MgSite -Property "siteCollection,webUrl,id" -All | Where-Object { !($_.WebUrl.Contains("my.sharepoint.com"))} -ErrorAction Stop
        return $sites #| where {$_.WebUrl.Contains("/home")}
    }
    catch {
        Write-Host " Error getting sites" -ForegroundColor Red
    }   
}

function Write-LogEntry($siteUrl, $siteName, $siteUsage, $user, [MemberTypes]$type, $message, $sharingCapability, $domainList, $lockStatus)
{
    $rootWebTemplate = $siteUsage.'Root Web Template'
    $siteTemplate = $rootWebTemplate

    if ($message -ne "Unable to get Site details") 
    { 
        if ($rootWebTemplate -eq "Group") { $siteTemplate = "Team site" }
        if ($rootWebTemplate -eq "Site Page Publishing") { $siteTemplate = "Communication site" }
        if ($rootWebTemplate -eq "Team Site") 
        {
            ## Need to get web template to check if classic or modern
            ConnectToPnP $adminSiteUrl
            $site = Get-PnPTenantSite -Identity $siteUrl
            if ($site.Template -eq "STS#0") { $siteTemplate = "Team site (classic experience)" }
            if ($site.Template -eq "STS#3") { $siteTemplate = "Team site (no Microsoft 365 group)" }
        }
        # if ($rootWebTemplate -eq "Team Channel") 
        # {
        #     ConnectToPnP $siteUrl
        #     $site = Get-PnPSite -Includes RelatedGroupId
        #     $relatedGroup = $site.RelatedGroupId
        # }
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

    $logLine | Export-Csv -Path $logFileLocation -NoTypeInformation -Append
}

function GetSiteUsers([MemberTypes]$type, $graphObj, $pnpObj, $group)
{
    $userEmails = @()
        
    # drop in pnp to check if there are mutiple owners
    ConnectToPnP -siteUrl $site.WebUrl

    $domainObjs = @()

    try{
        if ($type -eq [MemberTypes]::Owner)
        {
            Write-Host "  Getting Site Owners" -ForegroundColor White
            $domainObjs = Get-PnPGroup -AssociatedOwnerGroup | Get-PnPGroupMember
        }

        if ($type -eq [MemberTypes]::Member)
        {
            Write-Host "  Getting Site Members" -ForegroundColor White
            $domainObjs = Get-PnPGroup -AssociatedMemberGroup | Get-PnPGroupMember
        }

        if ($type -eq [MemberTypes]::Visitor)
        {
            Write-Host "  Getting Site Visitors" -ForegroundColor White
            $domainObjs = Get-PnPGroup -AssociatedVisitorGroup | Get-PnPGroupMember
        }
    }
    catch {
        Write-Host "   Error getting site $type groups" -ForegroundColor Yellow
        Write-Host "    Likely because Site has bespoke permissions" -ForegroundColor Yellow


        return "ErrorGettingSPOGroups"
    }

    

    if ($type -eq [MemberTypes]::Admin)
    {
        Write-Host "  Getting Site Admins" -ForegroundColor White
        $domainObjs = Get-PnPSiteCollectionAdmin
    }


    foreach ($domainObj in $domainObjs)
    {
        # We may have Groups, Sec Groups or users

        # User - just add the email
        if ($domainObj.LoginName.Contains("|membership|"))
        {
            Write-Host "  Found: $($domainObj.Email) as $type" -ForegroundColor Green
            $userEmails += [Tuple]::Create($domainObj.Email, "User", "N/A")
            continue
        }

        # Role Manager - Special Groups
        if ($domainObj.LoginName.Contains("|rolemanager|"))
        {
            Write-Host "  Found: $($domainObj.Title) role as $type" -ForegroundColor Green
            $userEmails += [Tuple]::Create($domainObj.Title, "RoleManager", "N/A")
            continue
        }

        # Group - get the group *Members*
        # M365 Group membership can only be users, not groups so no need to recurse
        if ($domainObj.LoginName.Contains("|federateddirectoryclaimprovider|"))
        {
            Write-Host "  Found: $($domainObj.Email) (Group) as $type" -ForegroundColor DarkGreen
            $groupId = $domainObj.LoginName.Split("|")[2]

            # Check if we have a group to skip
            if ($groupsToSkip -contains $groupId.Substring(0, [Math]::Min($groupId.Length, 36)))
            {
                Write-Host "  Found: $($domainObj.Title) as $type - Skipping as in ignore list" -ForegroundColor Yellow
                $userEmails += [Tuple]::Create($domainObj.Title, "IgnoredGroup", "N/A")
                continue
            }

            ## Seems to only return ids, so we need to get the users
            ## Instances where we the group id isn't properly formed
            ## If we have a group connected site, group owners are the admin and owners rather than the members

            ## If we have a group we need to check if the group is the primary group
            ## If it is, we need to get the owners, if not we need to get the members
            $getOwners = $false
            if ($group -and ([MemberTypes]::Owner -eq $type -or [MemberTypes]::Admin -eq $type))
            {
                # We need to confirm if this is the base group
                # Get the drive owner and check if it's the same as the group
                $drive = Get-MgSiteDefaultDrive -SiteId $graphObj.Id
                if ($drive.Owner.AdditionalProperties.group.id -eq $groupId.Substring(0, [Math]::Min($groupId.Length, 36)))
                {
                    $getOwners = $true
                }
            }

            if ($getOwners)
            {
                Write-Host "  Getting Group Owners" -ForegroundColor Gray
                $members = Get-MgGroupOwner -GroupId $groupId.Substring(0, [Math]::Min($groupId.Length, 36)) -Property "userPrincipalName" -All
            }
            else {
                Write-Host "  Getting Group Members" -ForegroundColor Gray
                $members = Get-MgGroupMember -GroupId $groupId.Substring(0, [Math]::Min($groupId.Length, 36)) -Property "userPrincipalName" -All
            }

            foreach ($member in $members)
            {
                Write-Host "   Found: $($member.AdditionalProperties.userPrincipalName) as an infered $type" -ForegroundColor Green
                $userEmails += [Tuple]::Create($member.AdditionalProperties.userPrincipalName, "Group", $groupId.Substring(0, [Math]::Min($groupId.Length, 36)), $domainObj.Email)
            }
            continue
        }

        # Sec Group | Mail Enabled | OnPrem (Synced) ADGroup - get the group *Members* and recurse any child groups
        # Sec Group membership can be all sorts so we need to recurse
        if ($domainObj.LoginName.Contains("|tenant|"))
        {

            if ($domainObj.LoginName.Equals("c:0t.c|tenant|b71daa58-3cb3-4b97-a6e8-eae7f2a30f20") -or $domainObj.LoginName.Equals("c:0t.c|tenant|e8d578f9-c761-4097-b616-a1111909a468"))
            {
                Write-Host "   Found: $($domainObj.Title) as an infered $type" -ForegroundColor DarkGreen
                continue
            }

            ## We also need to ignore Global Administrator ownership
            if ($domainObj.Title.Equals("Global Administrator"))
            {
                Write-Host "   Found: $($domainObj.Title) as an infered $type" -ForegroundColor DarkGreen
                continue
            }

            Write-Host "  Found: $($domainObj.LoginName) (SecGroup) as $type" -ForegroundColor DarkGreen
            $groupId = $domainObj.LoginName.Split("|")[2]

            # Check if we have a group to skip
            if ($groupsToSkip -contains $groupId.Substring(0, [Math]::Min($groupId.Length, 36)))
            {
                Write-Host "  Found: $($domainObj.Title) as $type - Skipping as in ignore list" -ForegroundColor Yellow
                $userEmails += [Tuple]::Create($domainObj.Title, "IgnoredGroup", "N/A")
                continue
            }

            $groupEmail = $domainObj.Email
            if ($null -eq $groupEmail) 
            { 
                $groupEmail = $domainObj.Title
            }

            Write-Host "   Getting SecGroup Members" -ForegroundColor Gray

            try 
            {
                $members = Get-MgGroupMember -GroupId $groupId -Property "userPrincipalName,id,securityEnabled" -All -ErrorAction Stop
            }
            catch {
                Write-Host "    Error getting members for $($domainObj.Title)" -ForegroundColor Red
                Write-Host "    $($Error[0].Exception.Message)" -ForegroundColor Red
            }

            

            foreach ($member in $members)
            {
                if ($null -ne $member.AdditionalProperties.userPrincipalName)
                {
                    Write-Host "    Found: $($member.AdditionalProperties.userPrincipalName) as an infered $type" -ForegroundColor Green
                    ## We have a user
                    $userEmails += [Tuple]::Create($member.AdditionalProperties.userPrincipalName, "SecGroup", $groupId, $groupEmail)
                    continue
                }
                else {

                    # We may have an DL - need to ignore
                    if (!$member.AdditionalProperties.securityEnabled)
                    {
                        Write-Host "    Found: $($member.Id) (DL) - Does not effect permissions - Skipping" -ForegroundColor DarkGray
                        continue
                    }

                    Write-Host "    Found: $($member.Id) (group) as an infered $type" -ForegroundColor DarkGreen
                    ## We have a group and must get the members of that group
                    $users = @()
                    ## To format the output
                    $initialSpace = "     "
                    $userEmailsFromGroups += GetSecGroupMembers -groupId $member.Id -users $users -space $initialSpace

                    ## Issue with recursive function returning the tuples
                    foreach($userFomGroup in $userEmailsFromGroups)
                    {
                        $userEmails += [Tuple]::Create($userFomGroup, "SecGroup", $groupId, $groupEmail)
                    }
                }
            }
            continue
        }
        
    }

    # Return unique list of owners (emails)
    return $userEmails | Get-Unique
}

function ProcessSite($site, $usage)
{
    Write-Host "Processing Site ($($site.WebUrl))" -ForegroundColor White

    try {
        # Get the site object
        $siteObject = GetSite -siteUrl $site.WebUrl
    }
    catch {
        Write-Host " Unable to get Site details" -ForegroundColor Yellow
        Write-LogEntry -siteUrl $site.WebUrl -siteName $null -lockStatus $null -sharingCapability $null -domainList $null -user $null -siteUsage $usage -message "Unable to get Site details"
        return
    }
    

    # Get site Lock Status
    $lockStatus = $siteObject.LockState

    # Get the sharing capability
    $sharingCapability = $siteObject.SharingCapability

    # Get the domain list
    $domainList = $siteObject.SharingAllowedDomainList

    # Get the site name
    $siteName = $siteObject.Title

    # Is site locked? Won't be able to get owners or members
    if ($lockStatus -eq "NoAccess")
    {
        Write-Host " Site is locked with no access. Cannot get member details" -ForegroundColor Yellow
        Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $null -siteUsage $usage -message "Site is locked with no access"
        return
    }

    $group = $false
    if ($usage.'Root Web Template' -eq "Group")
    {
        $group = $true
    }

    try {
        ## Get the site owners
        Write-Host " Getting Site Owners" -ForegroundColor White
        $owners = GetSiteUsers -type ([MemberTypes]::Owner) -graphObj $site -pnpObj $siteObject -group $group

        if ($owners -eq "ErrorGettingSPOGroups")
        {
            Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $null -siteUsage $usage -message "Site has bespoke owner permissions"
            
        } else {
            foreach ($owner in $owners)
            {
                Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $owner -siteUsage $usage -type ([MemberTypes]::Owner)
            }
        }

        
    }
    catch {
        Write-Host "Error getting owners for $($site.WebUrl)" -ForegroundColor Red

        Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $null -siteUsage $usage -message "Error getting owners for $($site.WebUrl) - $($_.Exception.Message)"
    }
    
    if ($getMembers)
    {
        try {
            ## Get the site members
            Write-Host " Getting Site Members" -ForegroundColor White
            $members = GetSiteUsers -type ([MemberTypes]::Member) -graphObj $site -pnpObj $siteObject -group $group
    
            if ($members -eq "ErrorGettingSPOGroups")
            {
                Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $null -siteUsage $usage -message "Site has bespoke member permissions"
                
            } else {
    
                foreach ($member in $members)
                {
                    Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $member -siteUsage $usage -type ([MemberTypes]::Member)
                }
            }
        } catch {
            Write-Host "Error getting members for $($site.WebUrl)" -ForegroundColor Red
            Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $null -siteUsage $usage -message "Error getting members for $($site.WebUrl) - $($_.Exception.Message)"
        }
    }

    if ($getVisitors){
        try {
            ## Get the site visitors
            Write-Host " Getting Site Visitors" -ForegroundColor White
            $visitors = GetSiteUsers -type ([MemberTypes]::Visitor) -graphObj $site -pnpObj $siteObject -group $group
    
            if ($visitors -eq "ErrorGettingSPOGroups")
            {
                Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $null -siteUsage $usage -message "Site has bespoke visitor permissions"
                
            } else {
    
                foreach ($visitor in $visitors)
                {
                    Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $visitor -siteUsage $usage -type ([MemberTypes]::Visitor)
                }
            }
        } catch {
            Write-Host "Error getting visitors for $($site.WebUrl)" -ForegroundColor Red
            Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $null -siteUsage $usage -message "Error getting visitors for $($site.WebUrl) - $($_.Exception.Message)"
        }
    }

    try {
        ## Get the Site Admins
        Write-Host " Getting Site Admins" -ForegroundColor White
        $admins = GetSiteUsers -type ([MemberTypes]::Admin) -graphObj $site -pnpObj $siteObject -group $group
        foreach ($admin in $admins)
        {
            Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $admin -siteUsage $usage -type ([MemberTypes]::Admin)
        }
    }
    catch {
        Write-Host "Error getting admins for $($site.WebUrl)" -ForegroundColor Red
        Write-LogEntry -siteUrl $site.WebUrl -siteName $siteName -lockStatus $lockStatus -sharingCapability $sharingCapability -domainList $domainList -user $null -siteUsage $usage -message "Error getting owners for $($site.WebUrl) - $($_.Exception.Message)"
    }
    
    
}

## Recursive function to get members (users) of a security group
function GetSecGroupMembers($groupId, $users, $space)
{
    Write-Host "$($space) Getting Nested Group Members" -ForegroundColor White
    
    $members = Get-MgGroupMember -GroupId $groupId -Property "userPrincipalName,id" -All 

    foreach ($member in $members)
    {
        if ($null -ne $member.AdditionalProperties.userPrincipalName)
        {
            Write-Host "$($space) Found Member - $($member.AdditionalProperties.userPrincipalName) " -ForegroundColor Green
            ## We have a user
            $userEmails += $member.AdditionalProperties.userPrincipalName
        }
        else {
            Write-Host "$($space) Found Nested Group - $($member.Id) " -ForegroundColor DarkGreen
            ## We have a group and must get the members (that are users) of that group
            $users += GetSecGroupMembers -groupId $member.Id -users $users -space "$($space) "
        }
    }
    return $users
}

function GetSiteUsageReport($sites) {
    mkdir temp
    Get-MgReportSharePointSiteUsageDetail -Period D7 -OutFile .\temp\
    $siteUsage = Import-Csv .\temp\SharePointSiteUsageDetail*.csv
    ## delete report file
    Remove-Item .\temp\SharePointSiteUsageDetail*.csv
    Remove-Item .\temp\
    return $siteUsage | where { $sites -contains $_.'Site URL' }
}

function GetSite($siteUrl) {
    ConnectToPnP -siteUrl $adminSiteUrl
    $site = Get-PnPTenantSite -Identity $siteUrl
    return $site
}

##############################################
# Main
##############################################

## Connect to Graph
ConnectToMSGraph

## Get all sites
$sites = Get-Sites 

## hold the log entries
$outputObjs = @()

## Clear the CSV
$outputObjs | Export-Csv -Path $logFileLocation -NoTypeInformation -Force

## Get the site usage report for the sites we care about
$siteUsage = GetSiteUsageReport $sites.WebUrl

# Current Item
$currentItem = 0

## initilise progress bar
$percent = 0
Write-Progress -Activity "Processing Site $currentItem / $($sites.Count)" -Status "$percent% Complete:" -PercentComplete $percent
     
## Loop through sites
foreach ($site in $sites)
{
    # Get the site Usage
    $siteUsageEntry = $siteUsage | where { $_.'Site URL' -eq $site.WebUrl }

    ProcessSite -site $site -usage $siteUsageEntry

    $currentItem++
    $percent = [Math]::Round(($currentItem / $sites.Count) * 100)
    Write-Progress -Activity "Processed Site $currentItem / $($sites.Count)" -Status "$percent% Complete:" -PercentComplete $percent
}
