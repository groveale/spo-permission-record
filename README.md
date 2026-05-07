# spo-permission-record
Export Site Level Permissions across SPO

This script can be used to export Site Collection Level permissions. Export is a line per users. Admins, Owners, Members and Visitors are exported. Any bespoke site permissions are ignored.

This includes both group connected and non group connected sites.

## Versions

### `GetSitePermissionDataSet.ps1` — Local / Interactive
Runs locally with certificate-based app-only auth for both Graph and PnP. Outputs CSV to the local `Output\` folder.

### `GetSitePermissionDataSet-Automation.ps1` — Azure Automation Account
Designed to run as an Azure Automation runbook. Uses Managed Identity for Graph API calls and certificate-based app registration for PnP. Outputs CSV directly to a SharePoint document library via the Graph API.

This version avoids the **PnP.PowerShell / Microsoft.Graph assembly conflict** (see below) by pinning module versions in the Automation Account.

## Assembly Conflict — PnP.PowerShell & Microsoft.Graph

`PnP.PowerShell` bundles its own copy of `Microsoft.Graph.Core`. When both PnP and the Microsoft.Graph SDK are loaded in the same PowerShell session, the first assembly version loaded wins — .NET cannot unload or swap assemblies mid-session.

If the Graph SDK expects a newer `Microsoft.Graph.Core` than what PnP bundles, you get errors like:

```
Could not load type 'Microsoft.Graph.Authentication.AzureIdentityAccessTokenProvider'
from assembly 'Microsoft.Graph.Core, Version=1.25.1.0'
```

### How the Automation Account version solves this

Azure Automation lets you pin exact module versions. By installing **Microsoft.Graph.\* 1.28.0** alongside **PnP.PowerShell 3.1.0**, both modules use a compatible `Microsoft.Graph.Core` (v1.25.x) and the conflict is eliminated.

| Module | Pinned Version |
|---|---|
| PnP.PowerShell | 3.1.0 |
| Microsoft.Graph.Authentication | 1.28.0 |
| Microsoft.Graph.Sites | 1.28.0 |
| Microsoft.Graph.Groups | 1.28.0 |
| Microsoft.Graph.Reports | 1.28.0 |

> **Do not** install Microsoft.Graph SDK 2.x in the same Automation Account — it will reintroduce the conflict.

## Dependencies

Both `PnP.PowerShell` and Microsoft Graph PowerShell modules are used in the scripts.

### Local version — App Registration (certificate auth)

An app registration with the following permissions is required:

#### Graph Permissions
* Sites.Read.All
* GroupMember.Read.All
* User.Read.All
* Reports.Read.All

#### SPO Permissions
* Sites.Read.All
* Site.FullControl.All

### Automation version — Managed Identity + App Registration

The Automation Account's **Managed Identity** is used for Graph API calls and needs:

* Sites.Read.All
* GroupMember.Read.All
* Reports.Read.All
* Files.ReadWrite.All (for uploading the CSV to SharePoint)

A separate **App Registration** (with certificate) is still required for PnP, as PnP.PowerShell does not support Managed Identity:

* Sites.Read.All
* Site.FullControl.All

### Automation parameters

| Parameter | Description |
|---|---|
| `DriveId` | SharePoint document library drive ID for CSV upload |
| `OutputFolderPath` | Folder path within the drive (e.g. `Reports/Permissions`) |
| `AdminSiteUrl` | SPO Admin centre URL |
| `ClientId` | App registration client ID (for PnP) |
| `TenantId` | Tenant ID or domain |
| `CertificateThumbprint` | Certificate thumbprint (for PnP) |
| `AllSites` | Process all sites (`$true`) or only listed sites |
| `SiteList` | Array of site URLs to process (ignored if `AllSites` is `$true`) |
| `GetMembers` | Include site members in output |
| `GetVisitors` | Include site visitors in output |
| `GroupsToSkip` | Array of group names/IDs to skip expanding |

## Broken Inheritance

Not covered

Note. This script only covers Site Collection level permissions. Broken inheritance at SubSite, Library or Item level is not picked up.

## Sharing Links

Not covered

