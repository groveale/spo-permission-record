# spo-permission-record
Export Site Level Permissions across SPO

This script can be used to export Site Collection Level permissions. Export is a line per users. Admins, Owners, Members and Visitors are exported. Any bespoke site permissions are ignored.

This includes both group connected and non group connected sites.

## Dependencies

This script use app only authentication. An app registration with the following permissions are required

### Graph Permissions
* Sites.Read.All
* GroupMember.Read.All
* User.Read.All
* Reports.Read.All

### SPO Permissions
* Sites.Read.All
* Site.FullControl.All

Both `Pnp.PowerShell` and MSGraph Powershell modules are used in the script

## Broken Inheritance

Not covered

Note. This script only covers Site Collection level permissions. Broken inheritance at SubSite, Library or Item level is not picked up.

## Sharing Links

Not covered

