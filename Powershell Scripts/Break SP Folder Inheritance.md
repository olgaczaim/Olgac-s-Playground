
---

# Break SharePoint Folder Permission Inheritance (PowerShell)

## Overview

This script iterates through **all folders in a SharePoint document library** and **breaks permission inheritance** for each folder that is currently inheriting permissions from its parent.

It is useful when:

* You need folder-level permissions
* A library was migrated or restored and permissions need reapplying
* You want to prepare folders for custom security assignments

The script preserves existing permissions when inheritance is broken.

## Prerequisites

* SharePoint Management Shell **or** PowerShell running as Administrator
* Farm or Site Collection Administrator permissions
* On-prem SharePoint (uses `Microsoft.SharePoint.PowerShell` snap-in)

## Script

```powershell
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# Variables for Processing
$SiteURL  = "http://acme/supplies"
$ListName = "Documents"

Try {
    $web = Get-SPWeb $SiteURL

    # Get the List
    $List = $web.GetList($SiteURL + "/" + $ListName)
    Write-Host -ForegroundColor Red $List.Title

    # Query for folders only
    $Query = New-Object Microsoft.SharePoint.SPQuery
    $Query.Query = "
        <Query>
            <Where>
                <Eq>
                    <FieldRef Name='ContentType' />
                    <Value Type='Computed'>Folder</Value>
                </Eq>
            </Where>
        </Query>
    "

    $ListItems = $List.GetItems($Query)

    $i = 0

    # Loop through each folder
    foreach ($Item in $ListItems) {
        if ($Item.HasUniqueRoleAssignments -eq $false) {
            $i++
            $Item.BreakRoleInheritance($true)
            Write-Host -ForegroundColor Green `
                "Permission Inheritance Broken for $($Item['FileLeafRef'])"
        }
    }

    Write-Host -ForegroundColor Yellow "Folders processed: $i"
}
Catch {
    Write-Host -ForegroundColor Red "Error:" $_.Exception.Message
}
```

## What the Script Does

1. Loads the SharePoint PowerShell snap-in
2. Connects to the specified site
3. Retrieves the target document library
4. Queries for items where `ContentType = Folder`
5. Loops through each folder
6. Breaks permission inheritance **only if it is currently inherited**
7. Preserves existing role assignments (`BreakRoleInheritance($true)`)

## Key Notes

* `BreakRoleInheritance($true)`
  → Copies existing permissions instead of wiping them
* Only folders are processed — files are ignored
* Existing unique permissions are left untouched
* Safe to re-run (idempotent for already-broken folders)

## Common Use Cases

* Post-migration permission fixes
* Preparing libraries for granular access control
* SharePoint cleanup after restores
* Automating security setup

## Warnings ⚠️

* Breaking inheritance at scale can increase permission complexity
* Large libraries may take time to process
* Always test in **non-production** first

## Possible Enhancements

* Assign permissions after breaking inheritance
* Filter specific folder paths
* Add logging to file
* Add support for files as well as folders

---
