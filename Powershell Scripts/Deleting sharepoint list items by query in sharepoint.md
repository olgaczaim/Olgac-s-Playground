
---

# Delete SharePoint List Items by CAML Query (PowerShell)

## Overview

This script deletes SharePoint list items **based on a CAML query** that matches a specific field value.
It iterates over a **range of index values**, queries matching list items, and deletes them in batches.

This approach is useful when:

* Cleaning up migrated or duplicated data
* Deleting items by custom ID or index field
* Bulk-removing items without deleting the entire list
* Working around list view thresholds using paging

⚠️ **This script permanently deletes list items.**

## Prerequisites

* SharePoint **on-premises**
* SharePoint Management Shell or PowerShell running as Administrator
* Farm or Site Collection Administrator permissions
* `Microsoft.SharePoint.PowerShell` snap-in available

## Script

```powershell id="pk1s2a"
for ($i = 59; $i -le 122; $i++) {

    Write-Host "Preparing to delete items with index - $i"

    Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

    $web  = Get-SPWeb "https://xxxxxxxxxxx"
    $list = $web.Lists["ListName"]

    $query = New-Object Microsoft.SharePoint.SPQuery

    # CAML Query: Match items by custom index field
    $internalFieldName = 'xID'
    $caml = "<Where>
                <Eq>
                    <FieldRef Name='$internalFieldName' />
                    <Value Type='Text'>$i</Value>
                </Eq>
             </Where>"

    $query.Query = $caml

    Write-Host "Query - $caml"

    $query.ViewAttributes = "Scope='Recursive'"
    $query.RowLimit = 1000
    $query.ViewFields = "
        <FieldRef Name='ID' />
        <FieldRef Name='Title' />
        <FieldRef Name='xID' />
    "
    $query.ViewFieldsOnly = $true

    do {
        $listItems = $list.GetItems($query)
        $query.ListItemCollectionPosition = $listItems.ListItemCollectionPosition

        foreach ($item in $listItems) {
            Write-Host "Deleting Item - $($item.ID) - $($item.Title) - $($item['xID'])"
            $list.GetItemById($item.ID).Delete()
        }
    }
    while ($query.ListItemCollectionPosition -ne $null)
}
```

## How It Works

1. Loops through a numeric range (`59 → 122`)
2. Builds a CAML query for each value
3. Queries list items where a custom field (`xID`) equals the current index
4. Uses paging (`RowLimit` + `ListItemCollectionPosition`)
5. Deletes each matching item

## Key Configuration Points

| Setting               | Description                                   |
| --------------------- | --------------------------------------------- |
| `$i = 59; $i -le 122` | Range of values to match                      |
| `xID`                 | Internal name of the field used for filtering |
| `Scope='Recursive'`   | Includes items in folders                     |
| `RowLimit = 1000`     | Controls batch size                           |

## Notes

* Field name must be the **internal name**, not the display name
* CAML `Value Type` must match the field type (`Text`, `Number`, etc.)
* Paging avoids threshold issues on large lists
* Script can be safely re-run (idempotent per index value)

## Warnings ⚠️

* Items are **permanently deleted**
* Recycle Bin may be bypassed depending on configuration
* Always test CAML queries using `GetItems()` without deletion first
* Consider running against a backup or test environment

## Common Use Cases

* Cleaning corrupted or duplicated list data
* Removing items by external system ID
* Post-migration cleanup
* Large-scale list maintenance

## Possible Enhancements

* Add `-WhatIf` / dry-run mode
* Log deleted item IDs to a file
* Parameterize site URL, list name, field name, and range
* Add try/catch per delete for resilience

---

