
---

# Clear All Items from a SharePoint List (PowerShell)

## Overview

This script **deletes all items from a SharePoint list** using **batch processing** for performance.

Instead of deleting items one-by-one, it builds a CAML `<ows:Batch>` request and submits it in a single operation, making it significantly faster for large lists.

⚠️ **This operation is destructive and irreversible.**

## Prerequisites

* SharePoint **on-premises**
* SharePoint Management Shell or PowerShell running as Administrator
* Farm or Site Collection Administrator permissions
* `Microsoft.SharePoint.PowerShell` snap-in available

## Parameters

| Parameter   | Description                 | Required |
| ----------- | --------------------------- | -------- |
| `-weburl`   | URL of the SharePoint site  | Yes      |
| `-listname` | Name of the SharePoint list | Yes      |

## Script

```powershell id="c6hdg1"
# Clear SharePoint List
#################################################################################

param($weburl, $listname)

if ($weburl -eq $null -or $listname -eq $null) {
    Write-Host -ForegroundColor Red "-weburl or -listname are null."
    return
}

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

$web  = Get-SPWeb $weburl
$list = $web.Lists[$listname]

$stringBuilder = New-Object System.Text.StringBuilder

try {
    $stringBuilder.Append(
        "<?xml version=`"1.0`" encoding=`"UTF-8`"?>" +
        "<ows:Batch OnError=`"Return`">"
    ) > $null

    $i = 0

    $spQuery = New-Object Microsoft.SharePoint.SPQuery
    $spQuery.ViewFieldsOnly = $true

    $items = $list.GetItems($spQuery)
    $count = $items.Count

    while ($i -le ($count - 1)) {
        Write-Host $i
        $item = $items[$i]

        $stringBuilder.AppendFormat("<Method ID=`"{0}`">", $i) > $null
        $stringBuilder.AppendFormat("<SetList Scope=`"Request`">{0}</SetList>", $list.ID) > $null
        $stringBuilder.AppendFormat("<SetVar Name=`"ID`">{0}</SetVar>", $item.ID) > $null
        $stringBuilder.Append("<SetVar Name=`"Cmd`">Delete</SetVar>") > $null
        $stringBuilder.Append("</Method>") > $null

        $i++
    }

    $stringBuilder.Append("</ows:Batch>") > $null

    $web.ProcessBatchData($stringBuilder.ToString()) > $null
}
catch {
    Write-Host -ForegroundColor Red $_.Exception.ToString()
}

Write-Host -ForegroundColor Green "Done."
```

## How It Works

1. Validates input parameters
2. Loads the SharePoint PowerShell snap-in
3. Retrieves the target site and list
4. Queries all list items
5. Builds a CAML batch delete request
6. Deletes all items in one operation using `ProcessBatchData`

## Usage Example

```powershell
.\ClearList.ps1 -weburl http://app-server -listname RecordList
```

## Performance Notes

* Much faster than deleting items individually
* Suitable for large lists
* Avoids throttling and UI timeouts

## Warnings ⚠️

* Items are **permanently deleted**
* Recycle Bin may be bypassed depending on configuration
* Always test in **non-production** environments first
* Consider backing up list data before execution

## Common Use Cases

* Resetting staging or test environments
* Cleaning up migrated data
* Reinitializing lists for reprocessing
* Automated environment rebuilds
* Err when deleting: The attempted operation is prohibited because it exceeds the list view threshold.
* Err: This list is too big to delete
* If you don't want to modify List View Threshold

## Possible Enhancements

* Add confirmation prompt
* Filter items by CAML query
* Log deleted item IDs to file
* Support list GUID instead of name

---

