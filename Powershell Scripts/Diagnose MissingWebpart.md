
---

# Diagnose Missing Web Parts in SharePoint

## Overview

This document describes how to *diagnose and clean up missing web parts, event receivers, or assemblies* in a SharePoint content database.

These issues commonly appear after:

* Solution retractions
* Failed deployments
* Farm rebuilds or restores
* Missing or deleted custom assemblies
* Incomplete upgrades

Symptoms often include **“Error” web parts**, broken pages, or upgrade/configuration failures.

⚠️ **All SQL operations shown here are READ-ONLY unless explicitly deleting objects via PowerShell. Do not modify SharePoint databases directly.**

---

## Identify Missing Web Parts (SQL – Read Only)

This query finds pages that reference a **missing web part type**.

```sql
SELECT *
FROM AllDocs
INNER JOIN AllWebParts
    ON AllDocs.Id = AllWebParts.tp_PageUrlID
WHERE AllWebParts.tp_WebPartTypeID = '8f45bd4f-bd2c-93cc-0ea0-2288d323a0e7'
```

### What This Tells You

* Which pages still reference a web part
* Where the broken web part exists
* Helps identify cleanup scope

Replace the GUID with the **WebPartTypeID** you are investigating.

---

## Identify Missing Assemblies (SQL – Read Only)

This query identifies **event receivers** that reference a missing assembly.

```sql
SELECT *
FROM EventReceivers
WHERE Assembly = 'MainNewsEventRecevier, Version=1.0.0.0, Culture=neutral, PublicKeyToken=f49d00a34b63827b'
```

### Common Causes

* Solution (.wsp) removed but receivers remain
* Assembly version mismatch
* GAC deployment failed
* Incomplete upgrade

---

## Locate Site and Web (PowerShell)

Use this to confirm the **site and web context** before performing cleanup.

```powershell
$site = Get-SPSite -Limit All | Where-Object {
    $_.Id -eq "27E90E41-2426-4C6B-A243-2E6287875099"
}

$web = $site | Get-SPWeb -Limit All | Where-Object {
    $_.Id -eq "EBDE660A-EF0D-44F7-A7BD-AD720F3C710C"
}

$web.Url
```

---

## Remove Broken Event Receivers (PowerShell)

Once the list and receiver IDs are known, you can safely remove orphaned event receivers using PowerShell.

### Example: Remove Event Receivers by ID

```powershell
$list = $web.Lists | Where-Object {
    $_.Id -eq "ECA5473D-E9EE-44C9-B26D-1700670E57A7"
}

$er = $list.EventReceivers | Where-Object {
    $_.Id -eq "CD53BBCB-E568-4743-A907-2EA3A5F218D6"
}
$er.Delete()
```

Repeat for additional event receiver IDs as needed:

```powershell
$er = $list.EventReceivers | Where-Object {
    $_.Id -eq "99E667EF-06ED-47BA-ACEF-2248EA9E8706"
}
$er.Delete()
```

```powershell
$er = $list.EventReceivers | Where-Object {
    $_.Id -eq "355A19A2-DFAF-4D0C-A354-F59536F462F0"
}
$er.Delete()
```

```powershell
$er = $list.EventReceivers | Where-Object {
    $_.Id -eq "509A2829-0749-49DC-B10C-84A848C617C2"
}
$er.Delete()
```

---

## Best Practices ⚠️

* **Never update SharePoint databases directly**
* Use SQL **only for diagnostics**
* Always confirm IDs before deleting anything
* Test cleanup in **non-production** first
* Take database backups before remediation

---

## Common Symptoms This Fixes

* “Missing Web Part” errors on pages
* Upgrade failures during PSConfig
* Timer job or feature activation errors
* ULS errors referencing missing assemblies
* Broken list functionality

---

## References

* Original script inspiration:
  [http://get-spscripts.com/2011/08/diagnose-missingwebpart-and.html](http://get-spscripts.com/2011/08/diagnose-missingwebpart-and.html)

---

