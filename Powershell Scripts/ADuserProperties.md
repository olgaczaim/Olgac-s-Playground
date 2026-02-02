
---

# Get User Properties from Active Directory Using PowerShell

## Overview

This document shows a few simple ways to retrieve **Active Directory user properties** using **Windows PowerShell** and `System.DirectoryServices.DirectorySearcher`.

These snippets are useful in SharePoint environments (or any AD-integrated system) when you need to:

* Look up a user by `sAMAccountName`
* Inspect all available AD attributes
* Retrieve group memberships (`memberOf`)

No SharePoint-specific cmdlets are required — this uses standard .NET directory services.

## Prerequisites

* Windows PowerShell
* Access to Active Directory
* Permissions to query user objects

## Set the Username

Replace the placeholder value with the account you want to query:

```powershell
$username = "xxxxx"
```

## Retrieve All User Properties

This example creates a `DirectorySearcher`, filters on `person` objects, and returns **all available properties** for the user.

```powershell
$searcher = New-Object System.DirectoryServices.DirectorySearcher
$searcher.Filter = "(&(objectCategory=person)(samAccountName=$($username)))"
$searcher.FindOne().Properties
```

This is useful when you’re not sure which attributes you need and want to inspect everything returned by AD.

## One-Liner: Get User Properties

A more compact version of the same query:

```powershell
(New-Object System.DirectoryServices.DirectorySearcher("(&(objectCategory=User)(samAccountName=$($username)))")).Properties
```

Good for quick testing or ad-hoc scripts.

## Get User Group Memberships

To retrieve the groups the user belongs to (`memberOf`):

```powershell
(New-Object System.DirectoryServices.DirectorySearcher("(&(objectCategory=User)(samAccountName=$($username)))")).FindOne().GetDirectoryEntry().memberOf
```

This returns the distinguished names (DNs) of the groups. You can further process these if you need friendly group names.

## Notes

* Results depend on domain permissions and AD schema
* Some attributes may return multiple values
* `memberOf` does **not** include primary group membership

## Use Cases

* SharePoint user troubleshooting
* Validating AD profile sync data
* Auditing permissions
* Debugging authentication or claims issues

---
