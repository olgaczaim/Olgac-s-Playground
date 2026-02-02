
---

# Running SharePoint Configuration Wizard via PowerShell (PSConfig)

## Overview

This document explains how to run the **SharePoint Products Configuration Wizard** from the **SharePoint Management Shell** using `psconfig.exe`.

Running PSConfig manually is commonly required after:

* Installing a **Cumulative Update (CU)**
* Failed or interrupted upgrades
* Database schema mismatches
* Restoring or moving SharePoint databases
* Rebuilding farm servers

## Basic Commands

### Run Upgrade (In-Place, Build-to-Build)

```powershell
psconfig -cmd upgrade -inplace b2b -wait
```

### Full PSConfig Command (Recommended)

```powershell
psconfig.exe -cmd upgrade -inplace b2b -force `
  -cmd applicationcontent `
  -cmd installfeatures
```

## What These Commands Do

PSConfig is effectively running **three commands in sequence**:

```text
psconfig.exe -cmd upgrade
psconfig.exe -cmd applicationcontent
psconfig.exe -cmd installfeatures
```

### Command Breakdown

| Parameter                 | Description                                                |
| ------------------------- | ---------------------------------------------------------- |
| `-cmd upgrade`            | Performs the SharePoint upgrade process                    |
| `-inplace b2b`            | Build-to-build in-place upgrade (CU-level upgrade)         |
| `v2v`                     | Version-to-version upgrade (e.g. SharePoint 2013 → 2016)   |
| `-force`                  | Stops any running upgrade actions and restarts the upgrade |
| `-cmd applicationcontent` | Manages shared application content                         |
| `-cmd installfeatures`    | Registers SharePoint features on the server                |
| `-wait`                   | Keeps the console open until completion                    |

## Running from Management Shell

* Start **SharePoint Management Shell**
* Run **as Administrator**
* You may need to be logged in as a **Farm Administrator**

## Common Upgrade Error

### Error Message

```text
Failed to upgrade SharePoint Products

An exception of type
Microsoft.SharePoint.PostSetupConfiguration.PostSetupConfigurationTaskException was thrown.

Upgrade [SPContentDatabase Name=WSS_Content_47xxxxxxxxx] failed.
The upgraded database schema doesn’t match the TargetSchema

Upgrade Timer job is exiting due to exception:
Microsoft.SharePoint.Upgrade.SPUpgradeException
```

### Cause

* One or more content databases were **not fully upgraded**
* PSConfig exited or failed part-way through
* Database schema versions are inconsistent

### Resolution

1. Open **SharePoint 2016 Management Shell**
2. Run the following command to upgrade incomplete databases:

```powershell
Get-SPContentDatabase |
  Where-Object { $_.NeedsUpgrade -eq $true } |
  Upgrade-SPContentDatabase -Confirm:$false
```

3. Once completed, re-run PSConfig:

```powershell
psconfig.exe -cmd upgrade -inplace b2b
```

This forces SharePoint to complete any unfinished database upgrades.

## Important: Always On / WSS_Logging Database ⚠️

### Before Running PSConfig

If your farm uses **SQL Server Always On Availability Groups**, you **must remove the `WSS_Logging` database** from the Availability Group **before** running the Configuration Wizard.

Failure to do this can cause upgrade failures.

---

## Removing `WSS_Logging` from Availability Group

### Prerequisites

* Must be connected to the **primary replica**
* Required permissions:

  * `ALTER AVAILABILITY GROUP`
  * `CONTROL AVAILABILITY GROUP`
  * or `CONTROL SERVER`

### Steps (SQL Server Management Studio)

1. Open **SQL Server Management Studio**
2. Connect to the primary replica
3. Expand:

   ```
   Always On High Availability
     → Availability Groups
       → Availability Databases
   ```
4. Right-click **WSS_Logging**
5. Select **Remove Database from Availability Group**
6. Complete the wizard

---

## Add `WSS_Logging` Back After Upgrade

Once PSConfig and the Configuration Wizard complete successfully:

1. Open **SQL Server Management Studio**
2. Expand:

   ```
   Always On High Availability
     → Availability Groups
   ```
3. Right-click **Availability Databases**
4. Select **Add Database**
5. Add **WSS_Logging**
6. Complete the wizard

---

## Best Practices

* Always run PSConfig on **one server at a time**
* Ensure **all SharePoint services are stopped** before upgrade
* Verify database backups before starting
* Monitor ULS logs during execution
* Never interrupt PSConfig once started

## Common Use Cases

* Post-CU configuration
* Fixing failed upgrade states
* Completing partial database upgrades
* Farm recovery after restore
* Re-running configuration after patching

---
