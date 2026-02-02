
---

# Cannot Open MSSQL Backup Directory

## Overview

This issue occurs when **SQL Server is unable to access a backup directory**, even though the path exists and appears accessible from Windows Explorer.

The most common cause is **incorrect NTFS permissions** on the backup folder for the account running the SQL Server service.

## Symptoms

You may see errors such as:

* *Cannot open backup device*
* *Access is denied*
* Backup jobs fail when run manually or via SQL Server Agent

## Root Cause

SQL Server accesses the filesystem **using the account defined in the SQL Server service**, not your logged-in user.

If that service account does not have permission to the backup directory, SQL Server will fail to read or write backups.

## Resolution

### 1. Identify the SQL Server Service Account

1. Open **Windows Services**
   `Start → Administrative Tools → Services`

2. Locate the service:

   ```
   SQL Server (MSSQLSERVER)
   ```

   *(For named instances, this may appear as `SQL Server (InstanceName)`)*

3. Check the **Log On As** column

   * If it’s not visible, enable it via:

     * Right-click the column header → **Select Columns**
     * Enable **Log On As**

4. Note the account listed — this is the account SQL Server uses to access the filesystem.

### 2. Grant Folder Permissions

1. Open **File Explorer**

2. Right-click the SQL backup directory

3. Select **Properties**

4. Go to:

   * **Security** tab (NTFS permissions)
   * *(Optionally)* **Sharing** tab if using a network share

5. Grant the SQL Server service account:

   * **Read / Write** permissions
   * Or **Modify** (recommended for backup directories)

6. Apply and close

### 3. Retry the Backup

Once permissions are applied:

* Retry the backup operation
* SQL Server should now be able to access the directory successfully

## Notes

* Grant permissions to the **service account**, not individual users
* If using SQL Server Agent, ensure the Agent service account also has access
* Network shares require both **share** and **NTFS** permissions

## Common Service Accounts

Examples you may see in **Log On As**:

* `NT SERVICE\MSSQLSERVER`
* `NT SERVICE\MSSQL$InstanceName`
* `DOMAIN\SqlServiceAccount`
* `LocalSystem` *(not recommended)*

## Use Cases

* SharePoint database backups
* SQL Server maintenance plans
* Manual `.bak` file creation
* Restored environments with new folder structures

---

