
---

# Managing Windows Services via Command Line

## Overview

This document outlines how to **delete**, **create**, and **update** Windows Services using the built-in `sc.exe` command-line tool.

These steps are useful when deploying applications, updating service executables, or cleaning up old or broken services — especially in server, SharePoint, or CI/CD environments.

> ⚠️ All commands must be run from an **elevated Command Prompt** (Run as Administrator).

## Delete an Existing Windows Service

Use this when removing an old or incorrectly configured service.

```cmd
sc delete "Your Service Name"
```

### Notes

* The service is marked for deletion and may disappear after a reboot
* Ensure the service is **stopped** before deleting it

## Create a New Windows Service

Create a new service and point it at the executable you want to run.

```cmd
sc create "Your New Service Name" binPath= "C:\directory\some.exe"
```

### Important Syntax Rules

* There **must** be a space after the `=` sign
* `binPath=` is case-insensitive, but spacing is not optional
* Use quotes if the path contains spaces

### Optional Parameters

```cmd
sc create "Your New Service Name" ^
  binPath= "C:\directory\some.exe" ^
  start= auto ^
  DisplayName= "Friendly Service Name"
```

## Update an Existing Windows Service

Windows services **cannot be edited directly**, but you can update most settings using `sc config`.

### Update the Executable Path

```cmd
sc config "Your Service Name" binPath= "C:\directory\updated.exe"
```

### Change Startup Type

```cmd
sc config "Your Service Name" start= auto
```

Startup options:

* `auto` – Automatic
* `demand` – Manual
* `disabled` – Disabled

### Change the Service Account

```cmd
sc config "Your Service Name" obj= "DOMAIN\ServiceAccount" password= "password"
```

> ⚠️ Be careful with passwords in command history or scripts.

## Restart the Service After Changes

After updating configuration, restart the service to apply changes:

```cmd
net stop "Your Service Name"
net start "Your Service Name"
```

## Verify Service Configuration

To check the current service configuration:

```cmd
sc qc "Your Service Name"
```

## Common Pitfalls

* Forgetting the space after `=`
* Running commands without admin rights
* Updating the executable without restarting the service
* Leaving orphaned services after failed deployments

## Use Cases

* Deploying background workers
* Updating Windows services during releases
* Cleaning up failed service installs
* Automating server setup scripts

---


