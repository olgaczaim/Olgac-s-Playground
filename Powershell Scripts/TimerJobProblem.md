
---

# SharePoint Timer Jobs Not Running After Upgrade / Restore

## Overview

Over the course of a couple of days working on a SharePoint environment, a *perfect storm* of changes led to a particularly frustrating issue where **SharePoint Timer Jobs simply refused to behave**.

This was not a normal or clean scenario. The environment had been through:

* Major database work to split a single, very large site collection database into smaller ones
* Installation of a service pack
* Restoring a copy of the **Production** database into a **Test** environment
* Multiple upgrade runs (and, regrettably, a few human errors along the way)

The result was a cascade of problems that consumed far more time than they reasonably should have.

## Symptoms

After finally getting the environment back into a usable state, timer-service-dependent functionality was clearly broken.

Observed behaviour included:

* Scheduled timer jobs **existed and appeared to run**
* The **“Running Jobs”** page was *always empty*
* Editing timer job schedules had **no effect**
* Clicking **“Run Now”** on any timer job did absolutely nothing
* Custom timer jobs and built-in jobs (e.g. Content Organizer processing) were affected

Standard troubleshooting steps were attempted with no success:

* Resetting the timer job cache
* Reverting permission changes
* Restarting servers (more than once)

Still nothing. The timer jobs simply would not run on demand.

## Root Cause

After a lot of searching (ironically, Bing was completely useless here), the breakthrough came from a Microsoft Support article:

> **“Administrative Timer jobs not running after upgrade”**

The key detail:

> Every SharePoint server has one `SPTimerServiceInstance` object which represents the **SPTimerV4** Windows Service.
> In certain circumstances (typically after an upgrade), the Windows service may be running, but the corresponding `SPTimerServiceInstance` object is **not Online**.

When this happens:

* Administrative timer jobs will not execute correctly
* Services that depend on timer jobs may silently fail
* “Run Now” does nothing

The likely cause is:

> *“An unexpected event during the upgrade prevented the timer service instance object from being brought back online.”*

Given the number of upgrades and restores performed, this checked out.

## Solution

Microsoft provided a PowerShell script to detect and fix timer service instances that are not online.

### PowerShell Script

```powershell
$farm = Get-SPFarm

$disabledTimers = $farm.TimerService.Instances | Where-Object {
    $_.Status -ne "Online"
}

if ($disabledTimers) {
    foreach ($timer in $disabledTimers) {
        Write-Host "Timer service instance on server" $timer.Server.Name "is not Online. Current status:" $timer.Status
        Write-Host "Attempting to set the status of the service instance to Online"

        $timer.Status = [Microsoft.SharePoint.Administration.SPObjectStatus]::Online
        $timer.Update()
    }
}
else {
    Write-Host "All Timer Service Instances in the farm are online! No problems found"
}
```

> Yes, it’s ugly. At this point, desperation outweighed aesthetics.

### Final Step (Important!)

Although the script detected and fixed a disabled timer instance, **the issue was not fully resolved until the SharePoint Timer Service was restarted**.

Once the timer service was restarted:

* Timer jobs began running again
* “Run Now” worked as expected
* Dependent functionality sprang back to life immediately

💥 Problem solved.

## Conclusion

This issue is easy to miss because:

* The Windows service appears to be running
* Scheduled jobs may still execute
* There are no obvious errors

If you’ve recently:

* Upgraded SharePoint
* Restored databases
* Moved environments
* Re-run configuration or upgrade steps

…and timer jobs are behaving strangely, **check the status of the `SPTimerServiceInstance` objects**.

## Footnote

This exists mainly so that when (not *if*) this happens again in the future, it’s easy to find.
If it helps someone else avoid losing two days to a broken timer service, even better.

Such is the life of a SharePoint geek. ❤️

---
>ref: https://matt-thornton.net/tech/powershell-tech/yes-timer-jobs-not-running/
