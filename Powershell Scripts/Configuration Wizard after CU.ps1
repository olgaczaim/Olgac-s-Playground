#Running configuration wizard on management shell

psconfig -cmd upgrade -inplace b2b -wait

#Or

PSConfig.exe -cmd upgrade -inplace b2b -force -cmd applicationcontent
-install -cmd installfeatures

#There are actually 3 commands in one:

#psconfig.exe –cmd upgrade
#psconfig.exe –cmd applicationcontent
#psconfig.exe –cmd installfeatures

#Psconfig.exe -cmd upgrade => Perform SharePoint upgrade. This command is automatically executed when running the SharePoint Products configuration Wizard if SharePoint needs to be upgraded.
#inplace b2b ==> If b2b is chosen, then an in-place build to build upgrade will be performed. The other option is v2v (version to version)
#force ==> The SharePoint Products Configuration Wizard stops any currently running upgrade actions, and then restart upgrade.
#-cmd applicationcontent ==> Manages Shared application content
#-cmd installfeatures ==> registers the SharePoint features in the server farm that are located on this server.

#------------------------------------------------------------------
#After the upgrade do not forgot to remove mirroring db logging before the running conf wizard
#------------------------------------------------------------------


#Before You Begin
#Prerequisites and Restrictions
#This task is supported only on primary replicas. You must be connected to the server instance that hosts the primary replica.
#Security
#Permissions
#Requires ALTER AVAILABILITY GROUP permission on the availability group, CONTROL AVAILABILITY GROUP permission, ALTER ANY AVAILABILITY GROUP permission, or CONTROL SERVER permission.
#
#Using SQL Server Management Studio
#To remove an availability database
#
#In Object Explorer, connect to the server instance that hosts the primary replica of the database or databases to be removed, and expand the server tree.
#
#Expand the Always On High Availability node and the Availability Groups node.
#
#Select the availability group, and expand the Availability Databases node.
#
#This step depends on whether you want to remove multiple databases groups or only one database, as follows:
#
#To remove multiple databases, use the Object Explorer Details pane to view and select all the databases that you want to remove. For more information, see Use the Object Explorer Details to Monitor Availability Groups (SQL Server Management Studio).
#
#To remove a single database, select it in either the Object Explorer pane or the Object Explorer Details pane.
#
#Right-click the selected database or databases, and select Remove Database from Availability Group in the command menu.
#
#In the Remove Databases from Availability Group dialog box, to remove all the listed databases, click OK. If you do not want to remove all them, click Cancel.