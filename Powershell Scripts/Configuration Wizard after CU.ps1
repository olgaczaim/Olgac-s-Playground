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