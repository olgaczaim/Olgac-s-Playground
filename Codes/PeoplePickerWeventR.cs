public override void ItemDeleting(SPItemEventProperties properties)
{
    //base.ItemDeleting(properties);
    using (SPWeb web = properties.Web)
    {
        SPUser user = web.CurrentUser;

        bool isMember = user.Groups.Cast<SPGroup>().Any(g => g.Name.Equals("TestGroup"));
      
       //if current user is admin or member of TestGrop, can delete item
        if (user.LoginName != @"SHAREPOINT\system" || !isMember)
        {
            properties.Status = SPEventReceiverStatus.CancelWithError;
            properties.ErrorMessage = "No Delete Permission";
        }
    }

}


public override void ItemAdding(SPItemEventProperties properties)
{
    //base.ItemDeleting(properties);
    using (SPWeb web = properties.Web)
    {
        SPUser user = web.CurrentUser;
        bool isMember = user.Groups.Cast<SPGroup>().Any(g => g.Name.Equals("TestGroup"));

      //if current user is admin or member of TestGrop, can add item
        if (user.LoginName != @"SHAREPOINT\system" || !isMember)
        {
            properties.Status = SPEventReceiverStatus.CancelWithError;
            properties.ErrorMessage = "No Item Add Permission";
        }
    }

}

//Example of before and after values for list items during updating event
public override void ItemUpdating(SPItemEventProperties properties)
{
    using (SPWeb web = properties.Web)
    {
        base.ItemUpdating(properties);
        string titleBefore = properties.ListItem["Title"].ToString();
        string titleAfter = properties.AfterProperties["Title"].ToString();

        string dateBefore = Convert.ToDateTime(properties.ListItem["Date1"]).ToShortDateString();
        string dateAfter = Convert.ToDateTime(properties.AfterProperties["Date1"]).ToShortDateString();

        SPFieldUser peopleUserPick = (SPFieldUser)properties.List.Fields.GetFieldByInternalName("PeoplePickerTest");
        SPFieldUserValue peopleValBefore = new SPFieldUserValue(web, properties.ListItem[peopleUserPick.InternalName].ToString());
        SPFieldUserValue peopleValAfter = new SPFieldUserValue(web, properties.AfterProperties[peopleUserPick.InternalName].ToString());

        string peopleBefore = peopleValBefore.User.LoginName;
        string peopleAfter = peopleValAfter.LookupValue;

        if (titleBefore != titleAfter || dateBefore != dateAfter || peopleBefore != peopleAfter)
        {
            properties.Status = SPEventReceiverStatus.CancelWithError;
            properties.ErrorMessage = "You can't change items";
        }
    }

}

/// <summary>
/// An item was updated. And Change that item's permission
/// </summary>
public override void ItemUpdated(SPItemEventProperties properties)
{
    base.ItemUpdated(properties);

    using (SPWeb web = properties.Web)
    {
        SPSecurity.RunWithElevatedPrivileges(delegate ()
        {
            //base.ItemAdded(properties);
            
              properties.ListItem.ResetRoleInheritance();
              properties.ListItem.BreakRoleInheritance(false, false);

              SPFieldUser userField = (SPFieldUser)getItem.Fields.GetField("PeoplePickerColumn");
              SPFieldUserValue userFieldValue = (SPFieldUserValue)userField.GetFieldValue(getItem["PeoplePickerColumn"].ToString());
              SPUser adminUser = userFieldValue.User;

              SPRoleDefinition roleDefinition = web.RoleDefinitions.GetByType(SPRoleType.Reader);
              SPRoleAssignment roleAssignment = new SPRoleAssignment(adminUser);
              roleAssignment.RoleDefinitionBindings.Add(roleDefinition);
              properties.ListItem.RoleAssignments.Add(roleAssignment);

              //Give perm to group
              SPGroup ygroup = web.SiteGroups["ContSPGroup"];
              SPRoleDefinition roleDefinitiongr = web.RoleDefinitions.GetByType(SPRoleType.Contributor);
              SPRoleAssignment roleAssignmentGroup = new SPRoleAssignment(ygroup);
              roleAssignmentGroup.RoleDefinitionBindings.Add(roleDefinitiongr);
              properties.ListItem.RoleAssignments.Add(roleAssignmentGroup);


              properties.ListItem.Update();
        });
    }

}
