
//Select from a list
SPSite thissite = SPContext.Current.Site;
SPWeb thisweb = SPContext.Current.Web;
//SPSecurity.RunWithElevatedPrivileges(delegate()
//{
using (SPSite currentSite = new SPSite(thissite.ID))
{
    using (SPWeb currentWeb = currentSite.OpenWeb(thisweb.ID))
    {

        SPList list = currentWeb.Lists["ListName"];

        IEnumerable<SPListItem> items = (from SPListItem a in list.Items
                                        where a["Title"].Equals("1")
                                        orderby a["Title"] descending
                                        select a);


        foreach (SPListItem item in items)
        {
            string title = item["Title"].ToString();
        }

    }
}
//}



//Update list item by id
SPSecurity.RunWithElevatedPrivileges(delegate ()
{
   using (SPSite currentSite = new SPSite(thissite.ID))
   {
       using (SPWeb currentWeb = currentSite.OpenWeb(thisweb.ID))
       {

           SPList formList = currentWeb.Lists["DummyList"];
           currentWeb.AllowUnsafeUpdates = true;

            SPListItem item = formList.GetItemById(1);
            item["Title"] = "Test";

            item.Update();


        }
    }
});



//Delete List Item
SPSecurity.RunWithElevatedPrivileges(delegate ()
{
    using (SPSite currentSite = new SPSite(thissite.ID))
    {
        using (SPWeb currentWeb = currentSite.OpenWeb(thisweb.ID))
        {

            SPList formList = currentWeb.Lists["DummyList"];
            currentWeb.AllowUnsafeUpdates = true;
            
            SPListItem item = formList.GetItemById(1);
            item.Delete();

        }
    }
});

//Add to List
SPSecurity.RunWithElevatedPrivileges(delegate ()
{
    using (SPSite currentSite = new SPSite(thissite.ID))
    {
        using (SPWeb currentWeb = currentSite.OpenWeb(thisweb.ID))
        {
            SPList formList = currentWeb.Lists["DummyList"];
            
            SPListItem newItem = formList.AddItem();
            newItem["Title"] = "Test Title Content";
            newItem.Update();
        }
    }
});
