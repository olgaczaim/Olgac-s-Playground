SPList list = null;
SPListItem item;

SPSite thissite = SPContext.Current.Site;
SPWeb thisweb = SPContext.Current.Web;

SPUtility.ValidateFormDigest();
SPSecurity.RunWithElevatedPrivileges(delegate ()
{
    using (SPSite currentSite = new SPSite(thissite.ID))
    {
        using (SPWeb currentWeb = currentSite.OpenWeb(thisweb.ID))
        {
            
            list = currentWeb.Lists.TryGetList("TestDocLib");

            string batchDataFormat = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><ows:Batch OnError=\"Continue\">{0}</ows:Batch>";
            string batchDataSetVar = "<SetVar Name=\"urn:schemas-microsoft-com:office:office#{0}\">{1}</SetVar>";
            string batchDataUpdateMethodFormat = "<Method ID=\"{0}\"><SetList>{1}</SetList><SetVar Name=\"ID\">{2}</SetVar><SetVar Name=\"Cmd\">Save</SetVar><SetVar Name='owsfileref'>{3}</SetVar>{4}</Method>";

            StringBuilder sbBatchDataMethod = new StringBuilder();


            string batchDataSetVarLines = string.Format(batchDataSetVar, "Stok", 5);
            sbBatchDataMethod.AppendFormat(batchDataUpdateMethodFormat, 13, list.ID, 13, currentWeb.Url + "/" + list.Title + "/image1.JPG", batchDataSetVarLines);

           
            string batchDataXml = string.Format(batchDataFormat, sbBatchDataMethod.ToString());
            string result = currentWeb.ProcessBatchData(batchDataXml);



            //currentWeb.AllowUnsafeUpdates = false;
        }
    }
});

//if you need to update sharepoint list (not document library) then remove <SetVar Name='owsfileref'>{3}</SetVar> and " currentWeb.Url + "/" + list.Title + "/image1.JPG","
//also run batchDataSetVarLines and sbBatchDataMethod inside foreach loop
