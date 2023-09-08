using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Linq;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using MSC = Microsoft.SharePoint.Client;
using System.Net.Http;
using System.Security.Policy;
using Newtonsoft.Json;
using System.Data;
using DocumentFormat.OpenXml;
using System.Data.OleDb;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using static System.Net.Mime.MediaTypeNames;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.SharePoint.Client;
using System.Configuration;
using Newtonsoft.Json.Linq;

namespace KwanOneOffMultipleTalent.Application
{

    public static class sharepointOperation
    {


        public static AppConfiguration GetSharepointCredentials(string siteUrl)
        {
            AppConfiguration _AppConfiguration = new AppConfiguration();

            _AppConfiguration.ServiceSiteUrl = siteUrl;// _UserOperation.ReadValue("SP_Address");
            _AppConfiguration.ServiceUserName = ConfigurationManager.AppSettings["SP_USER_ID_Live"]; //_UserOperation.ReadValue("SP_USER_ID_Live");
            _AppConfiguration.ServicePassword = ConfigurationManager.AppSettings["SP_Password_Live"];// _UserOperation.ReadValue("SP_Password_Live");// Decrypt(_UserOperation.ReadValue("SP_Password_Live"));
            _AppConfiguration.SP_PanAPI = ConfigurationManager.AppSettings["SP_PanAPI"];
            _AppConfiguration.TalentMasterListName = ConfigurationManager.AppSettings["TalentMasterListName"];
            _AppConfiguration.BusinessCenterMasterListName = ConfigurationManager.AppSettings["BusinessCenterMasterListName"];
            _AppConfiguration.OneOffTalentListName = ConfigurationManager.AppSettings["OneOffTalentListName"];
            _AppConfiguration.OneOffMainListName = ConfigurationManager.AppSettings["OneOffMainListName"];
            _AppConfiguration.SubVerticalMasterListName = ConfigurationManager.AppSettings["SubVerticalMasterListName"];
            _AppConfiguration.Log_OneofTalentListName = ConfigurationManager.AppSettings["Log_OneofTalentListName"];
            _AppConfiguration.WorkFlowMasterListName = ConfigurationManager.AppSettings["WorkFlowMasterListName"];

            return _AppConfiguration;
        }

        public static MSC.ClientContext GetContext(string siteUrl)
        {
            try
            {
                AppConfiguration _AppConfiguration = GetSharepointCredentials(siteUrl);
                var securePassword = new SecureString();
                foreach (char c in _AppConfiguration.ServicePassword)
                {
                    securePassword.AppendChar(c);
                }
                var onlineCredentials = new MSC.SharePointOnlineCredentials(_AppConfiguration.ServiceUserName, securePassword);
                var context = new MSC.ClientContext(_AppConfiguration.ServiceSiteUrl);
                context.Credentials = onlineCredentials;
                return context;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public static async Task<List<ErrorApplication>> GetActiveErrorListAsync(string siteUrl, string listName)
        {
            List<ErrorApplication> errorModels = new List<ErrorApplication>();
            try
            {
                using (MSC.ClientContext context = sharepointOperation.GetContext(siteUrl))
                {
                    if (context != null)
                    {
                        AppConfiguration AppConfic = GetSharepointCredentials(siteUrl);
                        MSC.List list = context.Web.Lists.GetByTitle(listName);
                        MSC.ListItemCollectionPosition itemPosition = null;
                        DataTable dtExcel = new DataTable();
                        DataTable dtUniqueTalent = new DataTable();
                        DataTable dtVarifyTalentPan = new DataTable();
                        DataTable dtBusinessCenter = new DataTable();
                        DataTable dtOneofMain = new DataTable();
                        DataTable dtNotavailablePan = new DataTable();
                        DataTable dtValidateTalent = new DataTable();
                        DataTable AlldetailExcel = new DataTable();

                        dtValidateTalent.Columns.Add("TalentName", typeof(string));
                        dtValidateTalent.Columns.Add("PanNo", typeof(string));
                        dtValidateTalent.Columns.Add("PANName", typeof(string));

                        dtNotavailablePan.Columns.Add("Talent Name", typeof(string));
                        dtNotavailablePan.Columns.Add("PanNo", typeof(string));

                        var URLName = "";
                        string newJsonString = "";
                        string LogStatus = "Success";

                        MSC.CamlQuery camlQuery = new MSC.CamlQuery();
                        camlQuery.ListItemCollectionPosition = itemPosition;
                        camlQuery.ViewXml = @"<View><Query><Where><Eq><FieldRef Name='Status' /><Value Type='Text'>Pending</Value></Eq></Where></Query></View>";
                        MSC.ListItemCollection Items = list.GetItems(camlQuery);
                        context.Load(Items);
                        context.ExecuteQuery();
                        itemPosition = Items.ListItemCollectionPosition;

                        foreach (MSC.ListItem item in Items)
                        {
                            URLName = Convert.ToString(item["FileRef"]).Trim();
                            int LID = Convert.ToInt32(item["LID"]);
                            int ID = Convert.ToInt32(item["ID"]);
                            UpdateOneofTalentLibStatus(siteUrl, listName, ID);
                            dtExcel = loopexcel(siteUrl, URLName);
                            dtOneofMain = GetOneofMaindata(siteUrl, LID);
                            GetUniqueBusinessCenter(siteUrl, dtExcel, dtOneofMain);
                            AlldetailExcel = GetBusinessId(siteUrl, dtExcel, dtOneofMain);
                            dtUniqueTalent = RemoveDuplicatePan(AlldetailExcel);

                            for (int i = 0; i < dtUniqueTalent.Rows.Count; i++)
                            {
                                string apiUrl = AppConfic.SP_PanAPI + dtUniqueTalent.Rows[i]["Pan no."] + "";
                                using (HttpClient client = new HttpClient())
                                {
                                    HttpResponseMessage response = client.GetAsync(apiUrl).Result;

                                    if (response.IsSuccessStatusCode)
                                    {
                                        string responseBody = await response.Content.ReadAsStringAsync();
                                        dynamic jsonObject = JsonConvert.DeserializeObject(responseBody);
                                        JObject jsObject = JObject.Parse(jsonObject);
                                        string RequestStatus = "success"; //commented below line because Pan checking api is exhausted that's why giving status value                                        
                                        //string RequestStatus = (string)jsObject["status"]; //uncomment after
                                        if (RequestStatus == "success")
                                        {
                                            //JObject dataObject = (JObject)jsObject["data"]; //uncomment after
                                            //string name = (string)dataObject["name"]; //uncomment after

                                            //string name = dtUniqueTalent.Rows[i]["name"].ToString(); //uncomment after

                                            string TalentName = dtUniqueTalent.Rows[i]["Talent Name"].ToString(); //commented below line because Pan checking api is exhausted that's why giving status value                                        

                                            DataRow sourceRow = dtUniqueTalent.Rows[i];
                                            DataRow newRow = dtValidateTalent.NewRow();

                                            // Copy data from source row to new row
                                            newRow["TalentName"] = sourceRow["Talent Name"];
                                            newRow["PanNo"] = sourceRow["Pan No."];
                                            //newRow["PANName"] = name; //uncomment after

                                            newRow["PANName"] = TalentName; //commented below line because Pan checking api is exhausted that's why giving status value

                                            // Add the new row to the destination table
                                            dtValidateTalent.Rows.Add(newRow);
                                        }
                                        else
                                        {
                                            JObject newJsonObject = new JObject();
                                            newJsonObject["Talent Name"] = dtUniqueTalent.Rows[i]["Talent Name"].ToString();
                                            newJsonObject["PAN Number"] = dtUniqueTalent.Rows[i]["Pan no."].ToString();
                                            newJsonString = newJsonObject.ToString();
                                            LogStatus = "Error";

                                            DataRow sourceRow = dtUniqueTalent.Rows[i];
                                            DataRow newRow = dtNotavailablePan.NewRow();
                                            newRow["Talent Name"] = sourceRow["Talent Name"];
                                            newRow["PanNo"] = sourceRow["Pan No."];
                                            dtNotavailablePan.Rows.Add(newRow);
                                        }
                                    }
                                }
                            }

                            DataTable dtOneTalentRecords = RemoveOneofTalent(dtNotavailablePan, AlldetailExcel);
                            InsertLogTalent(siteUrl, LID, LogStatus, newJsonString);
                            SaveNewTalent(siteUrl, dtValidateTalent);
                            SaveOneOfTalent(siteUrl, dtOneTalentRecords, LID);
                            UpdateOneOfMainStatus(siteUrl, dtOneTalentRecords, LID);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return errorModels;
        }

        public static void UpdateOneofTalentLibStatus(string siteUrl, string libraryName, int ID)
        {
            try
            {
                AppConfiguration _AppConfiguration = GetSharepointCredentials(siteUrl);
                using (var context = new ClientContext(siteUrl))
                {
                    SecureString securePassword = new SecureString();
                    foreach (char c in _AppConfiguration.ServicePassword)
                    {
                        securePassword.AppendChar(c);
                    }
                    context.Credentials = new SharePointOnlineCredentials(_AppConfiguration.ServiceUserName, securePassword);
                    List list = context.Web.Lists.GetByTitle(libraryName);

                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                    ListItem newItem = list.AddItem(itemCreateInfo);

                    List targetList = context.Web.Lists.GetByTitle(libraryName);
                    ListItem targetItem = targetList.GetItemById(ID);

                    context.Load(targetItem);
                    context.ExecuteQuery();

                    // Update the value of a specific column
                    //targetItem["ID"] = ID; // Replace "StatusColumn" with the actual column name
                    targetItem["Status"] = "Inprogress";
                    targetItem.Update();

                    context.ExecuteQuery();

                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public static DataTable loopexcel(string siteUrl, string URL)
        {
            using (MSC.ClientContext context = sharepointOperation.GetContext(siteUrl))
            {

                var fileimage = context.Web.GetFileByServerRelativeUrl(URL);

                context.Load(fileimage);
                context.ExecuteQuery();
                ClientResult<System.IO.Stream> data = fileimage.OpenBinaryStream();
                context.Load(fileimage);
                context.ExecuteQuery();
                using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                {
                    data.Value.CopyTo(mStream);
                    var SheetName = "Sheet1";
                    DataTable dt = new DataTable();
                    using (SpreadsheetDocument document = SpreadsheetDocument.Open(mStream, false))
                    {
                        WorkbookPart workbookPart = document.WorkbookPart;
                        Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == SheetName);

                        WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

                        // Get the SharedStringTable from the SharedStringTablePart of the worksheet part.
                        SharedStringTable sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable;

                        // Get all rows from the worksheet.
                        IEnumerable<Row> rows = worksheetPart.Worksheet.Descendants<Row>();

                        // Assuming you have three columns in the Excel sheet, modify as needed.
                        dt.Columns.Add("Talent Name", typeof(string));
                        dt.Columns.Add("Pan No.", typeof(string));
                        dt.Columns.Add("GST No", typeof(string));
                        dt.Columns.Add("Business Center", typeof(string));
                        dt.Columns.Add("External Cost", typeof(string));
                        dt.Columns.Add("Internal Cost", typeof(string));
                        dt.Columns.Add("GMV", typeof(string));
                        dt.Columns.Add("Margin", typeof(string));
                        dt.Columns.Add("Margin%", typeof(string));

                        int rowCount = 0;

                        foreach (Row row in rows) // This will iterate through each row in the worksheet.
                        {
                            if (rowCount != 0) // Skip the header row.
                            {
                                DataRow tempRow = dt.NewRow();

                                // Modify this part according to the range of columns you want to read.
                                for (int i = 0; i < 9; i++) // Assuming you want to read the first three columns.
                                {
                                    // Get the cell at the specified column index from the current row.
                                    Cell cell = row.Elements<Cell>().ElementAt(i);

                                    // Get the cell value using the GetCellValue method and pass the sharedStringTable.
                                    string cellValue = GetCellValue(cell, sharedStringTable);

                                    // Set the cell value in the corresponding column of the DataRow.
                                    tempRow[i] = cellValue;
                                }

                                // Add the populated DataRow to the DataTable.
                                dt.Rows.Add(tempRow);
                            }

                            rowCount++;
                        }

                    }
                    return dt;
                }

            }

        }

        public static DataTable GetTalentMasterdata(string siteUrl)
        {
            try
            {
                AppConfiguration _AppConfiguration = GetSharepointCredentials(siteUrl);
                using (var context = new ClientContext(siteUrl))
                {
                    SecureString securePassword = new SecureString();
                    foreach (char c in _AppConfiguration.ServicePassword)
                    {
                        securePassword.AppendChar(c);
                    }
                    context.Credentials = new SharePointOnlineCredentials(_AppConfiguration.ServiceUserName, securePassword);

                    // Load the list
                    List list = context.Web.Lists.GetByTitle(_AppConfiguration.TalentMasterListName);

                    // Define a CamlQuery to retrieve all items from the list
                    CamlQuery query = CamlQuery.CreateAllItemsQuery();

                    // Execute the query
                    ListItemCollection items = list.GetItems(query);
                    context.Load(items);
                    context.ExecuteQuery();
                    DataTable dataTable = new DataTable();
                    dataTable.Columns.Add("ID", typeof(string));
                    dataTable.Columns.Add("TalentName", typeof(string));
                    dataTable.Columns.Add("PANNumber", typeof(string));
                    foreach (ListItem item in items)
                    {
                        DataRow newRow = dataTable.NewRow();
                        newRow["ID"] = item["ID"];
                        newRow["TalentName"] = item["TalentName"];
                        newRow["PANNumber"] = item["PANNumber"];
                        dataTable.Rows.Add(newRow);
                    }
                    return dataTable;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static DataTable GetOneofMaindata(string siteUrl, int LID)
        {
            try
            {
                DataTable dataTable = new DataTable();
                dataTable.Columns.Add("SubVerticalName", typeof(string));
                dataTable.Columns.Add("FXRate", typeof(string));

                AppConfiguration _AppConfiguration = GetSharepointCredentials(siteUrl);
                using (ClientContext context = new ClientContext(siteUrl))
                {

                    SecureString securePassword = new SecureString();
                    foreach (char c in _AppConfiguration.ServicePassword)
                    {
                        securePassword.AppendChar(c);
                    }
                    context.Credentials = new SharePointOnlineCredentials(_AppConfiguration.ServiceUserName, securePassword);

                    List targetList = context.Web.Lists.GetByTitle(_AppConfiguration.OneOffMainListName);

                    CamlQuery query = new CamlQuery();
                    query.ViewXml = $"<View><Query><Where><Eq><FieldRef Name=ID /><Value Type='Number'>{LID}</Value></Eq></Where></Query></View>";

                    ListItemCollection items = targetList.GetItems(query);
                    context.Load(items);
                    context.ExecuteQuery();

                    foreach (ListItem item in items)
                    {
                        DataRow newRow = dataTable.NewRow();
                        newRow["SubVerticalName"] = Convert.ToString(item["SubVerticalName"]);
                        newRow["FXRate"] = Convert.ToString(item["FXRate"]);
                        dataTable.Rows.Add(newRow);
                    }
                }
                int GetSubVerticalId = GetSubVerticalName(siteUrl, dataTable);
                dataTable.Columns.Add("SubVerticalNameId", typeof(string));
                dataTable.Rows[0]["SubVerticalNameId"] = GetSubVerticalId;

                return dataTable;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static int GetSubVerticalName(string siteUrl, DataTable dt)
        {
            int SubVerticalNameId = 0;
            AppConfiguration _AppConfiguration = GetSharepointCredentials(siteUrl);
            string SubVerticalName = dt.Rows[0]["SubVerticalName"].ToString();

            using (ClientContext context = new ClientContext(siteUrl))
            {
                SecureString securePassword = new SecureString();
                foreach (char c in _AppConfiguration.ServicePassword)
                {
                    securePassword.AppendChar(c);
                }
                context.Credentials = new SharePointOnlineCredentials(_AppConfiguration.ServiceUserName, securePassword);

                List targetList = context.Web.Lists.GetByTitle(_AppConfiguration.SubVerticalMasterListName);

                CamlQuery query = new CamlQuery();
                query.ViewXml = $"<View><Query><Where><Eq><FieldRef Name='SubVerticalName'/><Value Type='Text'>{SubVerticalName}</Value></Eq></Where></Query></View>";

                ListItemCollection items = targetList.GetItems(query);
                context.Load(items);
                context.ExecuteQuery();

                foreach (ListItem item in items)
                {
                    SubVerticalNameId = Convert.ToInt32(item["ID"]);
                }
            }
            return SubVerticalNameId;
        }
        public static DataTable RemoveDuplicatePan(DataTable dt2)
        {
            try
            {
                string columnName = "Pan No.";
                DataTable deduplicatedTable = dt2.Clone();
                var uniqueRows = dt2.AsEnumerable()
               .GroupBy(row => row.Field<object>(columnName)) // Group by the specified column
               .Select(group => group.First()); // Select the first row from each group

                foreach (DataRow row in uniqueRows)
                {
                    deduplicatedTable.ImportRow(row); // Import the selected rows into the deduplicated DataTable
                }
                return deduplicatedTable;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public static void InsertLogTalent(string siteUrl, int LID, string LogSatus, string LogmsgJsonString)
        {

            AppConfiguration AppConfic = GetSharepointCredentials(siteUrl);
            using (var context = new ClientContext(siteUrl))
            {
                SecureString securePassword = new SecureString();
                foreach (char c in AppConfic.ServicePassword)
                {
                    securePassword.AppendChar(c);
                }
                context.Credentials = new SharePointOnlineCredentials(AppConfic.ServiceUserName, securePassword);

                List list = context.Web.Lists.GetByTitle(AppConfic.Log_OneofTalentListName);

                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem newItem = list.AddItem(itemCreateInfo);

                newItem["LID"] = LID;
                newItem["Status"] = LogSatus;
                newItem["ErrorLog"] = LogmsgJsonString;

                newItem.Update();
                context.ExecuteQuery();
            }
        }

        public static DataTable RemoveOneofTalent(DataTable dtNotavailablePan, DataTable dtExcel)
        {
            DataTable dtOneofTalent = new DataTable();
            try
            {
                var PanNoToRemove = dtNotavailablePan.AsEnumerable().Select(row => row.Field<string>("PanNo")).ToList();

                for (int i = dtExcel.Rows.Count - 1; i >= 0; i--)
                {
                    string PANNO = dtExcel.Rows[i]["Pan No."].ToString();
                    if (PanNoToRemove.Contains(PANNO))
                    {
                        dtExcel.Rows.RemoveAt(i);
                    }
                }

                return dtExcel;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public static void GetUniqueBusinessCenter(string siteUrl, DataTable dtTalentPayout, DataTable dtOneofMain)
        {
            try
            {
                string columnName = "Business Center";
                DataTable dtUnqBusinessCenter = new DataTable();
                dtUnqBusinessCenter.Columns.Add("Business Center", typeof(string));
                var uniqueRows = dtTalentPayout.AsEnumerable()
               .GroupBy(row => row.Field<object>(columnName))
               .Select(group => group.First());
                foreach (DataRow row in uniqueRows)
                {
                    string BusinessCenter = row["Business Center"].ToString();
                    DataRow newRow = dtUnqBusinessCenter.NewRow();
                    newRow["Business Center"] = BusinessCenter; ;
                    dtUnqBusinessCenter.Rows.Add(newRow);
                }
                DataTable dtNewBusinessCenter = GetBusinessCenterdata(siteUrl, dtUnqBusinessCenter);
                InsertBusinessListItem(siteUrl, dtNewBusinessCenter, dtOneofMain);
            }
            catch (Exception)
            {
                throw;
            }

        }

        public static DataTable GetBusinessId(string siteUrl, DataTable dtExcel, DataTable dtOneofMain)
        {
            try
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("BusinessCenterName", typeof(string));
                dt.Columns.Add("ID", typeof(string));

                dtExcel.Columns.Add("BusinessCenterID", typeof(int));
                dtExcel.Columns.Add("FXRate", typeof(int));

                ListItemCollection items = GetBusinessCenterListItem(siteUrl);

                foreach (ListItem item in items)
                {
                    DataRow newRow = dt.NewRow();
                    newRow["BusinessCenterName"] = item["BusinessCenterName"].ToString();
                    newRow["ID"] = item["ID"];
                    dt.Rows.Add(newRow);
                }

                foreach (DataRow row in dtExcel.Rows)
                {
                    string name = (string)row["Business Center"];
                    var searchRow = dt.AsEnumerable().FirstOrDefault(dtrow => dtrow.Field<string>("BusinessCenterName") == name);

                    row["BusinessCenterID"] = searchRow[1];
                    row["FXRate"] = dtOneofMain.Rows[0]["FXRate"];

                }

                return dtExcel;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public static void SaveNewTalent(string siteUrl, DataTable NewTalent)
        {
            try
            {
                DataTable dtExitsTalent = GetTalentMasterdata(siteUrl);

                DataTable dt = new DataTable();
                dt.Columns.Add("Talent Name", typeof(string));
                dt.Columns.Add("Pan No", typeof(string));
                dt.Columns.Add("PANName", typeof(string));


                for (int i = 0; i < NewTalent.Rows.Count; i++)
                {
                    string TalentNewName = NewTalent.Rows[i]["TalentName"].ToString();
                    string PanNewNo = NewTalent.Rows[i]["PanNo"].ToString();
                    string PANNewName = NewTalent.Rows[i]["PANName"].ToString();
                    bool valueExists = dtExitsTalent.AsEnumerable().Any(row => row.Field<string>("PANNumber") == PanNewNo);

                    if (valueExists == false)
                    {

                        DataRow newRow = dt.NewRow();
                        newRow["Talent Name"] = TalentNewName;
                        newRow["Pan No"] = PanNewNo;
                        newRow["PANName"] = PANNewName;
                        dt.Rows.Add(newRow);
                    }

                }
                InsertListItem(siteUrl, dt);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public static ListItemCollection GetBusinessCenterListItem(string siteUrl)
        {

            AppConfiguration _AppConfiguration = GetSharepointCredentials(siteUrl);
            using (var context = new ClientContext(siteUrl))
            {
                SecureString securePassword = new SecureString();
                foreach (char c in _AppConfiguration.ServicePassword)
                {
                    securePassword.AppendChar(c);
                }
                context.Credentials = new SharePointOnlineCredentials(_AppConfiguration.ServiceUserName, securePassword);

                // Load the list
                List list = context.Web.Lists.GetByTitle(_AppConfiguration.BusinessCenterMasterListName);

                // Define a CamlQuery to retrieve all items from the list
                CamlQuery query = CamlQuery.CreateAllItemsQuery();

                // Execute the query
                ListItemCollection items = list.GetItems(query);
                context.Load(items);
                context.ExecuteQuery();
                return items;
            }
        }
        public static DataTable GetBusinessCenterdata(string siteUrl, DataTable dtUnqBusinessCenter)
        {
            try
            {
                ListItemCollection items = GetBusinessCenterListItem(siteUrl);

                DataTable dtNewBusinessCenter = new DataTable();
                dtNewBusinessCenter.Columns.Add("Business Center", typeof(string));
                DataTable dataTable = new DataTable();
                dataTable.Columns.Add("BusinessCenterName", typeof(string));

                foreach (ListItem item in items)
                {
                    DataRow newRow = dataTable.NewRow();
                    newRow["BusinessCenterName"] = item["BusinessCenterName"];
                    dataTable.Rows.Add(newRow);
                }


                for (int i = 0; i < dtUnqBusinessCenter.Rows.Count; i++)
                {

                    string BusinessCenterNewName = dtUnqBusinessCenter.Rows[i]["Business Center"].ToString();
                    bool valueExists = dataTable.AsEnumerable().Any(row => row.Field<string>("BusinessCenterName") == BusinessCenterNewName);

                    if (valueExists == false)
                    {

                        DataRow newRow = dtNewBusinessCenter.NewRow();
                        newRow["Business Center"] = BusinessCenterNewName;

                        dtNewBusinessCenter.Rows.Add(newRow);
                    }
                }


                return dtNewBusinessCenter;

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void InsertBusinessListItem(string siteUrl, DataTable dtNewTalent, DataTable dtOneofMain)
        {
            try
            {
                AppConfiguration AppConfic = GetSharepointCredentials(siteUrl);
                for (int i = 0; i < dtNewTalent.Rows.Count; i++)
                {
                    using (var context = new ClientContext(siteUrl))
                    {
                        SecureString securePassword = new SecureString();
                        foreach (char c in AppConfic.ServicePassword)
                        {
                            securePassword.AppendChar(c);
                        }
                        context.Credentials = new SharePointOnlineCredentials(AppConfic.ServiceUserName, securePassword);

                        List list = context.Web.Lists.GetByTitle(AppConfic.BusinessCenterMasterListName);

                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                        ListItem newItem = list.AddItem(itemCreateInfo);

                        string BusinessCenterName = dtNewTalent.Rows[i]["Business Center"].ToString();

                        newItem["BusinessCenterName"] = BusinessCenterName;
                        newItem["SubVerticalName"] = dtOneofMain.Rows[0]["SubVerticalNameId"];

                        newItem.Update();
                        context.ExecuteQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static void InsertListItem(string siteUrl, DataTable dtNewTalent)
        {
            try
            {
                AppConfiguration _AppConfiguration = GetSharepointCredentials(siteUrl);
                for (int i = 0; i < dtNewTalent.Rows.Count; i++)
                {
                    using (var context = new ClientContext(siteUrl))
                    {
                        SecureString securePassword = new SecureString();
                        foreach (char c in _AppConfiguration.ServicePassword)
                        {
                            securePassword.AppendChar(c);
                        }
                        context.Credentials = new SharePointOnlineCredentials(_AppConfiguration.ServiceUserName, securePassword);

                        List list = context.Web.Lists.GetByTitle(_AppConfiguration.TalentMasterListName);

                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                        ListItem newItem = list.AddItem(itemCreateInfo);

                        string TalentNewName = dtNewTalent.Rows[i]["Talent Name"].ToString();
                        string PanNewName = dtNewTalent.Rows[i]["Pan No"].ToString();
                        newItem["TalentName"] = TalentNewName;
                        newItem["PANNumber"] = PanNewName;
                        newItem["PANName"] = dtNewTalent.Rows[i]["PANName"].ToString();
                        newItem.Update();

                        context.ExecuteQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static string GetCellValue(Cell cell, SharedStringTable sharedStringTable)
        {

            if (cell == null || cell.CellValue == null)
            {
                return string.Empty;
            }

            string cellValue = cell.CellValue.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                // If the cell contains a shared string, fetch the value from the shared string table.
                int index = int.Parse(cellValue);
                cellValue = sharedStringTable.ElementAt(index).InnerText;
            }

            return cellValue;
        }

        public static void SaveOneOfTalent(string siteUrl, DataTable dtExcel, int LID)
        {
            try
            {
                DataTable dtTalentMst = GetTalentMasterdata(siteUrl);
                dtExcel.Columns.Add("TalentId", typeof(string));

                for (int i = 0; i < dtExcel.Rows.Count; i++)
                {
                    string PanNumber = dtExcel.Rows[i]["Pan No."].ToString();
                    var searchRow = dtTalentMst.AsEnumerable().FirstOrDefault(row => row.Field<string>("PANNumber") == PanNumber);

                    if (searchRow != null)
                    {
                        string valueAtIndex0 = Convert.ToString(searchRow.ItemArray[0]);
                        dtExcel.Rows[i]["TalentId"] = valueAtIndex0;
                    }


                }
                InsertOneOffTalent(siteUrl, dtExcel, LID);
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public static void InsertOneOffTalent(string siteUrl, DataTable dtExcel, int LID)
        {
            try
            {
                AppConfiguration _AppConfiguration = GetSharepointCredentials(siteUrl);
                for (int i = 0; i < dtExcel.Rows.Count; i++)
                {
                    using (var context = new ClientContext(siteUrl))
                    {
                        SecureString securePassword = new SecureString();
                        foreach (char c in _AppConfiguration.ServicePassword)
                        {
                            securePassword.AppendChar(c);
                        }
                        context.Credentials = new SharePointOnlineCredentials(_AppConfiguration.ServiceUserName, securePassword);

                        List list = context.Web.Lists.GetByTitle(_AppConfiguration.OneOffTalentListName);

                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                        ListItem newItem = list.AddItem(itemCreateInfo);
                        if (dtExcel.Rows[i]["TalentId"] != DBNull.Value && dtExcel.Rows[i]["TalentId"] != null)
                        {
                            int TalentId = Convert.ToInt32(dtExcel.Rows[i]["TalentId"]);
                            FieldLookupValue lookupValue = new FieldLookupValue();
                            lookupValue.LookupId = TalentId;
                            double FXRate = Convert.ToDouble(dtExcel.Rows[i]["FXRate"]);
                            double ExternalCost = Convert.ToDouble(dtExcel.Rows[i]["External Cost"]);
                            double InternalCost = Convert.ToDouble(dtExcel.Rows[i]["Internal Cost"]);
                            double GMV = Convert.ToDouble(dtExcel.Rows[i]["GMV"]);
                            double Margin = Convert.ToDouble(dtExcel.Rows[i]["Margin"]);

                            double FXExternalCost = FXRate * ExternalCost;
                            double FXInternalCost = FXRate * InternalCost;
                            double FXGMV = FXRate * GMV;
                            double FXMargin = FXRate * Margin;

                            newItem["TalentName"] = TalentId;
                            newItem["BusinessCenterName"] = Convert.ToInt32(dtExcel.Rows[i]["BusinessCenterID"]);
                            newItem["PayoutPer"] = dtExcel.Rows[i]["Margin%"];
                            newItem["ExternalAmount"] = dtExcel.Rows[i]["External Cost"];
                            newItem["InternalAmount"] = dtExcel.Rows[i]["Internal Cost"];
                            newItem["GMV"] = dtExcel.Rows[i]["GMV"];
                            newItem["Margin"] = dtExcel.Rows[i]["Margin"];
                            newItem["OneOffID"] = LID;
                            newItem["FXRate"] = FXRate;
                            newItem["FXExternalAmount"] = FXExternalCost;
                            newItem["FXInternalAmount"] = FXInternalCost;
                            newItem["FXMargin"] = FXMargin;
                            newItem["FXGMV"] = FXGMV;
                            newItem.Update();

                            context.ExecuteQuery();
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        public static void UpdateOneOfMainStatus(string siteUrl, DataTable dtOneTalentRecords, int OneMainID)
        {

            try
            {
                DataTable Statusdt = ReturnStatusDetails(siteUrl);

                AppConfiguration _AppConfiguration = GetSharepointCredentials(siteUrl);
                using (var context = new ClientContext(siteUrl))
                {
                    SecureString securePassword = new SecureString();
                    foreach (char c in _AppConfiguration.ServicePassword)
                    {
                        securePassword.AppendChar(c);
                    }
                    context.Credentials = new SharePointOnlineCredentials(_AppConfiguration.ServiceUserName, securePassword);
                    List list = context.Web.Lists.GetByTitle(_AppConfiguration.OneOffTalentListName);

                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                    ListItem newItem = list.AddItem(itemCreateInfo);

                    List targetList = context.Web.Lists.GetByTitle(_AppConfiguration.OneOffMainListName);
                    ListItem targetItem = targetList.GetItemById(OneMainID);

                    context.Load(targetItem);
                    context.ExecuteQuery();

                    double ExternalCost = 0;
                    double InternalCost = 0;
                    double GMV = 0;
                    double Margin = 0;
                    double FXRate = 0;

                    foreach (DataRow row in dtOneTalentRecords.Rows)
                    {
                        ExternalCost += Convert.ToDouble(row["External Cost"]);
                        InternalCost += Convert.ToDouble(row["Internal Cost"]);
                        GMV += Convert.ToDouble(row["GMV"]);
                        Margin += Convert.ToDouble(row["Margin"]);
                        FXRate = Convert.ToDouble(row["FXRate"]);
                    }
                    double FXExternalCost = FXRate * ExternalCost;
                    double FXInternalCost = FXRate * InternalCost;
                    double FXGMV = FXRate * GMV;
                    double FXMargin = FXRate * Margin;

                    targetItem["TalentExternalAmount"] = ExternalCost;
                    targetItem["TalentInternalAmount"] = InternalCost;
                    targetItem["TotalMargin"] = Margin;
                    targetItem["GMV"] = GMV;
                    targetItem["FXTalentExternalAmount"] = FXExternalCost;
                    targetItem["FXTalentInternalAmount"] = FXInternalCost;
                    targetItem["FXGMV"] = FXGMV;
                    targetItem["FXTotalMargin"] = FXMargin;
                    targetItem["StatusName"] = Statusdt.Rows[0]["StatusId"];
                    targetItem["InternalStatus"] = Statusdt.Rows[0]["StatusName"];
                    targetItem.Update();

                    context.ExecuteQuery();

                }
            }
            catch (Exception ex)
            {

                throw ex;
            }

        }

        public static DataTable ReturnStatusDetails(string siteUrl)
        {
            int SatusId = 0;
            string InternalStatus = "";
            DataTable dtStatusdtls = new DataTable();
            dtStatusdtls.Columns.Add("StatusId", typeof(string));
            dtStatusdtls.Columns.Add("StatusName", typeof(string));
            AppConfiguration _AppConfiguration = GetSharepointCredentials(siteUrl);

            try
            {
                using (ClientContext context = new ClientContext(siteUrl))
                {

                    SecureString securePassword = new SecureString();
                    foreach (char c in _AppConfiguration.ServicePassword)
                    {
                        securePassword.AppendChar(c);
                    }
                    context.Credentials = new SharePointOnlineCredentials(_AppConfiguration.ServiceUserName, securePassword);

                    List targetList = context.Web.Lists.GetByTitle(_AppConfiguration.WorkFlowMasterListName);
                    
                    CamlQuery query = new CamlQuery();
                    //query.ViewXml = $"<View><Query><Where><And><Eq><FieldRef Name='FromStatus/InternalStatus' /><Value Type='Text'>{fromstatus}</Value></Eq><Eq><FieldRef Name='ToStatus/InternalStatus' /><Value Type='Text'>{"PendingForHODApproval"}</Value></Eq></And></Where></Query></View>";
                    query.ViewXml = $"<View><Query><Where><And><Eq><FieldRef Name='FromStatus' LookupId='TRUE' /><Value Type='Lookup'>10</Value></Eq><Eq><FieldRef Name='ToStatus' LookupId='TRUE' /><Value Type='Lookup'>1</Value></Eq></And></Where></Query></View>";

                    ListItemCollection items = targetList.GetItems(query);
                    context.Load(items);
                    context.ExecuteQuery();

                    foreach (ListItem item in items)
                    {
                        FieldLookupValue lookupValue = item["InternalStatus"] as FieldLookupValue;
                        SatusId = lookupValue.LookupId;
                        InternalStatus = lookupValue.LookupValue;

                        DataRow newRow = dtStatusdtls.NewRow();
                        newRow["StatusId"] = lookupValue.LookupId;
                        newRow["StatusName"] = lookupValue.LookupValue;
                        dtStatusdtls.Rows.Add(newRow);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return dtStatusdtls;
        }

    }
    public class AppConfiguration
    {
        public string ServiceSiteUrl;
        public string ServiceUserName;
        public string ServicePassword;
        public string SP_PanAPI;
        public string TalentMasterListName;
        public string BusinessCenterMasterListName;
        public string OneOffTalentListName;
        public string OneOffMainListName;
        public string SubVerticalMasterListName;
        public string Log_OneofTalentListName;
        public string WorkFlowMasterListName;

    }
}

