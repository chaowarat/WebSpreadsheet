﻿using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;
using System.Web.Security;
using System.Linq;
using System.IO;
using System.Xml.Linq;
using System.Xml.XPath;
using System.Data.OleDb;
using System.Data;
using Excel;

namespace Spreadsheet
{
    public partial class BenefitAdminCode : System.Web.UI.Page
    {
        BenefitAdminDataContext bfAdmin = new BenefitAdminDataContext();
        private static string fileName = "code.html";
        private static DataSet resultFromUpload = null;

        protected void Page_Load(object sender, EventArgs e)
        {
            //if (!Request.IsAuthenticated || !User.Identity.Name.Equals("BenefitAdmin"))
            //{
            //    Response.Redirect("Casemanager.aspx");
            //}
            //else
            //{
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                //Label_user.Text = "Welcome " + Session["user"];
            //}
        }

        [System.Web.Services.WebMethod]
        public static string getHTML()
        {
            HtmlManager.Clean();
            return HtmlManager.createHtml(fileName);
        }

        [System.Web.Services.WebMethod]
        public static string SaveData(string data)
        {
            data = data.Replace("&", "&amp;");
            MemoryStream mem = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(data));
            XPathDocument document;
            try
            {
                document = new XPathDocument(mem);
            }
            catch (Exception)
            {
                return "error";
            }
            BenefitAdminDataContext bfAdmin = new BenefitAdminDataContext();
            var allInSvr = from c in bfAdmin.Services select c;
            bfAdmin.Services.DeleteAllOnSubmit(allInSvr);
            try
            {
                bfAdmin.SubmitChanges();
            }
            catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }

            var allInAct = from c in bfAdmin.Activities select c;
            bfAdmin.Activities.DeleteAllOnSubmit(allInAct);
            try
            {
                bfAdmin.SubmitChanges();
            }
            catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }

            var allInSAct = from c in bfAdmin.SubActivities select c;
            bfAdmin.SubActivities.DeleteAllOnSubmit(allInSAct);
            try
            {
                bfAdmin.SubmitChanges();
            }
            catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }

            var allInMat = from c in bfAdmin.Materials select c;
            bfAdmin.Materials.DeleteAllOnSubmit(allInMat);
            try
            {
                bfAdmin.SubmitChanges();
            }
            catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }

            XPathNavigator navigator = document.CreateNavigator();
            XPathNodeIterator allDocs = navigator.Select("DOCUMENTS/DOCUMENT");
            while (allDocs.MoveNext())
            {
                string title = allDocs.Current.GetAttribute("title", "");
                if (title.Equals("Service"))
                {
                    #region
                    string col0 = "", col1 = "", col2 = "", col3 = "", col4 = "", col5 = "";
                    string col6 = "", col7 = "", col8 = "", col9 = "", col10 = "";
                    XPathNodeIterator countRowString = allDocs.Current.Select("METADATA/ROWS");
                    countRowString.MoveNext();
                    int countRow = Convert.ToInt32(countRowString.Current.Value);

                    XPathNodeIterator rows = allDocs.Current.Select("DATA");
                    while (rows.MoveNext())
                    {
                        for (int i = 1; i < countRow; i++)
                        {
                            XPathNodeIterator getCol0 = rows.Current.Select("R" + i + "/C0");
                            getCol0.MoveNext();
                            col0 = getCol0.Current.Value;
                            XPathNodeIterator getCol1 = rows.Current.Select("R" + i + "/C1");
                            getCol1.MoveNext();
                            col1 = getCol1.Current.Value;
                            XPathNodeIterator getCol2 = rows.Current.Select("R" + i + "/C2");
                            getCol2.MoveNext();
                            col2 = getCol2.Current.Value;
                            XPathNodeIterator getCol3 = rows.Current.Select("R" + i + "/C3");
                            getCol3.MoveNext();
                            col3 = getCol3.Current.Value;
                            XPathNodeIterator getCol4 = rows.Current.Select("R" + i + "/C4");
                            getCol4.MoveNext();
                            col4 = getCol4.Current.Value;
                            XPathNodeIterator getCol5 = rows.Current.Select("R" + i + "/C5");
                            getCol5.MoveNext();
                            col5 = getCol5.Current.Value;
                            XPathNodeIterator getCol6 = rows.Current.Select("R" + i + "/C6");
                            getCol6.MoveNext();
                            col6 = getCol6.Current.Value;
                            XPathNodeIterator getCol7 = rows.Current.Select("R" + i + "/C7");
                            getCol7.MoveNext();
                            col7 = getCol7.Current.Value;
                            XPathNodeIterator getCol8 = rows.Current.Select("R" + i + "/C8");
                            getCol8.MoveNext();
                            col8 = getCol8.Current.Value;
                            XPathNodeIterator getCol9 = rows.Current.Select("R" + i + "/C9");
                            getCol9.MoveNext();
                            col9 = getCol9.Current.Value;
                            XPathNodeIterator getCol10 = rows.Current.Select("R" + i + "/C10");
                            getCol10.MoveNext();
                            col10 = getCol10.Current.Value;
                            //insert to DB
                            if (!col0.Equals(""))
                            {
                                Service svr = new Service();
                                svr.SVCCode = col0;
                                svr.SVCName = col1;
                                svr.SVCDesc = col2;
                                svr.HostCode = col3;
                                svr.StaffRole = col4;
                                svr.SVCType = col5;
                                svr.SVCObjective = col6;
                                svr.SVCSupport = col7;
                                svr.SVCCoverage = col8;
                                svr.SVCStart = col9;
                                svr.SVCEnd = col10;
                                bfAdmin.Services.InsertOnSubmit(svr);
                                try
                                {
                                    bfAdmin.SubmitChanges();
                                }
                                catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
                            }
                        }
                    }
                    #endregion
                }
                else if (title.Equals("Activity"))
                {
                    #region
                    string col0 = "", col1 = "", col2 = "", col3 = "";
                    XPathNodeIterator countRowString = allDocs.Current.Select("METADATA/ROWS");
                    countRowString.MoveNext();
                    int countRow = Convert.ToInt32(countRowString.Current.Value);

                    XPathNodeIterator rows = allDocs.Current.Select("DATA");
                    while (rows.MoveNext())
                    {
                        for (int i = 1; i < countRow; i++)
                        {
                            XPathNodeIterator getCol0 = rows.Current.Select("R" + i + "/C0");
                            getCol0.MoveNext();
                            col0 = getCol0.Current.Value;
                            XPathNodeIterator getCol1 = rows.Current.Select("R" + i + "/C1");
                            getCol1.MoveNext();
                            col1 = getCol1.Current.Value;
                            XPathNodeIterator getCol2 = rows.Current.Select("R" + i + "/C2");
                            getCol2.MoveNext();
                            col2 = getCol2.Current.Value;

                            XPathNodeIterator getCol3 = rows.Current.Select("R" + i + "/C3");
                            getCol3.MoveNext();
                            col3 = getCol3.Current.Value;
                            //insert to DB
                            if (!col0.Equals("") && !col2.Equals(""))
                            {
                                Activity act = new Activity();
                                act.ACTCode = col0;
                                act.ACTDesc = col1;
                                act.SVCCode = col2;
                                act.ICF_Code = col3;
                                bfAdmin.Activities.InsertOnSubmit(act);
                                try
                                {
                                    bfAdmin.SubmitChanges();
                                }
                                catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
                            }
                        }
                    }
                    #endregion
                }
                else if (title.Equals("SubActivity"))
                {
                    #region
                    string col0 = "", col1 = "", col2 = "";
                    XPathNodeIterator countRowString = allDocs.Current.Select("METADATA/ROWS");
                    countRowString.MoveNext();
                    int countRow = Convert.ToInt32(countRowString.Current.Value);

                    XPathNodeIterator rows = allDocs.Current.Select("DATA");
                    while (rows.MoveNext())
                    {
                        for (int i = 1; i < countRow; i++)
                        {
                            XPathNodeIterator getCol0 = rows.Current.Select("R" + i + "/C0");
                            getCol0.MoveNext();
                            col0 = getCol0.Current.Value;
                            XPathNodeIterator getCol1 = rows.Current.Select("R" + i + "/C1");
                            getCol1.MoveNext();
                            col1 = getCol1.Current.Value;
                            XPathNodeIterator getCol2 = rows.Current.Select("R" + i + "/C2");
                            getCol2.MoveNext();
                            col2 = getCol2.Current.Value;
                            //insert to DB
                            if (!col0.Equals("") && !col2.Equals(""))
                            {
                                SubActivity sact = new SubActivity();
                                sact.SACTCode = col0;
                                sact.SACTDesc = col1;
                                sact.ACTCode = col2;
                                bfAdmin.SubActivities.InsertOnSubmit(sact);
                                try
                                {
                                    bfAdmin.SubmitChanges();
                                }
                                catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
                            }
                        }
                    }
                    #endregion
                }
                else if (title.Equals("Material"))
                {
                    #region
                    string col0 = "", col1 = "", col2 = "", col3 = "", col4 = "", col5 = "", col6 = "";
                    XPathNodeIterator countRowString = allDocs.Current.Select("METADATA/ROWS");
                    countRowString.MoveNext();
                    int countRow = Convert.ToInt32(countRowString.Current.Value);

                    XPathNodeIterator rows = allDocs.Current.Select("DATA");
                    while (rows.MoveNext())
                    {
                        for (int i = 1; i < countRow; i++)
                        {
                            XPathNodeIterator getCol0 = rows.Current.Select("R" + i + "/C0");
                            getCol0.MoveNext();
                            col0 = getCol0.Current.Value;
                            XPathNodeIterator getCol1 = rows.Current.Select("R" + i + "/C1");
                            getCol1.MoveNext();
                            col1 = getCol1.Current.Value;
                            XPathNodeIterator getCol2 = rows.Current.Select("R" + i + "/C2");
                            getCol2.MoveNext();
                            col2 = getCol2.Current.Value;
                            XPathNodeIterator getCol3 = rows.Current.Select("R" + i + "/C3");
                            getCol3.MoveNext();
                            col3 = getCol3.Current.Value;
                            XPathNodeIterator getCol4 = rows.Current.Select("R" + i + "/C4");
                            getCol4.MoveNext();
                            col4 = getCol4.Current.Value;
                            XPathNodeIterator getCol5 = rows.Current.Select("R" + i + "/C5");
                            getCol5.MoveNext();
                            col5 = getCol5.Current.Value;
                            XPathNodeIterator getCol6 = rows.Current.Select("R" + i + "/C6");
                            getCol6.MoveNext();
                            col6 = getCol6.Current.Value;
                            //insert to DB
                            if (!col0.Equals("") && !col5.Equals(""))
                            {
                                Material mt = new Material();
                                mt.MaterialCode = col0;
                                mt.MaterialDesc = col1;
                                mt.Unit = col2;
                                mt.EstimatedPrice = col3;
                                mt.RealPrice = col4;
                                mt.SVCCode = col5;
                                mt.Note = col6;
                                bfAdmin.Materials.InsertOnSubmit(mt);
                                try
                                {
                                    bfAdmin.SubmitChanges();
                                }
                                catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
                            }
                        }
                    }
                    #endregion
                }
            }
            return null;
        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            DropDownListFrom.Items.Clear();
            this.saveFile();
        }

        private void saveFile()
        {
            string fileName = Path.GetFileName(FileUpload1.PostedFile.FileName);
            string fileExtension = Path.GetExtension(FileUpload1.PostedFile.FileName);
            if (FileUpload1.HasFile && fileName != "" && fileExtension != "")
            {
                try
                {
                    Stream spreadsheet = FileUpload1.PostedFile.InputStream;
                    IExcelDataReader excelReader = null;
                    if (fileExtension.Equals(".xls"))
                    {
                        excelReader = ExcelReaderFactory.CreateBinaryReader(spreadsheet);
                    }
                    else if (fileExtension.Equals(".xlsx"))
                    {
                        excelReader = ExcelReaderFactory.CreateOpenXmlReader(spreadsheet);
                    }

                    //first row is column name not parse
                    //excelReader.IsFirstRowAsColumnNames = true;

                    resultFromUpload = excelReader.AsDataSet();
                    //resultFromUpload.Locale = new System.Globalization.CultureInfo("th-TH");
                    for (int i = 0; i < resultFromUpload.Tables.Count; i++)
                    {
                        DropDownListFrom.Items.Add(resultFromUpload.Tables[i].TableName.Trim());
                    }
                    excelReader.Close();
                }
                catch { }
            }
        }

        private void excelToDB(DataTable data)
        {
            //clean DB
            #region
            if (RadioButtonList_type.SelectedIndex == 1) //selecte create new sheet
            {
                if (DropDownListTo.SelectedIndex == 0)
                {
                    var allInSvr = from c in bfAdmin.Services select c;
                    bfAdmin.Services.DeleteAllOnSubmit(allInSvr);
                    try
                    {
                        bfAdmin.SubmitChanges();
                    }
                    catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
                }
                else if (DropDownListTo.SelectedIndex == 1)
                {
                    var allInAct = from c in bfAdmin.Activities select c;
                    bfAdmin.Activities.DeleteAllOnSubmit(allInAct);
                    try
                    {
                        bfAdmin.SubmitChanges();
                    }
                    catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
                }
                else if (DropDownListTo.SelectedIndex == 2)
                {
                    var allInSAct = from c in bfAdmin.SubActivities select c;
                    bfAdmin.SubActivities.DeleteAllOnSubmit(allInSAct);
                    try
                    {
                        bfAdmin.SubmitChanges();
                    }
                    catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
                }
                else
                {
                    var allInMat = from c in bfAdmin.Materials select c;
                    bfAdmin.Materials.DeleteAllOnSubmit(allInMat);
                    try
                    {
                        bfAdmin.SubmitChanges();
                    }
                    catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
                }
            }
                #endregion
            if (DropDownListTo.SelectedIndex == 0)
            {
                int skip = 0;
                foreach (DataRow row in data.Rows)
                {
                    if (row[0].ToString() != null && skip > 0)
                    {
                        Service svr = new Service();
                        try
                        {
                            svr.SVCCode = row[0].ToString();
                            svr.SVCName = row[1].ToString() == null ? "" : row[1].ToString();
                            svr.SVCDesc = row[2].ToString() == null ? "" : row[2].ToString();
                            svr.HostCode = row[3].ToString() == null ? "" : row[3].ToString();
                            svr.StaffRole = row[4].ToString() == null ? "" : row[4].ToString();
                            svr.SVCType = row[5].ToString() == null ? "" : row[5].ToString();
                            svr.SVCObjective = row[6].ToString() == null ? "" : row[6].ToString();
                            svr.SVCSupport = row[7].ToString() == null ? "" : row[7].ToString();
                            svr.SVCCoverage = row[8].ToString() == null ? "" : row[8].ToString();
                            svr.SVCStart = row[9].ToString() == null ? "" : row[9].ToString();
                            svr.SVCEnd = row[10].ToString() == null ? "" : row[10].ToString();
                        }
                        catch { }

                        var Svr = from c in bfAdmin.Services where c.SVCCode.ToString().Trim().Equals(row[0].ToString().Trim()) select c;
                        if (Svr != null)
                        {
                            bfAdmin.Services.DeleteAllOnSubmit(Svr);
                            try
                            {
                                bfAdmin.SubmitChanges();
                            }
                            catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
                        }
                        bfAdmin.Services.InsertOnSubmit(svr);
                        try
                        {
                            bfAdmin.SubmitChanges();
                        }
                        catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
                    }
                    skip = 1;
                }
            }
            else if (DropDownListTo.SelectedIndex == 1)
            {
                int skip = 0;
                foreach (DataRow row in data.Rows)
                {
                    if (row[0].ToString() != null && skip > 0)
                    {
                        Activity act = new Activity();
                        try
                        {
                            act.ACTCode = row[0].ToString();
                            act.ACTDesc = row[1].ToString() == null ? "" : row[1].ToString();
                            act.SVCCode = row[2].ToString() == null ? "" : row[2].ToString();
                        }
                        catch { }

                        var Act = from c in bfAdmin.Activities where c.ACTCode.ToString().Trim().Equals(row[0].ToString().Trim()) select c;
                        bfAdmin.Activities.DeleteAllOnSubmit(Act);
                        try
                        {
                            bfAdmin.SubmitChanges();
                        }
                        catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
                        bfAdmin.Activities.InsertOnSubmit(act);
                        try
                        {
                            bfAdmin.SubmitChanges();
                        }
                        catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
                    }
                    skip = 1;
                }
            }
            else if (DropDownListTo.SelectedIndex == 2)
            {
                int skip = 0;
                foreach (DataRow row in data.Rows)
                {
                    if (row[0].ToString() != null && skip > 0)
                    {
                        SubActivity sact = new SubActivity();
                        try
                        {
                            sact.SACTCode = row[0].ToString();
                            sact.SACTDesc = row[1].ToString() == null ? "" : row[1].ToString();
                            sact.ACTCode = row[2].ToString() == null ? "" : row[2].ToString();
                        }
                        catch { }

                        var Sact = from c in bfAdmin.SubActivities where c.SACTCode.ToString().Trim().Equals(row[0].ToString().Trim()) select c;
                        bfAdmin.SubActivities.DeleteAllOnSubmit(Sact);
                        try
                        {
                            bfAdmin.SubmitChanges();
                        }
                        catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
                        bfAdmin.SubActivities.InsertOnSubmit(sact);
                        try
                        {
                            bfAdmin.SubmitChanges();
                        }
                        catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
                    }
                    skip = 1;
                }
            }
            else
            {
                int skip = 0;
                foreach (DataRow row in data.Rows)
                {
                    if (row[0].ToString() != null && skip > 0)
                    {
                        Material mat = new Material();
                        try
                        {
                            mat.MaterialCode = row[0].ToString();
                            mat.MaterialDesc = row[1].ToString() == null ? "" : row[1].ToString();
                            mat.Unit = row[2].ToString() == null ? "" : row[2].ToString();
                            mat.EstimatedPrice = row[3].ToString() == null ? "" : row[3].ToString();
                            mat.RealPrice = row[4].ToString() == null ? "" : row[4].ToString();
                            mat.SVCCode = row[5].ToString() == null ? "" : row[5].ToString();
                        }
                        catch { }

                        var Mat = from c in bfAdmin.Materials where c.MaterialCode.ToString().Trim().Equals(row[0].ToString().Trim()) select c;
                        bfAdmin.Materials.DeleteAllOnSubmit(Mat);
                        try
                        {
                            bfAdmin.SubmitChanges();
                        }
                        catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }

                        bfAdmin.Materials.InsertOnSubmit(mat);
                        try
                        {
                            bfAdmin.SubmitChanges();
                        }
                        catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
                    }
                    skip = 1;
                }
            }

            //refresh page
            Response.Redirect(Request.RawUrl);
        }

        protected void LinkButton_signOut_Click(object sender, EventArgs e)
        {
            FormsAuthentication.SignOut();
            Response.Redirect("Casemanager.aspx");
        }

        protected void ButtonOK_Click(object sender, EventArgs e)
        {
            try
            {
                for (int i = 0; i < resultFromUpload.Tables.Count; i++)
                {
                    if (DropDownListFrom.SelectedItem.ToString().Trim().Equals(resultFromUpload.Tables[i].TableName.Trim()))
                    {
                        excelToDB(resultFromUpload.Tables[i]);
                        break;
                    }
                }
            }
            catch { }
            DropDownListFrom.Items.Clear();
            resultFromUpload = null;
        }

    }
}