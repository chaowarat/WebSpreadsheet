using System;
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

namespace Spreadsheet
{
    public partial class BenefitAdminCost : System.Web.UI.Page
    {
        BenefitAdminDataContext bfAdmin = new BenefitAdminDataContext();
        string fileName = "cost.html";
        private static string staticXmlFileName = "cost.xml";

        protected void Page_Load(object sender, EventArgs e)
        {
            //if (!Request.IsAuthenticated || !User.Identity.Name.Equals("CostAdmin"))
            //{
            //    Response.Redirect("Casemanager.aspx");
            //}
            //else
            //{
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                HtmlManager.createHtml(fileName);
                //Label_user.Text = "Welcome " + Session["user"];
            //}
        }

        [System.Web.Services.WebMethod]
        public static void SaveData(string data)
        {
            data = data.Replace("&", "%26");
            HtmlManager.saveHtml(data, staticXmlFileName);
            BenefitAdminDataContext bfAdmin = new BenefitAdminDataContext();
            var allInAcost = from c in bfAdmin.ActivityCosts select c;
            bfAdmin.ActivityCosts.DeleteAllOnSubmit(allInAcost);
            try
            {
                bfAdmin.SubmitChanges();
            }
            catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }

            var allInAnno = from c in bfAdmin.Annotations select c;
            bfAdmin.Annotations.DeleteAllOnSubmit(allInAnno);
            try
            {
                bfAdmin.SubmitChanges();
            }
            catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }

            MemoryStream mem = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(data));
            XPathDocument document = new XPathDocument(mem);
            XPathNavigator navigator = document.CreateNavigator();
            XPathNodeIterator allDocs = navigator.Select("DOCUMENTS/DOCUMENT");
            while (allDocs.MoveNext())
            {
                string title = allDocs.Current.GetAttribute("title", "");
                if (title.Equals("Annotation"))
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
                            if (!col0.Equals(""))
                            {
                                Annotation anno = new Annotation();
                                anno.AID = col0;
                                anno.AText = col1;
                                anno.AnnotationID = col2;
                                anno.Reference = col3;

                                bfAdmin.Annotations.InsertOnSubmit(anno);
                                try
                                {
                                    bfAdmin.SubmitChanges();
                                }
                                catch (Exception) { bfAdmin = new CaseManagerDBDataContext(); }
                            }
                        }
                    }
                    #endregion
                }
                else if (title.Equals("ActivityCost"))
                {
                    string col0 = "", col1 = "", col2 = "", col3 = "", col4 = "", col5 = "", col6 = "", col7 = "", col8 = "", col9 = "";
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
                            //insert to DB
                            if (!col0.Equals(""))
                            {
                                ActivityCost acost = new ActivityCost();
                                acost.ACTCode = col0;
                                acost.Unit = col1;
                                acost.LabourCost = col2;
                                acost.MaterialCost = col3;
                                acost.CC_Equipment = col4;
                                acost.CC_Building = col5;
                                acost.IndirectCost = col6;
                                acost.ProposedCost = col7;
                                acost.CurrentCost = col8;
                                acost.ReferencedCostOrg = col9;
                                bfAdmin.ActivityCosts.InsertOnSubmit(acost);
                                try
                                {
                                    bfAdmin.SubmitChanges();
                                }
                                catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
                            }
                        }
                    }
                }

            }
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
            string connectionString = "";
            if (FileUpload1.HasFile && fileName != "" && fileExtension != "")
            {
                string fileLocation = Server.MapPath("~/sheets/" + fileName);
                FileUpload1.SaveAs(fileLocation);

                //Check whether file extension is xls or xslx
                if (fileExtension == ".xls")
                {
                    connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileLocation + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                }
                else if (fileExtension == ".xlsx")
                {
                    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileLocation + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                }

                //Create OleDB Connection and OleDb Command
                OleDbConnection con = new OleDbConnection(connectionString);
                OleDbCommand cmd = new OleDbCommand();
                cmd.CommandType = System.Data.CommandType.Text;
                cmd.Connection = con;
                OleDbDataAdapter dAdapter = new OleDbDataAdapter(cmd);
                DataTable dtExcelRecords = new DataTable();
                try
                {
                    con.Open();
                    DataTable dtExcelSheetName = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                    //For many Sheet
                    String[] excelSheets = new String[dtExcelSheetName.Rows.Count];
                    int i = 0;
                    // Add the sheet name to the string array.
                    foreach (DataRow row in dtExcelSheetName.Rows)
                    {
                        excelSheets[i] = row["TABLE_NAME"].ToString();
                        DropDownListFrom.Items.Add(excelSheets[i].Substring(0, excelSheets[i].Length - 1));
                        i++;
                    }
                    Application["fileName"] = fileName;
                }
                catch (Exception) { }
                finally
                {
                    con.Close();
                }
            }
        }

        private void excelToDB(DataTable data)
        {
            //remove data if radio selected 1
            if (RadioButtonList_type.SelectedIndex == 1) //selecte create new sheet
            {
                if (DropDownListTo.SelectedIndex == 0)
                {
                    var allInAnno = from c in bfAdmin.Annotations select c;
                    bfAdmin.Annotations.DeleteAllOnSubmit(allInAnno);
                    try
                    {
                        bfAdmin.SubmitChanges();
                    }
                    catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
                }
                else if (DropDownListTo.SelectedIndex == 1)
                {
                    var allInAcost = from c in bfAdmin.ActivityCosts select c;
                    bfAdmin.ActivityCosts.DeleteAllOnSubmit(allInAcost);
                    try
                    {
                        bfAdmin.SubmitChanges();
                    }
                    catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
                } 
            }
            if (DropDownListTo.SelectedIndex == 0)
            {
                foreach (DataRow row in data.Rows)
                {
                    Annotation annot = new Annotation();
                    annot.AID = row[0].ToString();
                    annot.AText = row[1].ToString();
                    annot.AnnotationID = row[2].ToString();
                    annot.Reference = row[3].ToString();

                    var Annot = from c in bfAdmin.Annotations where c.AID.ToString().Trim().Equals(row[0].ToString().Trim()) select c;
                    bfAdmin.Annotations.DeleteAllOnSubmit(Annot);
                    try
                    {
                        bfAdmin.SubmitChanges();
                    }
                    catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
                    bfAdmin.Annotations.InsertOnSubmit(annot);
                    try
                    {
                        bfAdmin.SubmitChanges();
                    }
                    catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
                }
            }
            else if (DropDownListTo.SelectedIndex == 1)
            {
                foreach (DataRow row in data.Rows)
                {
                    ActivityCost act = new ActivityCost();
                    act.ACTCode = row[0].ToString();
                    act.Unit = row[1].ToString();
                    act.LabourCost = row[2].ToString();
                    act.MaterialCost = row[3].ToString();
                    act.CC_Equipment = row[4].ToString();
                    act.CC_Building = row[5].ToString();
                    act.IndirectCost = row[6].ToString();
                    act.ProposedCost = row[7].ToString();
                    act.CurrentCost = row[8].ToString();
                    act.ReferencedCostOrg = row[9].ToString();

                    var Acost = from c in bfAdmin.ActivityCosts where c.ACTCode.ToString().Trim().Equals(row[0].ToString().Trim()) select c;
                    bfAdmin.ActivityCosts.DeleteAllOnSubmit(Acost);
                    try
                    {
                        bfAdmin.SubmitChanges();
                    }
                    catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
                    bfAdmin.ActivityCosts.InsertOnSubmit(act);
                    try
                    {
                        bfAdmin.SubmitChanges();
                    }
                    catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
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
            string fileName = Application["fileName"].ToString();
            string fileExtension = Path.GetExtension(fileName);
            string connectionString = "";
            string fileLocation = Server.MapPath("~/sheets/" + fileName);

            //Check whether file extension is xls or xslx
            if (fileExtension == ".xls")
            {
                connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileLocation + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
            }
            else if (fileExtension == ".xlsx")
            {
                connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileLocation + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
            }

            //Create OleDB Connection and OleDb Command
            OleDbConnection con = new OleDbConnection(connectionString);
            OleDbCommand cmd = new OleDbCommand();
            cmd.CommandType = System.Data.CommandType.Text;
            cmd.Connection = con;
            OleDbDataAdapter dAdapter = new OleDbDataAdapter(cmd);
            DataTable dtExcelRecords = new DataTable();
            try
            {
                con.Open();
                DataTable dtExcelSheetName = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                //For many Sheet
                String excelSheets;
                // Add the sheet name to the string array.
                foreach (DataRow row in dtExcelSheetName.Rows)
                {
                    excelSheets = row["TABLE_NAME"].ToString();
                    if (excelSheets.Substring(0, excelSheets.Length - 1).Equals(DropDownListFrom.SelectedItem.ToString()))
                    {
                        cmd.CommandText = "SELECT * FROM [" + excelSheets + "]";
                        dAdapter.SelectCommand = cmd;
                        dAdapter.Fill(dtExcelRecords);
                        excelToDB(dtExcelRecords);
                        break;
                    }
                }

            }
            catch (Exception) { }
            finally
            {
                con.Close();
                System.IO.File.Delete(fileLocation);
            }
        }

    }
}