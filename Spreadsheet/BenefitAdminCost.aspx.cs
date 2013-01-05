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
using Excel;

namespace Spreadsheet
{
    public partial class BenefitAdminCost : System.Web.UI.Page
    {
        BenefitAdminDataContext bfAdmin = new BenefitAdminDataContext();
        private static string fileName = "cost.html";
        private static string staticXmlFileName = "cost.xml";
        private static DataSet resultFromUpload = null;

        protected void Page_Load(object sender, EventArgs e)
        {
            //if (!Request.IsAuthenticated || !User.Identity.Name.Equals("CostAdmin"))
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
                                catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }
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
                    for (int i = 0; i < resultFromUpload.Tables.Count; i++)
                    {
                        DropDownListFrom.Items.Add(resultFromUpload.Tables[i].TableName.Trim());
                    }
                    excelReader.Close();
                }
                catch{}
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
                int skip = 0;
                foreach (DataRow row in data.Rows)
                {
                    if (row[0].ToString() != null && skip > 0)
                    {
                        Annotation annot = new Annotation();
                        try
                        {
                            annot.AID = row[0].ToString();
                            annot.AText = row[1].ToString() == null ? "" : row[1].ToString();
                            annot.AnnotationID = row[2].ToString() == null ? "" : row[2].ToString();
                            annot.Reference = row[3].ToString() == null ? "" : row[3].ToString();
                        }
                        catch { }

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
                        ActivityCost act = new ActivityCost();
                        try
                        {
                            act.ACTCode = row[0].ToString();
                            act.Unit = row[1].ToString() == null ? "" : row[1].ToString();
                            act.LabourCost = row[2].ToString() == null ? "" : row[2].ToString();
                            act.MaterialCost = row[3].ToString() == null ? "" : row[3].ToString();
                            act.CC_Equipment = row[4].ToString() == null ? "" : row[4].ToString();
                            act.CC_Building = row[5].ToString() == null ? "" : row[5].ToString();
                            act.IndirectCost = row[6].ToString() == null ? "" : row[6].ToString();
                            act.ProposedCost = row[7].ToString() == null ? "" : row[7].ToString();
                            act.CurrentCost = row[8].ToString() == null ? "" : row[8].ToString();
                            act.ReferencedCostOrg = row[9].ToString();
                        }
                        catch { }

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