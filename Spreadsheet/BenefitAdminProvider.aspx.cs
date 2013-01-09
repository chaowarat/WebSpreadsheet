using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;
using System.Linq;
using System.IO;
using System.Web.Security;
using System.Xml.Linq;
using System.Xml.XPath;
using System.Data.OleDb;
using System.Data;

namespace Spreadsheet
{
    public partial class BenefitAdminProvider : System.Web.UI.Page
    {
        private static string staticHtmlFileName = "provider.html";

        protected void Page_Load(object sender, EventArgs e)
        {
            //if (!User.Identity.IsAuthenticated || !User.Identity.Name.Equals("ProviderAdmin"))
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
            return HtmlManager.createHtml(staticHtmlFileName);
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
            var allInMis = from c in bfAdmin.Ministries select c;
            bfAdmin.Ministries.DeleteAllOnSubmit(allInMis);
            try
            {
                bfAdmin.SubmitChanges();
            }
            catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }

            var allInOr = from c in bfAdmin.Organizations select c;
            bfAdmin.Organizations.DeleteAllOnSubmit(allInOr);
            try
            {
                bfAdmin.SubmitChanges();
            }
            catch (Exception) { bfAdmin = new BenefitAdminDataContext(); }

            var allInPro = from c in bfAdmin.Providers select c;
            bfAdmin.Providers.DeleteAllOnSubmit(allInPro);
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
                if (title.Equals("Ministry"))
                {
                    #region
                    string col0 = "", col1 = "";
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
                            if (!col0.Equals(""))
                            {
                                Ministry min = new Ministry();
                                min.MinistryCode = col0;
                                min.MinistryName = col1;
                                bfAdmin.Ministries.InsertOnSubmit(min);
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
                else if(title.Equals("Organization"))
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
                            if (!col0.Equals(""))
                            {
                                Organization org = new Organization();
                                org.OrgCode = col0;
                                org.OrgName = col1;
                                org.MinistryCode = col2;
                                bfAdmin.Organizations.InsertOnSubmit(org);
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
                else if(title.Equals("Provider"))
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
                            if (!col0.Equals(""))
                            {
                                Provider pro = new Provider();
                                pro.ProviderCode = col0;
                                pro.ProviderName = col1;
                                pro.OrgCode = col2;
                                bfAdmin.Providers.InsertOnSubmit(pro);
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

        protected void LinkButton_signOut_Click(object sender, EventArgs e)
        {
            FormsAuthentication.SignOut();
            Response.Redirect("Casemanager.aspx");
        }

    }
}