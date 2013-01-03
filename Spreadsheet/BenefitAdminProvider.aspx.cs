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
        string fileName = "provider.html";
        private static string staticXmlFileName = "provider.xml";
        private static string staticHtmlFileName = "provider.html";

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!User.Identity.IsAuthenticated || !User.Identity.Name.Equals("ProviderAdmin"))
            {
                Response.Redirect("Casemanager.aspx");
            }
            else
            {
                HtmlManager.createHtml(fileName);
                Label_user.Text = "Welcome " + Session["user"];
            }
        }

        [System.Web.Services.WebMethod]
        public static void SaveData(string xml, string html)
        {
            HtmlManager.saveHtml(xml, staticXmlFileName);
            HtmlManager.saveHtml(html, staticHtmlFileName);
        }

        protected void LinkButton_signOut_Click(object sender, EventArgs e)
        {
            FormsAuthentication.SignOut();
            Response.Redirect("Casemanager.aspx");
        }

    }
}