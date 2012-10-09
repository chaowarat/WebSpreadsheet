using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Spreadsheet
{
    public partial class Main : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            Login_admin.Focus();
        }

        protected void Login_admin_Authenticate(object sender, AuthenticateEventArgs e)
        {
            if (Membership.ValidateUser(Login_admin.UserName, Login_admin.Password))
            {
                Session["user"] = Login_admin.UserName;
                if (Roles.IsUserInRole(Login_admin.UserName, "provider"))
                {
                    FormsAuthentication.SetAuthCookie("provider", true);
                    Response.Redirect("BenefitAdminProvider.aspx");
                }
                else if (Roles.IsUserInRole(Login_admin.UserName, "code"))
                {
                    FormsAuthentication.SetAuthCookie("code", true);
                    Response.Redirect("BenefitAdminCode.aspx");
                }
                else if (Roles.IsUserInRole(Login_admin.UserName, "cost"))
                {
                    FormsAuthentication.SetAuthCookie("cost", true);
                    Response.Redirect("BenefitAdminCost.aspx");
                }
            }
        }
    }
}