using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Spreadsheet
{
    public partial class PreBenefitAdminCode : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                loadData();
            }
        }

        private void loadData()
        {
            BenefitAdminDataContext db = new BenefitAdminDataContext();
            var data = from a in db.WelfareCategories select a;
            foreach (WelfareCategory w in data)
            {
                ListItem li = new ListItem();
                li.Text = w.WelfareDescription;
                li.Value = w.WelfareID;
                RadioButtonList_type.Items.Add(li);
            }

        }

        protected void Button_ok_Click(object sender, EventArgs e)
        {
            var selectedItem = RadioButtonList_type.SelectedItem;
            if (selectedItem != null)
            {
                Session["childtype"] = selectedItem.Value.Trim();
                Response.Redirect("BenefitAdminCode.aspx");
            }
        }

    }
}