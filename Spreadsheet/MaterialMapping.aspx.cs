using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Spreadsheet
{
    public partial class MaterialMapping : System.Web.UI.Page
    {
        private static List<Activity> activityList;
        private static IQueryable<Material> mat;
        private static int index;
        BenefitAdminDataContext db = new BenefitAdminDataContext();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                index = 0;
                loadData();
                bindData();
            }
        }

        private void loadData()
        {
            activityList = HtmlManager.act.ToList();
            mat = from a in db.Materials select a;
        }

        private void bindData()
        {
            GridView_mapping.DataSource = mat;
            GridView_mapping.DataBind();

            List<Activity> tmp = new List<Activity>();
            tmp.Add(activityList[index]);
            GridView_activity.DataSource = tmp;
            GridView_activity.DataBind();

        }

        protected void Button_ok_Click(object sender, EventArgs e)
        {
            if (index < activityList.Count)
            {
                foreach (GridViewRow row in GridView_mapping.Rows)
                {
                    CheckBox chk = (CheckBox)row.FindControl("CheckBox_select");
                    if (chk.Checked)
                    {
                        Material_ACT mact = new Material_ACT();
                        mact.MaterialCode = row.Cells[1].Text.Trim();
                        mact.ACTCode = activityList[index].ACTCode.Trim();

                        db.Material_ACTs.InsertOnSubmit(mact);
                        try
                        {
                            db.SubmitChanges();
                        }
                        catch { db = new BenefitAdminDataContext(); }
                    }
                }
                index++;
                if (index < activityList.Count)
                {
                    bindData();
                }
                else
                {
                    // redirect
                }
            }
        }
    }
}