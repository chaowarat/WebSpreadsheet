using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Spreadsheet
{
    public partial class ServiceMapping : System.Web.UI.Page
    {
        private static List<Service> serviceList;
        private static IQueryable<DisabilityType> disType;
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
            serviceList = HtmlManager.ser.ToList();
            disType = from a in db.DisabilityTypes select a;
        }

        private void bindData()
        {
            GridView_mapping.DataSource = disType;
            GridView_mapping.DataBind();
        }

        protected void Button_ok_Click(object sender, EventArgs e)
        {
            if (index < serviceList.Count)
            {
                foreach (GridViewRow row in GridView_mapping.Rows)
                {
                    CheckBox chk = (CheckBox)row.FindControl("CheckBox_select");
                    if (chk.Checked)
                    {
                        string disCode = Session["childtype"].ToString().Trim() + row.Cells[2].Text.Trim();
                        ServiceChildTypeMapping scm = new ServiceChildTypeMapping();
                        scm.SVCCode = serviceList[index].SVCCode.Trim();
                        scm.SubWelfareID = disCode;
                        db.ServiceChildTypeMappings.InsertOnSubmit(scm);
                        try
                        {
                            db.SubmitChanges();
                        }
                        catch { db = new BenefitAdminDataContext(); }
                    }
                }
                index++;
                bindData();
            }
        }

        protected void GridView_mapping_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow && e.Row.RowIndex == 0 && index < serviceList.Count)
            {
                Label lblServiceId = (Label)e.Row.FindControl("Label_service_id");
                lblServiceId.Text = serviceList[index].SVCCode.Trim();
                Label lblServiceName = (Label)e.Row.FindControl("Label_service_name");
                lblServiceName.Text = serviceList[index].SVCName.Trim();
            }
        }
    }
}