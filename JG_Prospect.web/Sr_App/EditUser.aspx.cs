using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using JG_Prospect.BLL;
using JG_Prospect.Common;
using JG_Prospect.Common.modal;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Data.OleDb;
using System.Text;
//using Word = Microsoft.Office.Interop.Word;
using System.Net;
using System.Net.Mail;
using System.Web.UI.DataVisualization.Charting;
using System.Xml.Serialization;
using System.Xml;
using JG_Prospect.App_Code;
using OfficeOpenXml;
using Newtonsoft.Json;


namespace JG_Prospect
{
    #region '--Enums--'

    public class HrData
    {
        public string status { get; set; }
        public string count { get; set; }
    }

    #endregion

    public partial class EditUser : System.Web.UI.Page
    {
        #region '--Members--'

        #endregion

        #region '--Properties--'

        private SortDirection SalesUserSortDirection
        {
            get
            {
                if (ViewState["SalesUserSortDirection"] == null)
                {
                    return SortDirection.Descending;
                }
                return (SortDirection)ViewState["SalesUserSortDirection"];
            }
            set
            {
                ViewState["SalesUserSortDirection"] = value;
            }
        }

        private string SalesUserSortExpression
        {
            get
            {
                if (ViewState["SalesUserSortExpression"] == null)
                {
                    return "CreatedOn";
                }
                return Convert.ToString(ViewState["SalesUserSortExpression"]);
            }
            set
            {
                ViewState["SalesUserSortExpression"] = value;
            }
        }

        private DataTable SelectedUsers
        {
            get
            {
                if (ViewState["SelectedUsers"] == null)
                {
                    return null;
                }
                return (DataTable)ViewState["SelectedUsers"];
            }
            set
            {
                ViewState["SelectedUsers"] = value;
            }
        }

        #endregion

        #region '--Page Events--'

        protected void Page_Load(object sender, EventArgs e)
        {
            CommonFunction.AuthenticateUser();

            if (Convert.ToString(Session["usertype"]).Contains("Admin"))
            {
                btnExport.Visible = true;
            }
            else
            {
                btnExport.Visible = false;
            }

            if (JGSession.Designation.ToUpper() == "RECRUITER" || JGSession.Designation.ToUpper() == "ADMIN")
            {
                lbtnChangeStatusForSelected.Visible =
                lbtnDeleteSelected.Visible = true;
            }
            else
            {
                lbtnChangeStatusForSelected.Visible =
                lbtnDeleteSelected.Visible = false;
            }

            if (JGSession.DesignationId == (int)JGConstant.DesignationType.Admin || JGSession.DesignationId == (int)JGConstant.DesignationType.IT_Lead)
            {
                lbtnDeleteSelected.Visible = true;
            }
            else
            {
                lbtnDeleteSelected.Visible = false;
            }

            if (!IsPostBack)
            {
                CalendarExtender1.StartDate = DateTime.Now;

                Session["DeactivationStatus"] = "";
                Session["FirstNameNewSC"] = "";
                Session["LastNameNewSC"] = "";
                Session["DesignitionSC"] = "";
                Session["HighlightUsersForTypes"] = null;

                //binddata();
                //DataSet dsCurrentPeriod = UserBLL.Instance.Getcurrentperioddates();
                //bindPayPeriod(dsCurrentPeriod);
                //txtfrmdate.Text = DateTime.Now.AddDays(-14).ToString("MM/dd/yyyy");
                txtfrmdate.Text = "All";
                txtTodate.Text = DateTime.Now.ToString("MM/dd/yyyy");

                FillCustomer();
                BindDesignations();

                GetSalesUsersStaticticsAndData(true);
            }
            else
            {
                if (Session["HighlightUsersForTypes"] != null)
                {
                    HighlightUsersForTypes((DataTable)Session["HighlightUsersForTypes"], drpUser);
                }
            }
        }

        #endregion

        #region '--Control Events--'

        #region grdUsers - Filters

        protected void chkAllDates_CheckedChanged(object sender, EventArgs e)
        {
            if (chkAllDates.Checked)
            {
                txtfrmdate.Enabled = false;
                txtTodate.Enabled = false;
                txtfrmdate.Text = "All";
            }
            else
            {
                txtfrmdate.Enabled = true;
                txtTodate.Enabled = true;
                txtfrmdate.Text = DateTime.Now.AddDays(-14).ToString("MM/dd/yyyy");
            }
            //BindGrid();
            GetSalesUsersStaticticsAndData(true);
        }

        protected void ddlUserStatus_PreRender(object sender, EventArgs e)
        {
            DropDownList ddlStatus = (DropDownList)sender;
            ddlStatus = JG_Prospect.Utilits.FullDropDown.UserStatusDropDown_Set_ImageAtt(ddlStatus);
        }

        protected void ddlFilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetSalesUsersStaticticsAndData(true);
            //BindGrid();
        }

        protected void txtfrmdate_TextChanged(object sender, EventArgs e)
        {
            GetSalesUsersStaticticsAndData(true);
        }

        protected void txtTodate_TextChanged(object sender, EventArgs e)
        {
            GetSalesUsersStaticticsAndData(true);
        }

        protected void btnSearchGridData_Click(object sender, EventArgs e)
        {
            GetSalesUsersStaticticsAndData(true);
        }

        #endregion

        #region grdUsers - User List

        protected void grdUsers_PreRender(object sender, EventArgs e)
        {
            GridView gv = (GridView)sender;

            if (gv.Rows.Count > 0)
            {
                gv.UseAccessibleHeader = true;
                gv.HeaderRow.TableSection = TableRowSection.TableHeader;
                gv.FooterRow.TableSection = TableRowSection.TableFooter;
            }

            if (gv.TopPagerRow != null)
            {
                gv.TopPagerRow.TableSection = TableRowSection.TableHeader;
            }
            if (gv.BottomPagerRow != null)
            {
                gv.BottomPagerRow.TableSection = TableRowSection.TableFooter;
            }
        }

        protected void grdUsers_RowDataBound(object sender, GridViewRowEventArgs e)
        {

            try
            {
                if (e.Row.RowType == DataControlRowType.DataRow)
                {
                    Label lblPrimaryPhone = (e.Row.FindControl("lblPrimaryPhone") as Label);
                    DropDownList ddlStatus = (e.Row.FindControl("ddlStatus") as DropDownList);//Find the DropDownList in the Row
                    DropDownList ddlContactType = (e.Row.FindControl("ddlContactType") as DropDownList);
                    HyperLink hypTechTask = e.Row.FindControl("hypTechTask") as HyperLink;
                    LinkButton lnkDelete = e.Row.FindControl("lnkDelete") as LinkButton;
                    ImageButton imgBookmark = e.Row.FindControl("imgBookmark") as ImageButton;

                    if (!string.IsNullOrEmpty(DataBinder.Eval(e.Row.DataItem, "BookmarkByInstallUserId").ToString()))
                    {
                        imgBookmark.ImageUrl = "~/Sr_App/img/starGreen.png";
                        imgBookmark.Attributes["title"] = DataBinder.Eval(e.Row.DataItem, "BookmarkByInstallUserId").ToString() + "-"
                            + DataBinder.Eval(e.Row.DataItem, "BookmarkBy").ToString() + " & "
                            + DataBinder.Eval(e.Row.DataItem, "BookmarkTime").ToString();
                        if (Convert.ToString(Session["usertype"]).Contains("Admin") || Convert.ToString(Session["usertype"]).Contains("SM")
                            || Convert.ToString(Session["UserId"]) == DataBinder.Eval(e.Row.DataItem, "BookmarkByInstallUserId").ToString())
                        {
                            imgBookmark.CommandName = "UnStarUser";
                            imgBookmark.CommandArgument = DataBinder.Eval(e.Row.DataItem, "Id").ToString();
                        }
                        else
                        {
                            imgBookmark.OnClientClick = "return false;";
                        }
                    }
                    else
                    {
                        imgBookmark.ImageUrl = "~/Sr_App/img/starGray.png";
                        imgBookmark.CommandName = "BookmarkUser";
                        imgBookmark.CommandArgument = DataBinder.Eval(e.Row.DataItem, "Id").ToString();
                    }


                    ddlStatus = JG_Prospect.Utilits.FullDropDown.FillUserStatus(ddlStatus);

                    ddlContactType = BindContactDllForGrid(ddlContactType);

                    System.Web.UI.HtmlControls.HtmlAnchor aReasumePath = (e.Row.FindControl("aReasumePath") as System.Web.UI.HtmlControls.HtmlAnchor);

                    string Status = Convert.ToString((e.Row.FindControl("lblStatus") as HiddenField).Value);//Select the status in DropDownList
                    string orderStatus = Convert.ToString((e.Row.FindControl("lblOrderStatus") as HiddenField).Value);

                    //==TODO Removing Prefix - Need to fix from Creat User page.
                    if (aReasumePath.InnerText.Trim() != "" && aReasumePath.InnerText.Length > 12)
                        aReasumePath.InnerText = aReasumePath.InnerText.Substring(14, aReasumePath.InnerText.Length - 14);

                    char chaDelimiter = '$';

                    if (lblPrimaryPhone.Text.IndexOf(chaDelimiter) > 0)
                        lblPrimaryPhone = ManiPulatePrimaryPhone(lblPrimaryPhone, chaDelimiter);


                    if (Status != "")
                    {
                        ddlStatus.Items.FindByValue(Status).Selected = true;

                        switch ((JGConstant.InstallUserStatus)Convert.ToByte(Status))
                        {
                            case JGConstant.InstallUserStatus.Applicant:
                                {
                                    e.Row.Attributes["style"] = "background-color: #FFFF00";
                                    break;
                                }
                            case JGConstant.InstallUserStatus.InstallProspect:
                                {
                                    e.Row.Attributes["style"] = "background-color: #FFA500";
                                    break;
                                }
                            case JGConstant.InstallUserStatus.InterviewDate:
                                {
                                    if (!string.IsNullOrEmpty(DataBinder.Eval(e.Row.DataItem, "TechTaskId").ToString()))
                                    {
                                        string strParentTechTaskId = Convert.ToString(DataBinder.Eval(e.Row.DataItem, "ParentTechTaskId"));
                                        if (string.IsNullOrEmpty(strParentTechTaskId))
                                        {
                                            strParentTechTaskId = Convert.ToString(DataBinder.Eval(e.Row.DataItem, "TechTaskId"));
                                            hypTechTask.Text = string.Concat(
                                                                                "TaskID#",
                                                                                string.IsNullOrEmpty(DataBinder.Eval(e.Row.DataItem, "TechTaskInstallId").ToString()) ?
                                                                                DataBinder.Eval(e.Row.DataItem, "TechTaskId") :
                                                                                DataBinder.Eval(e.Row.DataItem, "TechTaskInstallId")
                                                                            );
                                        }
                                        else
                                        {
                                            hypTechTask.Text = string.Concat(
                                                                                "SubTaskID#",
                                                                                string.IsNullOrEmpty(DataBinder.Eval(e.Row.DataItem, "TechTaskInstallId").ToString()) ?
                                                                                DataBinder.Eval(e.Row.DataItem, "TechTaskId") :
                                                                                DataBinder.Eval(e.Row.DataItem, "TechTaskInstallId")
                                                                            );
                                        }

                                        hypTechTask.NavigateUrl = string.Format(
                                                                                Page.ResolveUrl("~/Sr_App/TaskGenerator.aspx?TaskId={0}&hstid={1}"),
                                                                                strParentTechTaskId,
                                                                                DataBinder.Eval(e.Row.DataItem, "TechTaskId")
                                                                               );
                                        hypTechTask.Visible = true;
                                    }
                                    break;
                                }
                            case JGConstant.InstallUserStatus.Rejected:
                                {
                                    e.Row.Attributes["style"] = "background-color: #AEAEAE";
                                    break;
                                }
                            case JGConstant.InstallUserStatus.Deactive:
                            case JGConstant.InstallUserStatus.Deleted:
                                {
                                    e.Row.Attributes["style"] = "background-color: #565656";
                                    break;
                                }
                            default:
                                break;
                        }
                    }


                    if (JGSession.DesignationId == (int)JGConstant.DesignationType.Admin || JGSession.DesignationId == (int)JGConstant.DesignationType.IT_Lead)
                    {
                        lnkDelete.Visible = true;
                    }
                    else
                    {
                        lnkDelete.Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                Response.Write("" + ex.Message);
            }
        }

        protected void grdUsers_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            string str = ConfigurationManager.ConnectionStrings["JGPA"].ConnectionString;
            SqlConnection con = new SqlConnection(str);

            if (e.CommandName == "AddNewContact")
            {
                GridViewRow gvRow = (GridViewRow)((Control)e.CommandSource).NamingContainer;
                TextBox txtNewContact = (TextBox)gvRow.FindControl("txtNewContact");
                CheckBox chkIsPrimaryPhone = (CheckBox)gvRow.FindControl("chkIsPrimaryPhone");
                DropDownList ddlContactType = (DropDownList)gvRow.FindControl("ddlContactType");
                //int Index = gvRow.RowIndex;
                bool IsPrimary = chkIsPrimaryPhone.Checked;

                String PhoneType = ddlContactType.SelectedItem.Text;
                int id = Convert.ToInt32(e.CommandArgument);

                if (txtNewContact.Text.Trim() != "")
                {
                    if (PhoneType == "EMAIL")
                    {
                        string strReturnValue = new_customerBLL.Instance.CheckDuplicateSalesUser(txtNewContact.Text, 2, id, 0);
                        if (strReturnValue != "")
                        {
                            //ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('User with email already Exist')", true);
                            ScriptManager.RegisterStartupScript(this, this.GetType(), "alertForEmail", "TheConfirm_OkOnly('User with this email already Exist','Email Alert')", true);
                        }
                        else
                        {
                            InstallUserBLL.Instance.AddNewEmailForUser(txtNewContact.Text, IsPrimary, id);
                        }
                        //binddata();
                        GetSalesUsersStaticticsAndData();
                    }
                    else
                    {
                        string strReturnValue = new_customerBLL.Instance.CheckDuplicateSalesUser(txtNewContact.Text, 1, id, 0);
                        if (strReturnValue != "")
                        {
                            ScriptManager.RegisterStartupScript(this, this.GetType(), "alertForEmail", "TheConfirm_OkOnly('User with this Phone already Exist','Phone Alert')", true);
                        }
                        else
                        {
                            InstallUserBLL.Instance.AddUserPhone(IsPrimary, txtNewContact.Text, Convert.ToInt32(ddlContactType.SelectedValue), id, null, null, false);
                            //binddata();
                            GetSalesUsersStaticticsAndData();
                        }
                    }
                }
                else
                {
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "alertForEmailPhone", "TheConfirm_OkOnly('Kindly Enter Phone / Email Value (It can not be blank)','Alert')", true);
                }

                //DropDownList ddlStatus = (DropDownList)gvRow.FindControl("ddlStatus");
                //int StatusId = Convert.ToInt32(e.CommandArgument);
                ////string Status = ddlStatus.SelectedValue;
                //bool result = InstallUserBLL.Instance.UpdateInstallUserStatus(Status, StatusId);

            }
            else if (e.CommandName == "EditSalesUser")
            {
                //GridViewRow row = (GridViewRow)((Control)e.CommandSource).NamingContainer;
                //int index = row.RowIndex;
                //Label desig = (Label)(grdUsers.Rows[index].Cells[4].FindControl("lblDesignation"));
                //string designation = desig.Text;
                //string ID1 = e.CommandArgument.ToString();
                //con.Open();
                //SqlCommand cmd = new SqlCommand("select Usertype from tblInstallUsers where Id='" + ID1 + "' ", con);
                //SqlDataReader rdr = cmd.ExecuteReader();
                //string type = "";
                //while (rdr.Read())
                //{
                //    type = rdr[0].ToString();

                //}
                //con.Close();
                //if (designation != "SubContractor" && type != "Sales")
                //{
                //    string ID = e.CommandArgument.ToString();
                //    Response.Redirect("InstallCreateUser.aspx?id=" + ID);
                //}
                //else if (designation == "SubContractor" && type != "Sales")
                //{
                //    string ID = e.CommandArgument.ToString();
                //    Response.Redirect("InstallCreateUser2.aspx?id=" + ID);
                //}
                //else if (type == "Sales")
                //{
                string ID = e.CommandArgument.ToString();
                Response.Redirect("ViewSalesUser.aspx?id=" + ID);
                //}

            }
            else if (e.CommandName == "DeactivateSalesUser")
            {
                List<int> lstIds = new List<int>() { Convert.ToInt32(e.CommandArgument.ToString()) };
                if (InstallUserBLL.Instance.DeactivateInstallUsers(lstIds))
                {
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "AlertBox", "alert('User Deactivated Successfully');", true);
                    GetSalesUsersStaticticsAndData();
                }
            }
            else if (e.CommandName == "DeleteSalesUser")
            {
                List<int> lstIds = new List<int>() { Convert.ToInt32(e.CommandArgument.ToString()) };
                if (InstallUserBLL.Instance.DeleteInstallUsers(lstIds))
                {
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "AlertBox", "alert('User Deleted Successfully');", true);
                    GetSalesUsersStaticticsAndData();
                }
            }
            else if (e.CommandName == "ShowPicture")
            {
                string ImagePath = "";
                string ImageName = e.CommandArgument.ToString();
                ImagePath = "UploadedFile/" + Path.GetFileName(ImageName);
                img_InstallerImage.ImageUrl = ImagePath;
                mp1.Show();
            }
            else if (e.CommandName == "ChangeStatus")
            {
                GridViewRow gvRow = (GridViewRow)((Control)e.CommandSource).NamingContainer;
                int Index = gvRow.RowIndex;
                DropDownList ddlStatus = (DropDownList)gvRow.FindControl("ddlStatus");
                int StatusId = Convert.ToInt32(e.CommandArgument);
                string Status = ddlStatus.SelectedValue;
                bool result = InstallUserBLL.Instance.UpdateInstallUserStatus(Status, StatusId);
            }
            else if (e.CommandName == "EditAddedByUserInstall")
            {
                //GridViewRow row = (GridViewRow)((Control)e.CommandSource).NamingContainer;
                //int index = row.RowIndex;
                ////Label desig = (Label)(grdUsers.Rows[index].Cells[4].FindControl("lblDesignation"));
                ////string designation = desig.Text;
                //string ID1 = e.CommandArgument.ToString();
                //con.Open();
                //SqlCommand cmd = new SqlCommand("select Usertype from tblInstallUsers where Id='" + ID1 + "' ", con);
                //SqlDataReader rdr = cmd.ExecuteReader();
                //string type = "";
                //while (rdr.Read())
                //{
                //    type = rdr[0].ToString();

                //}
                //con.Close();
                //if (designation != "SubContractor" && type != "Sales")
                //{
                //    string ID = e.CommandArgument.ToString();
                //    Response.Redirect("InstallCreateUser.aspx?id=" + ID);
                //}
                //else if (designation == "SubContractor" && type != "Sales")
                //{
                //    string ID = e.CommandArgument.ToString();
                //    Response.Redirect("InstallCreateUser2.aspx?id=" + ID);
                //}
                //else if (type == "Sales")
                //{
                string AddedById = e.CommandArgument.ToString();
                Response.Redirect("ViewSalesUser.aspx?id=" + AddedById);
                //}

            }
            else if (e.CommandName == "send-email")
            {
                LoadEmailContentToSentToUser(grdUsers.Rows[Convert.ToInt32(e.CommandArgument)]);
            }
            else if (e.CommandName == "BookmarkUser")
            {
                BookmarkUnStarInstallUser(Convert.ToInt32(e.CommandArgument), Convert.ToInt32(Session["UserId"]), "BookmarkUser");
                GetSalesUsersStaticticsAndData();
            }
            else if (e.CommandName == "UnStarUser")
            {
                BookmarkUnStarInstallUser(Convert.ToInt32(e.CommandArgument), Convert.ToInt32(Session["UserId"]), "UnStarUser");
                GetSalesUsersStaticticsAndData();
            }
        }

        protected void grdUsers_Sorting(object sender, GridViewSortEventArgs e)
        {
            if (this.SalesUserSortExpression == e.SortExpression)
            {
                if (this.SalesUserSortDirection == SortDirection.Ascending)
                {
                    this.SalesUserSortDirection = SortDirection.Descending;
                }
                else
                {
                    this.SalesUserSortDirection = SortDirection.Ascending;
                }
            }
            else
            {
                this.SalesUserSortExpression = e.SortExpression;
                this.SalesUserSortDirection = SortDirection.Ascending;
            }

            //binddata();
            GetSalesUsersStaticticsAndData();
        }

        protected void ddlPageSize_grdUsers_SelectedIndexChanged(object sender, EventArgs e)
        {
            grdUsers.PageSize = Convert.ToInt32(ddlPageSize_grdUsers.SelectedValue);
            grdUsers.PageIndex = 0;
            GetSalesUsersStaticticsAndData();
        }

        protected void grdUsers_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            grdUsers.PageIndex = e.NewPageIndex;
            GetSalesUsersStaticticsAndData();
        }

        protected void grdUsers_ddlStatus_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Below 4 lines is to get that particular row control values
            DropDownList ddlNew = sender as DropDownList;
            string strddlNew = ddlNew.SelectedValue;
            GridViewRow grow = (GridViewRow)((Control)sender).NamingContainer;
            Label lblDesignation = (Label)(grow.FindControl("lblDesignation"));
            Label lblFirstName = (Label)(grow.FindControl("lblFirstName"));
            Label lblLastName = (Label)(grow.FindControl("lblLastName"));
            LinkButton lbtnEmail = (LinkButton)(grow.FindControl("lbtnEmail"));
            HiddenField lblStatus = (HiddenField)(grow.FindControl("lblStatus"));
            Label Id = (Label)grow.FindControl("lblid");
            DropDownList ddl = (DropDownList)grow.FindControl("ddlStatus");
            Session["EditId"] = Id.Text;
            Session["EditStatus"] = ddl.SelectedValue;
            Session["DesignitionSC"] = lblDesignation.Text;
            Session["DesignitionIdSC"] = grdUsers.DataKeys[grow.RowIndex]["DesignationID"].ToString();
            Session["FirstNameNewSC"] = lblFirstName.Text;
            Session["LastNameNewSC"] = lblLastName.Text;

            lblName_InterviewDetails.Text =
            lblName_OfferMade.Text = lblFirstName.Text + " " + lblLastName.Text;

            lblDesignation_OfferMade.Text = lblDesignation.Text;

            if (ddlDesignationForTask.Items.FindByText(lblDesignation.Text) != null)
            {
                ddlDesignationForTask.ClearSelection();
                ddlDesignationForTask.Items.FindByText(lblDesignation.Text).Selected = true;
            }

            if (
                    (lblStatus.Value == Convert.ToByte(JGConstant.InstallUserStatus.Active).ToString()) &&
                    (
                        !(Convert.ToString(Session["usertype"]).Contains("Admin")) &&
                        !(Convert.ToString(Session["usertype"]).Contains("SM"))
                    )
                )
            {
                //binddata();
                GetSalesUsersStaticticsAndData();
                ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('You dont have rights change the status.')", true);
                return;
            }
            else if (
                        (
                            lblStatus.Value == Convert.ToByte(JGConstant.InstallUserStatus.Active).ToString() &&
                            ddl.SelectedValue != Convert.ToByte(JGConstant.InstallUserStatus.Deactive).ToString()
                        ) &&
                        (
                            (Convert.ToString(Session["usertype"]).Contains("Admin")) ||
                            (Convert.ToString(Session["usertype"]).Contains("SM"))
                        )
                    )
            {
                ScriptManager.RegisterStartupScript(this, this.GetType(), "Overlay", "overlayPassword();", true);
                return;
            }
            bool status = CheckRequiredFields(ddl.SelectedValue, Convert.ToInt32(Id.Text));
            if (!status)
            {
                if (ddl.SelectedValue == Convert.ToByte(JGConstant.InstallUserStatus.OfferMade).ToString())
                {
                    hdnFirstName.Value = lblFirstName.Text;
                    hdnLastName.Value = lblLastName.Text;
                    txtEmail.Text = lbtnEmail.Text;
                    txtPassword1.Attributes.Add("value", "jmgrove");
                    txtpassword2.Attributes.Add("value", "jmgrove");
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "Overlay", "OverlayPopupOfferMade();", true);
                    return;
                }
                else
                {
                    //binddata();
                    GetSalesUsersStaticticsAndData();
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('Status cannot be changed as required field for selected status are not field')", true);
                    return;
                }
            }

            if (
                    (
                        ddl.SelectedValue == Convert.ToByte(JGConstant.InstallUserStatus.Active).ToString() ||
                        ddl.SelectedValue == Convert.ToByte(JGConstant.InstallUserStatus.Deactive).ToString()
                    ) &&
                    (
                        !(Convert.ToString(Session["usertype"]).Contains("Admin")) &&
                        !(Convert.ToString(Session["usertype"]).Contains("SM"))
                    )
                )
            {
                ddl.SelectedValue = Convert.ToString(lblStatus.Value);
                ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('You dont have permission to Activate or Deactivate user')", true);
                return;
            }
            else if (ddl.SelectedValue == Convert.ToByte(JGConstant.InstallUserStatus.Rejected).ToString())
            {
                ScriptManager.RegisterStartupScript(this, this.GetType(), "Overlay", "overlay()", true);
                return;
            }
            else if (ddl.SelectedValue == Convert.ToByte(JGConstant.InstallUserStatus.InterviewDate).ToString())
            {
                LoadUsersByRecruiterDesgination(ddlUsers);
                FillTechTaskDropDown(ddlTechTask, ddlTechSubTask, Convert.ToInt32(ddlDesignationForTask.SelectedValue));
                ddlInsteviewtime.DataSource = GetTimeIntervals();
                ddlInsteviewtime.DataBind();
                dtInterviewDate.Text = DateTime.Now.AddDays(1).ToShortDateString();
                ddlInsteviewtime.SelectedValue = "10:00";
                ScriptManager.RegisterStartupScript(this, this.GetType(), "Overlay", "overlayInterviewDate()", true);
                return;
            }
            else if (
                        ddl.SelectedValue == Convert.ToByte(JGConstant.InstallUserStatus.Deactive).ToString() &&
                        (
                            (Convert.ToString(Session["usertype"]).Contains("Admin")) &&
                            (Convert.ToString(Session["usertype"]).Contains("SM"))
                        )
                    )
            {
                Session["DeactivationStatus"] = "Deactive";
                ScriptManager.RegisterStartupScript(this, this.GetType(), "Overlay", "overlay()", true);
                return;
            }
            else if (ddl.SelectedValue == Convert.ToByte(JGConstant.InstallUserStatus.OfferMade).ToString())
            {
                txtEmail.Text = lbtnEmail.Text;
                txtPassword1.Attributes.Add("value", "jmgrove");
                txtpassword2.Attributes.Add("value", "jmgrove");
                ScriptManager.RegisterStartupScript(this, this.GetType(), "Overlay", "OverlayPopupOfferMade();", true);
                return;
            }

            if (ddl.SelectedValue == Convert.ToByte(JGConstant.InstallUserStatus.InstallProspect).ToString())
            {
                if (lblStatus.Value != "")
                {
                    ddl.SelectedValue = lblStatus.Value;
                }
                ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('Status cannot be changed to Install Prospect')", true);
                return;
            }

            if (lblStatus.Value == Convert.ToByte(JGConstant.InstallUserStatus.Active).ToString() &&
                (!(Convert.ToString(Session["usertype"]).Contains("Admin"))))
            {
                ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('Status cannot be changed to any other status other than Deactive once user is Active')", true);
                if (Convert.ToString(Session["PreviousStatusNew"]) != "")
                {
                    ddl.SelectedValue = Convert.ToString(Session["PreviousStatusNew"]);
                }
                return;
            }
            else
            {
                // Adding a popUp...

                InstallUserBLL.Instance.ChangeStatus(Convert.ToString(Session["EditStatus"]), Convert.ToInt32(Session["EditId"]), DateTime.Today, DateTime.Now.ToShortTimeString(), Convert.ToInt32(Session[JG_Prospect.Common.SessionKey.Key.UserId.ToString()]), JGSession.IsInstallUser.Value, txtReason.Text);
                //binddata();
                GetSalesUsersStaticticsAndData();

                if ((ddl.SelectedValue == Convert.ToByte(JGConstant.InstallUserStatus.Active).ToString()) ||
                    (ddl.SelectedValue == Convert.ToByte(JGConstant.InstallUserStatus.Deactive).ToString()))
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "Overlay", "showStatusChangePopUp();", true);
                return;
            }

            //else
            //{

            //    int StatusId = Convert.ToInt32(Id.Text);
            //    string Status = ddl.SelectedValue;
            //    //bool result = InstallUserBLL.Instance.UpdateInstallUserStatus(Status, StatusId);
            //    InstallUserBLL.Instance.ChangeStatus(Status, StatusId, Convert.ToString(DateTime.Today.ToShortDateString()), DateTime.Now.ToShortTimeString(), Convert.ToInt32(Session[JG_Prospect.Common.SessionKey.Key.UserId.ToString()]));
            //}

            //call: updateStauts() function to update it in database.
        }

        private DropDownList BindContactDllForGrid(DropDownList ddlContactType)
        {
            // To Avoid multi call to DB
            if (ViewState["ContactDllForGrid"] != null)
            {   // Bind dropdown 
                ddlContactType = BindContactDllForVS(ddlContactType);
            }
            else
            {
                // Fill ViewState from DB.
                DataSet dsPhoneType;

                dsPhoneType = InstallUserBLL.Instance.GetAllUserPhoneType();

                foreach (DataRow RowItem in dsPhoneType.Tables[0].Rows)
                {
                    if (RowItem["ContactName"].ToString().ToUpper() == "OTHER")
                    {
                        RowItem["ContactName"] = "EMAIL";
                        RowItem["ContactValue"] = "EMAIL";
                        RowItem["UserContactID"] = "0";
                    }
                }

                ViewState["ContactDllForGrid"] = dsPhoneType;

                ddlContactType = BindContactDllForVS(ddlContactType);

            }
            return ddlContactType;
        }

        /// <summary>
        /// Bind DropDown on from the ViewState.
        /// </summary>
        /// <param name="ddlContactType"></param>
        /// <returns></returns>
        private DropDownList BindContactDllForVS(DropDownList ddlContactType)
        {
            DataSet dsPhoneType;
            dsPhoneType = (DataSet)ViewState["ContactDllForGrid"];

            if (dsPhoneType.Tables[0].Rows.Count > 0)
            {
                ddlContactType.DataSource = dsPhoneType.Tables[0];
                ddlContactType.DataTextField = "ContactName";
                ddlContactType.DataValueField = "UserContactID";
                ddlContactType.DataBind();
            }

            return ddlContactType;
        }

        /// <summary>
        /// If PromaryPhone with Phone Type will Inject Phone Type Image.
        /// </summary>
        /// <param name="lblPrimaryPhone"></param>
        /// <param name="chaDelimiter"></param>
        /// <returns></returns>
        private Label ManiPulatePrimaryPhone(Label lblPrimaryPhone, char chaDelimiter)
        {
            string[] strPrimaryPhone = lblPrimaryPhone.Text.Split(chaDelimiter);
            string strPhoneType = "";

            if (!string.IsNullOrEmpty(strPrimaryPhone[1]))
            {
                strPhoneType = strPrimaryPhone[1].ToString().Trim();

                switch (strPhoneType)
                {
                    case "skype":
                        strPhoneType = "../Sr_App/img/skype.png";
                        break;
                    case "whatsapp":
                        strPhoneType = "../Sr_App/img/WhatsApp.png";
                        break;
                    case "HousePhone":
                    case "House Phone":
                        strPhoneType = "../Sr_App/img/Phone_home.png";
                        break;
                    case "CellPhone":
                    case "Cell Phone":
                        strPhoneType = "../Sr_App/img/Cell_Phone.png";
                        break;
                    case "WorkPhone":
                    case "Work Phone":
                        strPhoneType = "../Sr_App/img/WorkPhone.png";
                        break;
                    case "AltPhone":
                    case "Alt. Phone":
                        strPhoneType = "../Sr_App/img/AltPhone.png";
                        break;
                    default:
                        strPhoneType = "../Sr_App/img/WorkPhone.png";
                        break;
                }
                strPhoneType = "<img src='" + strPhoneType + "' alt='' />";
                lblPrimaryPhone.Text = "<a style='color:red' class='PrimaryPhone'>" + strPrimaryPhone[0] + "</a>" + strPhoneType;
            }

            return lblPrimaryPhone;
        }

        #endregion

        #region grdUsers - Popups

        protected void btnChangeStatus_Click(object sender, EventArgs e)
        {
            int isvaliduser = 0;
            isvaliduser = UserBLL.Instance.chklogin(Convert.ToString(Session["loginid"]), txtPassword.Text);
            if (isvaliduser > 0)
            {
                InstallUserBLL.Instance.ChangeStatus(Convert.ToString(Session["EditStatus"]), Convert.ToInt32(Session["EditId"]), DateTime.Today, DateTime.Now.ToShortTimeString(), Convert.ToInt32(Session[JG_Prospect.Common.SessionKey.Key.UserId.ToString()]), JGSession.IsInstallUser.Value, txtReason.Text);
                //binddata();
                GetSalesUsersStaticticsAndData();
            }
        }

        protected void btnSaveReason_Click(object sender, EventArgs e)
        {
            if (Convert.ToString(Session["DeactivationStatus"]) == "Deactive")
            {
                DataSet ds = new DataSet();
                string email = "";
                string HireDate = "";
                string EmpType = "";
                string PayRates = "";
                ds = InstallUserBLL.Instance.ChangeStatus(Convert.ToString(Session["EditStatus"]), Convert.ToInt32(Session["EditId"]), DateTime.Today, DateTime.Now.ToShortTimeString(), Convert.ToInt32(Session[JG_Prospect.Common.SessionKey.Key.UserId.ToString()]), JGSession.IsInstallUser.Value, txtReason.Text);
                if (ds.Tables.Count > 0)
                {
                    if (ds.Tables[0].Rows.Count > 0)
                    {
                        if (Convert.ToString(ds.Tables[0].Rows[0][0]) != "")
                        {
                            email = Convert.ToString(ds.Tables[0].Rows[0][0]);
                        }
                        if (Convert.ToString(ds.Tables[0].Rows[0][1]) != "")
                        {
                            HireDate = Convert.ToString(ds.Tables[0].Rows[0][1]);
                        }
                        if (Convert.ToString(ds.Tables[0].Rows[0][2]) != "")
                        {
                            EmpType = Convert.ToString(ds.Tables[0].Rows[0][2]);
                        }
                        if (Convert.ToString(ds.Tables[0].Rows[0][3]) != "")
                        {
                            PayRates = Convert.ToString(ds.Tables[0].Rows[0][3]);
                        }
                    }
                }
                SendEmail(email, Convert.ToString(Session["FirstNameNewSC"]), Convert.ToString(Session["LastNameNewSC"]), "Deactivation", txtReason.Text, Convert.ToString(Session["DesignitionSC"]), Convert.ToInt32(Session["DesignitionIdSC"]), HireDate, EmpType, PayRates, 0);
            }
            else
            {
                InstallUserBLL.Instance.ChangeStatus(Convert.ToString(Session["EditStatus"]), Convert.ToInt32(Session["EditId"]), DateTime.Today, DateTime.Now.ToShortTimeString(), Convert.ToInt32(Session[JG_Prospect.Common.SessionKey.Key.UserId.ToString()]), JGSession.IsInstallUser.Value, txtReason.Text);
                //binddata();
                GetSalesUsersStaticticsAndData();
            }
            ScriptManager.RegisterStartupScript(this, this.GetType(), "Overlay", "ClosePopup()", true);
            return;
        }

        protected void ddlDesignationForTask_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillTechTaskDropDown(ddlTechTask, ddlTechSubTask, Convert.ToInt32(ddlDesignationForTask.SelectedValue));
        }

        protected void ddlTechTask_SelectedIndexChanged(object sender, EventArgs e)
        {
            ddlTechSubTask.ClearSelection();
            ddlTechSubTask.Items.Clear();

            if (ddlTechTask.SelectedIndex > 0)
            {
                //DataSet dsSubTasks = TaskGeneratorBLL.Instance.GetSubTasks(Convert.ToInt32(ddlTechTask.SelectedValue), CommonFunction.CheckAdminAndItLeadMode(), "Title ASC", "", null, null);
                //ddlTechSubTask.DataSource = dsSubTasks.Tables[0];
                //ddlTechSubTask.DataTextField = "Title";
                //ddlTechSubTask.DataValueField = "TaskId";
                //ddlTechSubTask.DataBind();
                FillSubTaskDropDown(Convert.ToInt64(ddlTechTask.SelectedValue), ddlTechSubTask);
            }
            else
            {
                ddlTechSubTask.DataSource = null;
                ddlTechSubTask.DataTextField = "Title";
                ddlTechSubTask.DataValueField = "TaskId";
                ddlTechSubTask.DataBind();
            }
            ddlTechSubTask.Items.Insert(0, new ListItem("--select--", "0"));
            ddlTechSubTask.SelectedValue = "0";
        }

        protected void btnSaveInterview_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            string email = "";
            string HireDate = "";
            string EmpType = "";
            string PayRates = "";
            string gitusername = string.Empty;


            //string InterviewDate = dtInterviewDate.Text;
            DateTime interviewDate;
            DateTime.TryParse(dtInterviewDate.Text, out interviewDate);
            if (interviewDate == null)
            {
                ScriptManager.RegisterStartupScript(this, this.GetType(), "Overlay", "alert('Invalid Interview Date, Please verify');", true);
                return;
            }
            ds = InstallUserBLL.Instance.ChangeStatus(Convert.ToString(Session["EditStatus"]), Convert.ToInt32(Session["EditId"]), interviewDate, ddlInsteviewtime.SelectedItem.Text, Convert.ToInt32(Session[JG_Prospect.Common.SessionKey.Key.UserId.ToString()]), JGSession.IsInstallUser.Value, txtReason.Text, ddlUsers.SelectedValue);
            if (ds.Tables.Count > 0)
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    if (Convert.ToString(ds.Tables[0].Rows[0][0]) != "")
                    {
                        email = Convert.ToString(ds.Tables[0].Rows[0][0]);
                    }
                    if (Convert.ToString(ds.Tables[0].Rows[0][1]) != "")
                    {
                        HireDate = Convert.ToString(ds.Tables[0].Rows[0][1]);
                    }
                    if (Convert.ToString(ds.Tables[0].Rows[0][2]) != "")
                    {
                        EmpType = Convert.ToString(ds.Tables[0].Rows[0][2]);
                    }
                    if (Convert.ToString(ds.Tables[0].Rows[0][3]) != "")
                    {
                        PayRates = Convert.ToString(ds.Tables[0].Rows[0][3]);
                    }
                    if (ds.Tables[0].Rows[0]["GitUserName"] != null)
                    {
                        gitusername = ds.Tables[0].Rows[0]["GitUserName"].ToString();
                    }
                }
            }

            SendEmail(email, Convert.ToString(Session["FirstNameNewSC"]), Convert.ToString(Session["LastNameNewSC"]),
                "Interview Date Auto Email", txtReason.Text, Convert.ToString(Session["DesignitionSC"]).Trim(), Convert.ToInt32(Session["DesignitionIdSC"]), HireDate, EmpType, PayRates, HTMLTemplates.InterviewDateAutoEmail
                , null, ddlUsers.SelectedItem != null ? ddlUsers.SelectedItem.Text : "");
            if (!string.IsNullOrEmpty(gitusername))
            {
                AddUserAsGitcollaborator(gitusername);
            }

            //AssignedTask if any or Default
            AssignedTaskToUser(Convert.ToInt32(Session["EditId"]), ddlTechTask, ddlTechSubTask);

            Response.Redirect(JG_Prospect.Common.JGConstant.PG_PATH_MASTER_CALENDAR);

            //binddata();
            //ScriptManager.RegisterStartupScript(this, this.GetType(), "Overlay", "ClosePopupInterviewDate()", true);
            //return;
        }

        protected void btnCancelInterview_Click(object sender, EventArgs e)
        {
            //binddata();
            GetSalesUsersStaticticsAndData();
        }

        protected void btnSaveOfferMade_Click(object sender, EventArgs e)
        {
            int EditId = 0;
            int.TryParse(Convert.ToString(Session["EditId"]), out EditId);
            InstallUserBLL.Instance.UpdateOfferMade(EditId, txtEmail.Text, txtPassword1.Text);

            DataSet ds = new DataSet();
            string email, HireDate, EmpType, PayRates, Desig, LastName, Address, FirstName;
            email = HireDate = EmpType = PayRates = Desig = LastName = Address = FirstName = String.Empty;

            ds = InstallUserBLL.Instance.ChangeStatus(Convert.ToString(Session["EditStatus"]), EditId, DateTime.Today, DateTime.Now.ToShortTimeString(), Convert.ToInt32(Session[JG_Prospect.Common.SessionKey.Key.UserId.ToString()]), JGSession.IsInstallUser.Value, txtReason.Text);
            if (ds.Tables.Count > 0)
            {
                if (ds.Tables[0].Rows.Count > 0)
                {
                    if (Convert.ToString(ds.Tables[0].Rows[0]["Email"]) != "")
                    {
                        email = Convert.ToString(ds.Tables[0].Rows[0]["Email"]);
                    }
                    if (Convert.ToString(ds.Tables[0].Rows[0]["HireDate"]) != "")
                    {
                        HireDate = Convert.ToString(ds.Tables[0].Rows[0]["HireDate"]);
                    }
                    if (Convert.ToString(ds.Tables[0].Rows[0]["EmpType"]) != "")
                    {
                        EmpType = Convert.ToString(ds.Tables[0].Rows[0]["EmpType"]);
                    }
                    if (Convert.ToString(ds.Tables[0].Rows[0]["PayRates"]) != "")
                    {
                        PayRates = Convert.ToString(ds.Tables[0].Rows[0]["PayRates"]);
                    }
                    if (Convert.ToString(ds.Tables[0].Rows[0]["Designation"]) != "")
                    {
                        Desig = Convert.ToString(ds.Tables[0].Rows[0]["Designation"]);
                    }
                    if (!String.IsNullOrEmpty(ds.Tables[0].Rows[0]["FristName"].ToString()))
                    {
                        FirstName = ds.Tables[0].Rows[0]["FristName"].ToString();
                    }
                    if (!String.IsNullOrEmpty(ds.Tables[0].Rows[0]["LastName"].ToString()))
                    {
                        LastName = ds.Tables[0].Rows[0]["LastName"].ToString();
                    }
                    if (!String.IsNullOrEmpty(ds.Tables[0].Rows[0]["Address"].ToString()))
                    {
                        Address = ds.Tables[0].Rows[0]["Address"].ToString();
                    }
                }
            }
            //string strHtml = JG_Prospect.App_Code.CommonFunction.GetContractTemplateContent(199, 0, Desig);
            DesignationHTMLTemplate objHTMLTemplate = HTMLTemplateBLL.Instance.GetDesignationHTMLTemplate(HTMLTemplates.Offer_Made_Attachment_Template, Convert.ToString(Session["DesignitionIdSC"]));
            string strHtml = objHTMLTemplate.Header + objHTMLTemplate.Body + objHTMLTemplate.Footer;
            strHtml = strHtml.Replace("#CurrentDate#", DateTime.Now.ToShortDateString());
            strHtml = strHtml.Replace("#FirstName#", FirstName);
            strHtml = strHtml.Replace("#LastName#", LastName);
            strHtml = strHtml.Replace("#Address#", Address);
            strHtml = strHtml.Replace("#Designation#", Desig);
            if (!string.IsNullOrEmpty(EmpType))
            {
                int intEmpType = 0;
                int.TryParse(EmpType, out intEmpType);

                if (intEmpType > 0)
                {
                    EmpType = CommonFunction.GetEnumDescription((JGConstant.EmploymentType)intEmpType);
                }

                strHtml = strHtml.Replace("#EmpType#", EmpType);

            }
            else
            {
                strHtml = strHtml.Replace("#EmpType#", "________________");
            }
            strHtml = strHtml.Replace("#JoiningDate#", HireDate);
            if (!string.IsNullOrEmpty(PayRates))
            {
                strHtml = strHtml.Replace("#RatePerHour#", PayRates);
            }
            else
            {
                strHtml = strHtml.Replace("#RatePerHour#", "____");
            }
            DateTime dtPayCheckDate;
            if (!string.IsNullOrEmpty(HireDate))
            {
                dtPayCheckDate = Convert.ToDateTime(HireDate);
            }
            else
            {
                dtPayCheckDate = DateTime.Now;
            }
            dtPayCheckDate = new DateTime(dtPayCheckDate.Year, dtPayCheckDate.Month, DateTime.DaysInMonth(dtPayCheckDate.Year, dtPayCheckDate.Month));
            strHtml = strHtml.Replace("#PayCheckDate#", dtPayCheckDate.ToShortDateString());

            string strPath = JG_Prospect.App_Code.CommonFunction.ConvertHtmlToPdf(strHtml, Server.MapPath(@"~\Sr_App\MailDocument\MailAttachments\"), "Job acceptance letter");
            List<Attachment> lstAttachments = new List<Attachment>();
            if (File.Exists(strPath))
            {
                Attachment attachment = new Attachment(strPath);
                attachment.Name = Path.GetFileName(strPath);
                lstAttachments.Add(attachment);
            }

            SendEmail(email, FirstName, LastName, "Offer Made", txtReason.Text, Desig, Convert.ToInt32(Session["DesignitionIdSC"]), HireDate, EmpType, PayRates,
                HTMLTemplates.Offer_Made_Auto_Email, lstAttachments);

            //binddata();
            GetSalesUsersStaticticsAndData();
            ScriptManager.RegisterStartupScript(this, this.GetType(), "Overlay", "ClosePopupOfferMade()", true);
            return;
        }

        protected void btnSendEmailToUser_Click(object sender, EventArgs e)
        {
            DataSet ds = AdminBLL.Instance.GetEmailTemplate(JGSession.Designation, 110);

            if (ds == null)
            {
                ds = AdminBLL.Instance.GetEmailTemplate("Admin");
            }
            else if (ds.Tables[0].Rows.Count == 0)
            {
                ds = AdminBLL.Instance.GetEmailTemplate("Admin");
            }

            List<Attachment> lstAttachments = new List<Attachment>();

            for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
            {
                string sourceDir = Server.MapPath(ds.Tables[1].Rows[i]["DocumentPath"].ToString());
                if (File.Exists(sourceDir))
                {
                    Attachment attachment = new Attachment(sourceDir);
                    attachment.Name = Path.GetFileName(sourceDir);
                    lstAttachments.Add(attachment);
                }
            }

            string strBody = txtEmailBody.Text + txtEmailCustomMessage.Text;

            try
            {
                JG_Prospect.App_Code.CommonFunction.SendEmail(JGSession.Designation, hdnEmailTo.Value, txtEmailSubject.Text, strBody, lstAttachments);

                ScriptManager.RegisterStartupScript(
                                                    this,
                                                    this.GetType(),
                                                    "HidePopup_divSendEmailToUserWithAlert",
                                                    string.Concat
                                                                (
                                                                    "alert('An email notification has sent on " + hdnEmailTo.Value + ".');",
                                                                    "HidePopup('#", divSendEmailToUser.ClientID, "');"
                                                                ),
                                                    true
                                                   );
            }
            catch
            {
                ScriptManager.RegisterStartupScript(this, this.GetType(), "UserMsg", "alert('Error while sending email notification on " + hdnEmailTo.Value + ".');", true);
            }
        }

        protected void btnCancelSendEmailToUser_Click(object sender, EventArgs e)
        {
            ScriptManager.RegisterStartupScript(
                                                this,
                                                this.GetType(),
                                                "HidePopup_divSendEmailToUser",
                                                string.Concat("HidePopup('#", divSendEmailToUser.ClientID, "');"),
                                                true
                                               );
        }

        #region '----Change Status For Selected - Popup----'

        protected void ddlStatus_Popup_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ddlStatus_Popup.SelectedValue == Convert.ToByte(JGConstant.InstallUserStatus.InterviewDate).ToString())
            {
                divInterviewDate.Visible =
                grdUsers_Popup.Columns[2].Visible =
                grdUsers_Popup.Columns[3].Visible = true;
                grdUsers_Popup.Columns[4].Visible = false;
            }
            else if (ddlStatus_Popup.SelectedValue == Convert.ToByte(JGConstant.InstallUserStatus.OfferMade).ToString())
            {
                divInterviewDate.Visible =
                grdUsers_Popup.Columns[2].Visible =
                grdUsers_Popup.Columns[3].Visible =
                grdUsers_Popup.Columns[4].Visible = false;
            }
            else if (ddlStatus_Popup.SelectedValue == Convert.ToByte(JGConstant.InstallUserStatus.Rejected).ToString())
            {
                divInterviewDate.Visible =
                grdUsers_Popup.Columns[2].Visible =
                grdUsers_Popup.Columns[3].Visible = false;
                grdUsers_Popup.Columns[4].Visible = true;
            }
            else
            {
                divInterviewDate.Visible =
                grdUsers_Popup.Columns[2].Visible =
                grdUsers_Popup.Columns[3].Visible =
                grdUsers_Popup.Columns[4].Visible = false;
            }

            grdUsers_Popup.DataSource = this.SelectedUsers;
            grdUsers_Popup.DataBind();
            upChangeStatusForSelected.Update();
        }

        protected void grdUsers_Popup_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                DropDownList ddlInterviewTime = e.Row.FindControl("ddlInterviewTime") as DropDownList;
                DropDownList ddlTechTask = e.Row.FindControl("ddlTechTask") as DropDownList;
                DropDownList ddlTechSubTask = e.Row.FindControl("ddlTechSubTask") as DropDownList;

                ddlInterviewTime.DataSource = GetTimeIntervals();
                ddlInterviewTime.DataBind();
                ddlInsteviewtime.SelectedValue = DataBinder.Eval(e.Row.DataItem, "InterviewTime").ToString();

                FillTechTaskDropDown(ddlTechTask, ddlTechSubTask, Convert.ToInt32(DataBinder.Eval(e.Row.DataItem, "DesignationID")));
            }
        }

        protected void grdUsers_Popup_ddlTechTask_SelectedIndexChanged(object sender, EventArgs e)
        {
            DropDownList ddlTechTask = sender as DropDownList;
            GridViewRow objGridRow = ddlTechTask.NamingContainer as GridViewRow;
            DropDownList ddlTechSubTask = objGridRow.FindControl("ddlTechSubTask") as DropDownList;
            if (ddlTechTask != null && ddlTechSubTask != null)
            {
                ddlTechSubTask.ClearSelection();
                ddlTechSubTask.Items.Clear();

                if (ddlTechTask.SelectedIndex > 0)
                {
                    //DataSet dsSubTasks = TaskGeneratorBLL.Instance.GetSubTasks(Convert.ToInt32(ddlTechTask.SelectedValue), CommonFunction.CheckAdminAndItLeadMode(), "Title ASC", "", null, null);
                    //ddlTechSubTask.DataSource = dsSubTasks.Tables[0];
                    //ddlTechSubTask.DataTextField = "Title";
                    //ddlTechSubTask.DataValueField = "TaskId";
                    //ddlTechSubTask.DataBind();
                    FillSubTaskDropDown(Convert.ToInt64(ddlTechTask.SelectedValue), ddlTechSubTask);
                }
                else
                {
                    ddlTechSubTask.DataSource = null;
                    ddlTechSubTask.DataTextField = "Title";
                    ddlTechSubTask.DataValueField = "TaskId";
                    ddlTechSubTask.DataBind();
                }
                ddlTechSubTask.Items.Insert(0, new ListItem("--select--", "0"));
                ddlTechSubTask.SelectedValue = "0";
            }
            upChangeStatusForSelected.Update();
        }

        protected void btnCancelChangeStatusForSelected_Click(object sender, EventArgs e)
        {
            ScriptManager.RegisterStartupScript
                                (
                                    this,
                                    this.GetType(),
                                    "HidePopup_divChangeStatusForSelected",
                                    string.Concat("HidePopup('#", divChangeStatusForSelected.ClientID, "');"),
                                    true
                                );
        }

        protected void btnSaveStatusForSelected_Click(object sender, EventArgs e)
        {
            int intId, intDesignationId;
            string strEmail, strHireDate, strEmployeeType, strPayRates, strFirstName, strLastName, strDesignation, strReason, strTime;
            DateTime? dtDate = null;

            foreach (GridViewRow objUserRow in grdUsers_Popup.Rows)
            {
                strEmail =
                strHireDate =
                strEmployeeType =
                strPayRates =
                strFirstName =
                strLastName =
                strDesignation =
                strReason =
                strTime = string.Empty;

                intId = Convert.ToInt32(grdUsers_Popup.DataKeys[objUserRow.RowIndex]["Id"]);
                intDesignationId = Convert.ToInt32(grdUsers_Popup.DataKeys[objUserRow.RowIndex]["DesignationID"]);
                strFirstName = ((Literal)objUserRow.FindControl("ltrlFirstName")).Text;
                strLastName = ((Literal)objUserRow.FindControl("ltrlLastName")).Text;
                strDesignation = ((Literal)objUserRow.FindControl("ltrlDesignation")).Text;
                strReason = ((TextBox)objUserRow.FindControl("txtReason")).Text;
                string strDate = ((TextBox)objUserRow.FindControl("txtInterviewDate")).Text;
                if (!string.IsNullOrEmpty(strDate))
                {
                    dtDate = Convert.ToDateTime(strDate);
                }
                strTime = ((DropDownList)objUserRow.FindControl("ddlInterviewTime")).SelectedValue;

                if (!dtDate.HasValue)
                {
                    dtDate = DateTime.Today;
                }

                if (string.IsNullOrEmpty(strTime))
                {
                    strTime = DateTime.Now.ToShortTimeString();
                }

                DataSet dsUser = InstallUserBLL.Instance.ChangeStatus
                                                    (
                                                        ddlStatus_Popup.SelectedValue,
                                                        intId,
                                                        dtDate.Value,
                                                        strTime,
                                                        JGSession.UserId,
                                                        JGSession.IsInstallUser.Value,
                                                        strReason
                                                    );

                if (dsUser.Tables.Count > 0 && dsUser.Tables[0].Rows.Count > 0)
                {
                    strEmail = Convert.ToString(dsUser.Tables[0].Rows[0][0]);
                    strHireDate = Convert.ToString(dsUser.Tables[0].Rows[0][1]);
                    strEmployeeType = Convert.ToString(dsUser.Tables[0].Rows[0][2]);
                    strPayRates = Convert.ToString(dsUser.Tables[0].Rows[0][3]);
                }

                switch ((JGConstant.InstallUserStatus)Convert.ToByte(ddlStatus_Popup.SelectedValue))
                {
                    case JGConstant.InstallUserStatus.Deactive:
                        SendEmail(
                                    strEmail,
                                    strFirstName,
                                    strLastName,
                                    "Deactivation",
                                    strReason,
                                    strDesignation,
                                    intDesignationId,
                                    strHireDate,
                                    strEmployeeType,
                                    strPayRates,
                                    0
                                );
                        break;

                    case JGConstant.InstallUserStatus.InterviewDate:
                        SendEmail(
                                    strEmail,
                                    strFirstName,
                                    strLastName,
                                    "Interview Date Auto Email",
                                    strReason,
                                    strDesignation,
                                    intDesignationId,
                                    strHireDate,
                                    strEmployeeType,
                                    strPayRates,
                                    HTMLTemplates.InterviewDateAutoEmail,
                                    null,
                                    ddlRecruiter_Popup.SelectedItem != null ? ddlRecruiter_Popup.SelectedItem.Text : ""
                                );

                        //AssignedTask if any or Default
                        AssignedTaskToUser(intId, (DropDownList)objUserRow.FindControl("ddlTechTask"), (DropDownList)objUserRow.FindControl("ddlTechSubTask"));

                        break;

                    case JGConstant.InstallUserStatus.OfferMade:
                        InstallUserBLL.Instance.UpdateOfferMade(intId, strEmail, JGSession.UserPassword);

                        #region '-- PDF Attachment --'

                        //string strHtml = JG_Prospect.App_Code.CommonFunction.GetContractTemplateContent(199, 0, strDesignation);
                        DesignationHTMLTemplate objHTMLTemplate = HTMLTemplateBLL.Instance.GetDesignationHTMLTemplate(HTMLTemplates.Offer_Made_Attachment_Template, intDesignationId.ToString());
                        string strHtml = objHTMLTemplate.Header + objHTMLTemplate.Body + objHTMLTemplate.Footer;
                        strHtml = strHtml.Replace("#CurrentDate#", DateTime.Now.ToShortDateString());
                        strHtml = strHtml.Replace("#FirstName#", strFirstName);
                        strHtml = strHtml.Replace("#LastName#", strLastName);
                        strHtml = strHtml.Replace("#Address#", string.Empty);
                        strHtml = strHtml.Replace("#Designation#", strDesignation);

                        if (!string.IsNullOrEmpty(strEmployeeType) && strEmployeeType.Length > 1)
                        {
                            strHtml = strHtml.Replace("#EmpType#", strEmployeeType);
                        }
                        else
                        {
                            strHtml = strHtml.Replace("#EmpType#", "________________");
                        }
                        strHtml = strHtml.Replace("#JoiningDate#", strHireDate);
                        if (!string.IsNullOrEmpty(strPayRates))
                        {
                            strHtml = strHtml.Replace("#RatePerHour#", strPayRates);
                        }
                        else
                        {
                            strHtml = strHtml.Replace("#RatePerHour#", "____");
                        }
                        DateTime dtPayCheckDate;
                        if (!string.IsNullOrEmpty(strHireDate))
                        {
                            dtPayCheckDate = Convert.ToDateTime(strHireDate);
                        }
                        else
                        {
                            dtPayCheckDate = DateTime.Now;
                        }
                        dtPayCheckDate = new DateTime(dtPayCheckDate.Year, dtPayCheckDate.Month, DateTime.DaysInMonth(dtPayCheckDate.Year, dtPayCheckDate.Month));
                        strHtml = strHtml.Replace("#PayCheckDate#", dtPayCheckDate.ToShortDateString());

                        string strPath = JG_Prospect.App_Code.CommonFunction.ConvertHtmlToPdf
                                                                (
                                                                    strHtml,
                                                                    Server.MapPath(@"~\Sr_App\MailDocument\MailAttachments\"),
                                                                    "Job acceptance letter"
                                                                );
                        List<Attachment> lstAttachments = new List<Attachment>();
                        if (File.Exists(strPath))
                        {
                            Attachment attachment = new Attachment(strPath);
                            attachment.Name = Path.GetFileName(strPath);
                            lstAttachments.Add(attachment);
                        }

                        #endregion

                        SendEmail(
                                    strEmail,
                                    strFirstName,
                                    strLastName,
                                    "Offer Made",
                                    strReason,
                                    strDesignation,
                                    intDesignationId,
                                    strHireDate,
                                    strEmployeeType,
                                    strPayRates,
                                    HTMLTemplates.Offer_Made_Auto_Email,
                                    lstAttachments
                                );
                        break;

                    default:
                        break;
                }
            }

            if (ddlStatus_Popup.SelectedValue == Convert.ToByte(JGConstant.InstallUserStatus.InterviewDate).ToString())
            {
                Response.Redirect(JG_Prospect.Common.JGConstant.PG_PATH_MASTER_CALENDAR);
            }
            else
            {

                GetSalesUsersStaticticsAndData();

                ScriptManager.RegisterStartupScript
                                    (
                                        this,
                                        this.GetType(),
                                        "HidePopup_divChangeStatusForSelected",
                                        string.Concat("HidePopup('#", divChangeStatusForSelected.ClientID, "');"),
                                        true
                                    );
            }
        }

        #endregion

        #endregion

        protected void btnUploadNew_Click(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(hdnBulkUploadFile.Value))
                {
                    string strFilePath = Server.MapPath("~/UploadedExcel/" + hdnBulkUploadFile.Value.Split('^')[0].Split('@')[0]);

                    string FileName = Path.GetFileName(strFilePath);
                    string Extension = Path.GetExtension(strFilePath);

                    if (File.Exists(strFilePath) && (Extension == ".xlsx" || Extension == ".csv"))
                    {
                        DataTable dtExcel = new DataTable();

                        switch (Extension)
                        {
                            case ".xls":
                            case ".xlsx":
                                Stream objStream = File.OpenRead(strFilePath);
                                dtExcel = ReadExcelFile(new ExcelPackage(objStream));
                                break;

                            case ".csv":
                                dtExcel = ReadCsvFile(strFilePath);
                                break;
                        }

                        if (ValidateUploadedData(dtExcel))
                        {
                            ImportIntsallUsers(dtExcel);
                            GetSalesUsersStaticticsAndData();
                        }
                        else
                        {
                            GetSalesUsersStaticticsAndData(); ;
                            UcStatusPopUp.changeText();
                            ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "showStatusChangePopUp();", true);
                        }
                    }
                    else
                    {
                        UcStatusPopUp.ucPopUpHeader = "";
                        UcStatusPopUp.ucPopUpMsg = "Please Select xlsx or csv file.";
                        UcStatusPopUp.changeText();
                    }
                }
            }
            catch (Exception ex)
            {
                UtilityBAL.AddException("EditUser-btnUploadNew_Click", Session["loginid"] == null ? "" : Session["loginid"].ToString(), ex.Message, ex.StackTrace);
            }
        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            try
            {
                if (BulkProspectUploader.HasFile)
                {
                    string ext = Path.GetExtension(BulkProspectUploader.FileName);
                    if (ext == ".xlsx" || ext == ".csv")
                    {
                        string FileName = Path.GetFileName(BulkProspectUploader.PostedFile.FileName);
                        string Extension = Path.GetExtension(BulkProspectUploader.PostedFile.FileName);

                        //string FolderPath = ConfigurationManager.AppSettings["FolderPath"];
                        //string FilePath = Server.MapPath(FolderPath + FileName);

                        string GenFolderPath = DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString() + DateTime.Now.Second.ToString();

                        string directoryPath = Server.MapPath("/UploadedExcel/" + GenFolderPath + "/");
                        if (!Directory.Exists(directoryPath))
                        {
                            Directory.CreateDirectory(directoryPath);
                        }
                        directoryPath = directoryPath + FileName;

                        BulkProspectUploader.SaveAs(directoryPath);

                        DataTable dtExcel = FillValueFromFiles(directoryPath, Extension);

                        if (ValidateUploadedData(dtExcel) == false)
                        {
                            //binddata();
                            GetSalesUsersStaticticsAndData(); ;
                            UcStatusPopUp.changeText();
                            ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "showStatusChangePopUp();", true);
                            return;
                        }
                        else
                        {
                            ImportIntsallUsers(dtExcel);
                            //binddata();
                            GetSalesUsersStaticticsAndData();
                        }
                    }
                    else
                    {
                        UcStatusPopUp.ucPopUpHeader = "";
                        UcStatusPopUp.ucPopUpMsg = "Please Select xlsx or csv file.";
                        UcStatusPopUp.changeText();
                    }
                }
            }
            catch (Exception ex)
            {
                UtilityBAL.AddException("EditUser-btnUpload_Click", Session["loginid"] == null ? "" : Session["loginid"].ToString(), ex.Message, ex.StackTrace);
            }
        }

        protected void btnExport_Click(object sender, EventArgs e)
        {
            /*
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", string.Format("attachment; filename={0}", "Users.xls"));
            Response.ContentType = "application/ms-excel";
            // DataSet ds = InstallUserBLL.Instance.getalluserdetails();
            DataTable dt = (DataTable)(Session["GridDataExport"]);
            // dt.Columns.Remove("PrimeryTradeId");
            // dt.Columns.Remove("SecondoryTradeId");
            string str = string.Empty;
            foreach (DataColumn dtcol in dt.Columns)
            {
                Response.Write(str + dtcol.ColumnName);
                str = "\t";
            }
            Response.Write("\n");
            foreach (DataRow dr in dt.Rows)
            {
                str = "";
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    Response.Write(str + Convert.ToString(dr[j]));
                    str = "\t";
                }
                Response.Write("\n");
            }
            Response.End();
          * */

            DataSet dsUsers_Export = InstallUserBLL.Instance.GetAllSalesUserToExport();
            DataTable dt = dsUsers_Export.Tables[0];

            string filename = "SalesUser.xls";
            System.IO.StringWriter tw = new System.IO.StringWriter();
            System.Web.UI.HtmlTextWriter hw = new System.Web.UI.HtmlTextWriter(tw);
            DataGrid dgGrid = new DataGrid();
            dgGrid.DataSource = dt;
            dgGrid.DataBind();

            //Get the HTML for the control.
            dgGrid.RenderControl(hw);
            //Write the HTML back to the browser.
            //Response.ContentType = application/vnd.ms-excel;
            Response.ContentType = "application/vnd.ms-excel";
            Response.AppendHeader("Content-Disposition", "attachment; filename=" + filename + "");
            this.EnableViewState = false;
            string style = @"<style> .textmode { mso-number-format:\@; } </style>";
            Response.Write(style);
            Response.Write(tw.ToString());
            Response.End();
        }

        protected void btnYesEdit_Click(object sender, EventArgs e)
        {
            if (Session["DuplicateUsers"] != null)
            {
                DataTable dt = (DataTable)Session["DuplicateUsers"];

                XmlDocument xmlDoc = new XmlDocument();
                CreateDuplicateUserObjectXml(dt, out xmlDoc);

                bool result = InstallUserBLL.Instance.BulkUpdateIntsallUser(xmlDoc.InnerXml, Session["loginid"].ToString());

                UcStatusPopUp.ucPopUpMsg = "Data updated successfully.";
                UcStatusPopUp.ucPopUpHeader = "";
                UcStatusPopUp.changeText();

                ScriptManager.RegisterStartupScript(this, this.GetType(), "Overlay", "showStatusChangePopUp();", true);

            }
            else
            {
                ScriptManager.RegisterStartupScript(this, this.GetType(), "AlertBox", "alert('Session expired , Please try again.');", true);
            }






            //if (result)
            //{
            //    ScriptManager.RegisterStartupScript(this, this.GetType(), "AlertBox", "alert('Data updated successfully!');ClosePopupUploadBulk();", true);
            //}
            //else
            //{
            //    ScriptManager.RegisterStartupScript(this, this.GetType(), "AlertBox", "alert('There is some error.');", true);
            //}
        }

        protected void btnNoEdit_Click(object sender, EventArgs e)
        {
            Session["DuplicateUsers"] = null;
            ScriptManager.RegisterStartupScript(this, this.GetType(), "overlay", "ClosePopupUploadBulk();", true);
            return;
        }

        protected void lbtnDeactivateSelected_Click(object sender, EventArgs e)
        {
            List<Int32> lstIDs = new List<int>();

            foreach (GridViewRow objUserRow in grdUsers.Rows)
            {
                if (((CheckBox)objUserRow.FindControl("chkSelected")).Checked)
                {
                    lstIDs.Add(Convert.ToInt32(grdUsers.DataKeys[objUserRow.RowIndex]["Id"]));
                }
            }

            if (lstIDs.Count > 0)
            {
                if (InstallUserBLL.Instance.DeactivateInstallUsers(lstIDs))
                {
                    CommonFunction.ShowAlertFromUpdatePanel(this, "User deactivated Successfully.");
                    GetSalesUsersStaticticsAndData();
                }
                else
                {
                    CommonFunction.ShowAlertFromUpdatePanel(this, "User can not be deactivated. Please try again.");
                }
            }
            else
            {
                CommonFunction.ShowAlertFromUpdatePanel(this, "Please select user(s) to deactivated.");
            }
        }

        protected void lbtnDeleteSelected_Click(object sender, EventArgs e)
        {
            List<Int32> lstIDs = new List<int>();

            foreach (GridViewRow objUserRow in grdUsers.Rows)
            {
                if (((CheckBox)objUserRow.FindControl("chkSelected")).Checked)
                {
                    lstIDs.Add(Convert.ToInt32(grdUsers.DataKeys[objUserRow.RowIndex]["Id"]));
                }
            }

            if (lstIDs.Count > 0)
            {
                if (InstallUserBLL.Instance.DeleteInstallUsers(lstIDs))
                {
                    CommonFunction.ShowAlertFromUpdatePanel(this, "User deleted Successfully.");
                    GetSalesUsersStaticticsAndData();
                }
                else
                {
                    CommonFunction.ShowAlertFromUpdatePanel(this, "User can not be deleted. Please try again.");
                }
            }
            else
            {
                CommonFunction.ShowAlertFromUpdatePanel(this, "Please select user(s) to deleted.");
            }
        }

        protected void lbtnChangeStatusForSelected_Click(object sender, EventArgs e)
        {
            this.SelectedUsers = null;

            DataTable dtUsers = new DataTable();
            dtUsers.Columns.Add("Id");
            dtUsers.Columns.Add("FirstName");
            dtUsers.Columns.Add("LastName");
            dtUsers.Columns.Add("Designation");
            dtUsers.Columns.Add("DesignationID");
            dtUsers.Columns.Add("InterviewDate");
            dtUsers.Columns.Add("InterviewTime");

            foreach (GridViewRow objUserRow in grdUsers.Rows)
            {
                if (((CheckBox)objUserRow.FindControl("chkSelected")).Checked)
                {
                    dtUsers.Rows.Add(
                                        Convert.ToString(grdUsers.DataKeys[objUserRow.RowIndex]["Id"]),
                                        (objUserRow.FindControl("lblFirstName") as Label).Text,
                                        (objUserRow.FindControl("lblLastName") as Label).Text,
                                        (objUserRow.FindControl("lblDesignation") as Label).Text,
                                        Convert.ToString(grdUsers.DataKeys[objUserRow.RowIndex]["DesignationID"]),
                                        DateTime.Now.AddDays(1).ToShortDateString(),
                                        "10:00 AM"
                                    );
                }
            }

            if (dtUsers.Rows.Count > 0)
            {
                JG_Prospect.Utilits.FullDropDown.FillUserStatus(ddlStatus_Popup, "--Select--", "0");
                LoadUsersByRecruiterDesgination(ddlRecruiter_Popup);

                this.SelectedUsers = dtUsers;
                grdUsers_Popup.DataSource = dtUsers;
                grdUsers_Popup.DataBind();

                upChangeStatusForSelected.Update();

                ScriptManager.RegisterStartupScript
                                (
                                    this,
                                    this.GetType(),
                                    "ShowPopup_divChangeStatusForSelected",
                                    string.Concat("ShowPopupWithTitle('#", divChangeStatusForSelected.ClientID, "');"),
                                    true
                                );
            }
            else
            {
                CommonFunction.ShowAlertFromUpdatePanel(this, "Please select user(s) to change status.");
            }
        }

        //protected void btnshow_Click(object sender, EventArgs e)
        //{
        //    DateTime fromDate = Convert.ToDateTime(txtfrmdate.Text, JG_Prospect.Common.JGConstant.CULTURE);
        //    DateTime toDate = Convert.ToDateTime(txtTodate.Text, JG_Prospect.Common.JGConstant.CULTURE);
        //    if (fromDate < toDate)
        //    {
        //        DataSet ds = InstallUserBLL.Instance.GetHrData(fromDate, toDate, Convert.ToInt16(drpUser.SelectedValue));
        //        if (ds != null)
        //        {
        //            DataTable dtHrData = ds.Tables[0];
        //            DataTable dtgridData = ds.Tables[1];
        //            List<HrData> lstHrData = new List<HrData>();
        //            foreach (DataRow row in dtHrData.Rows)
        //            {
        //                HrData hrdata = new HrData();
        //                hrdata.status = row["status"].ToString();
        //                hrdata.count = row["cnt"].ToString();
        //                lstHrData.Add(hrdata);
        //            }

        //            if (dtHrData.Rows.Count > 0)
        //            {

        //                var rowOfferMade = lstHrData.Where(r => r.status == "OfferMade").FirstOrDefault();
        //                if (rowOfferMade != null)
        //                {
        //                    string count = rowOfferMade.count;
        //                    lbljoboffercount.Text = count;
        //                }
        //                else
        //                {
        //                    lbljoboffercount.Text = "0";
        //                }
        //                var rowActive = lstHrData.Where(r => r.status == "Active").FirstOrDefault();
        //                if (rowActive != null)
        //                {
        //                    string count = rowActive.count;
        //                    lblActiveCount.Text = count;
        //                }
        //                else
        //                {
        //                    lblActiveCount.Text = "0";
        //                }
        //                var rowRejected = lstHrData.Where(r => r.status == "Rejected").FirstOrDefault();
        //                if (rowRejected != null)
        //                {
        //                    string count = rowRejected.count;
        //                    lblRejectedCount.Text = count;
        //                }
        //                else
        //                {
        //                    lblRejectedCount.Text = "0";
        //                }
        //                var rowDeactive = lstHrData.Where(r => r.status == "Deactive").FirstOrDefault();
        //                if (rowDeactive != null)
        //                {
        //                    string count = rowDeactive.count;
        //                    lblDeactivatedCount.Text = count;
        //                }
        //                else
        //                {
        //                    lblDeactivatedCount.Text = "0";
        //                }
        //                var rowInstallProspect = lstHrData.Where(r => r.status == "Install Prospect").FirstOrDefault();
        //                if (rowInstallProspect != null)
        //                {
        //                    string count = rowInstallProspect.count;
        //                    lblInstallProspectCount.Text = count;
        //                }
        //                else
        //                {
        //                    lblInstallProspectCount.Text = "0";
        //                }
        //                var rowPhoneScreened = lstHrData.Where(r => r.status == "PhoneScreened").FirstOrDefault();
        //                if (rowPhoneScreened != null)
        //                {
        //                    string count = rowPhoneScreened.count;
        //                    lblPhoneVideoScreenedCount.Text = count;
        //                }
        //                else
        //                {
        //                    lblPhoneVideoScreenedCount.Text = "0";
        //                }
        //                var rowInterviewDate = lstHrData.Where(r => r.status == "InterviewDate").FirstOrDefault();
        //                if (rowInterviewDate != null)
        //                {
        //                    string count = rowInterviewDate.count;
        //                    lblInterviewDateCount.Text = count;
        //                }
        //                else
        //                {
        //                    lblInterviewDateCount.Text = "0";
        //                }
        //                var rowApplicant = lstHrData.Where(r => r.status == "Applicant").FirstOrDefault();
        //                string Applicantcount = "0";
        //                if (rowApplicant != null)
        //                {
        //                    Applicantcount = rowApplicant.count;

        //                }
        //                else
        //                {
        //                    Applicantcount = "0";

        //                }
        //                // Ratio Calculation
        //                lblAppInterviewRatio.Text = Convert.ToString(Convert.ToDouble(lblInterviewDateCount.Text) / Convert.ToDouble(Applicantcount));
        //                //lblAppHireRatio.Text = Convert.ToString(Convert.ToDouble(lblActiveCount.Text) / Convert.ToDouble(Applicantcount));
        //                //lblJobOfferHireRatio.Text = Convert.ToString(Convert.ToDouble(lblActive.Text) / Convert.ToDouble(lblInterviewDateCount.Text));
        //            }
        //            else
        //            {
        //                lbljoboffercount.Text = "0";
        //                lblActiveCount.Text = "0";
        //                lblRejectedCount.Text = "0";
        //                lblDeactivatedCount.Text = "0";
        //                lblInstallProspectCount.Text = "0";
        //                lblPhoneVideoScreenedCount.Text = "0";
        //                lblInterviewDateCount.Text = "0";
        //                lblAppInterviewRatio.Text = "0";
        //                //lblAppHireRatio.Text = "0";
        //            }
        //            if (dtgridData.Rows.Count > 0)
        //            {
        //                Session["UserGridData"] = dtgridData;
        //                BindUsers(dtgridData);
        //                //grdUsers.DataSource = dtgridData;
        //                //grdUsers.DataBind();
        //            }
        //            else
        //            {
        //                Session["UserGridData"] = null;
        //                grdUsers.DataSource = null;
        //                grdUsers.DataBind();
        //            }
        //        }
        //    }
        //    else
        //    {
        //        ScriptManager.RegisterStartupScript(this, this.GetType(), "AlertBox", "alert('ToDate must be greater than FromDate');", true);
        //    }
        //}

        #endregion

        #region '--Methods--'

        //private void BindUsers(DataTable dt)
        //{
        //    try
        //    {
        //        string usertype = Session["usertype"].ToString().ToLower();

        //        if (dt != null && dt.Columns["OrderStatus"] == null)
        //        {
        //            dt.Columns.Add("OrderStatus");
        //            int st = 0;

        //            if (usertype == "jg account" || usertype == "sales manager" || usertype == "office manager" || usertype == "recruiter")
        //            {
        //                foreach (DataRow dr in dt.Rows)
        //                {
        //                    st = (int)((OrderStatus1)Enum.Parse(typeof(OrderStatus1), dr["Status"].ToString().Replace(" ", "")));
        //                    dr["OrderStatus"] = st.ToString();
        //                }
        //            }
        //            else if (usertype == "admin" || usertype == "jr. sales" || usertype == "project manager")
        //            {
        //                foreach (DataRow dr in dt.Rows)
        //                {
        //                    st = (int)((OrderStatus2)Enum.Parse(typeof(OrderStatus2), dr["Status"].ToString().Replace(" ", "")));
        //                    dr["OrderStatus"] = st.ToString();
        //                }
        //            }
        //        }

        //        if (dt != null)
        //        {
        //            DataView dv = dt.DefaultView;
        //            dv.Sort = "OrderStatus asc";
        //            grdUsers.DataSource = dv;
        //            grdUsers.DataBind();
        //        }
        //        else
        //        {
        //            grdUsers.DataSource = null;
        //            grdUsers.DataBind();
        //        }

        //        //DataView dv = dt.DefaultView;
        //        //dv.Sort = "OrderStatus asc";
        //        //grdUsers.DataSource = dv;
        //        //grdUsers.DataBind();
        //    }
        //    catch (Exception ex)
        //    {
        //        UtilityBAL.AddException("Edituser", Session["loginid"] == null ? "" : Session["loginid"].ToString(), ex.Message, ex.StackTrace);
        //    }
        //}

        //private void binddata()
        //{
        //    DataSet dsUsers_Export = InstallUserBLL.Instance.GetAllSalesUserToExport();
        //    DataSet dsUsers_Grid = InstallUserBLL.Instance.GetAllEditSalesUser();

        //    BindPieChart(dsUsers_Grid.Tables[0]);

        //    //DS.Tables[0].Columns[4].DataType = typeof(Int32);
        //    Session["GridDataExport"] = dsUsers_Export.Tables[0];
        //    Session["UserGridData"] = dsUsers_Grid.Tables[0];

        //    BindUsers(dsUsers_Grid.Tables[0]);
        //    //grdUsers.DataSource = DS.Tables[0];
        //    //grdUsers.DataBind();

        //    //List<string> lstDesignation= (from ptrade in DS.Tables[0].AsEnumerable()
        //    // where !string.IsNullOrEmpty(ptrade.Field<string>("Designation"))
        //    // select Convert.ToString(ptrade["Designation"])).Distinct().ToList();

        //    //lstDesignation.Sort((x, y) => string.Compare(x, y));
        //    //ddlDesignation.DataSource = lstDesignation;

        //    //ddlDesignation.DataBind();            
        //    BindDesignations();
        //}

        private void BindDesignations()
        {
            DataSet dsDesignation = new DataSet();
            dsDesignation = DesignationBLL.Instance.GetAllDesignationsForHumanResource();
            if (dsDesignation.Tables.Count > 0)
            {
                ddlDesignation.DataSource = dsDesignation.Tables[0];
                ddlDesignation.DataTextField = "DesignationName";
                ddlDesignation.DataValueField = "ID";
                ddlDesignation.DataBind();

                ddlDesignationForTask.DataSource = dsDesignation.Tables[0];
                ddlDesignationForTask.DataTextField = "DesignationName";
                ddlDesignationForTask.DataValueField = "ID";
                ddlDesignationForTask.DataBind();
            }
            ddlDesignation.Items.Insert(0, new ListItem("--All--", "0"));
            ddlDesignationForTask.Items.Insert(0, new ListItem("--All--", "0"));
        }

        private void FillCustomer()
        {

            fillFilterUserDDL();

            ddlUserStatus = JG_Prospect.Utilits.FullDropDown.FillUserStatus(ddlUserStatus, "--All--", "0");

            ddlUserStatus.Items.FindByValue(Convert.ToString((byte)JGConstant.InstallUserStatus.Active)).Enabled = false;
            ddlUserStatus.Items.FindByValue(Convert.ToString((byte)JGConstant.InstallUserStatus.OfferMade)).Enabled = false;


            DataSet dsSource = new DataSet();
            dsSource = InstallUserBLL.Instance.GetSource();
            DataRow drSource = dsSource.Tables[0].NewRow();
            drSource["Id"] = "0";
            drSource["Source"] = "--All--";
            dsSource.Tables[0].Rows.InsertAt(drSource, 0);
            if (dsSource.Tables[0].Rows.Count > 0)
            {
                ddlSource.DataSource = dsSource.Tables[0];
                ddlSource.DataValueField = "Id";
                ddlSource.DataTextField = "Source";
                ddlSource.DataBind();
            }
        }

        private void fillFilterUserDDL()
        {
            //DataSet dds = new DataSet();
            //dds = new_customerBLL.Instance.GeUsersForDropDown();
            //DataRow dr = dds.Tables[0].NewRow();

            //dr["Id"] = "0";
            //dr["Username"] = "--All--";
            //dds.Tables[0].Rows.InsertAt(dr, 0);
            //if (dds.Tables[0].Rows.Count > 0)
            //{
            //    drpUser.DataSource = dds.Tables[0];
            //    drpUser.DataValueField = "Id";
            //    drpUser.DataTextField = "Username";
            //    drpUser.DataBind();
            //}

            DataSet dsInstalledUser = InstallUserBLL.Instance.GetUsersNDesignationForSalesFilter();
            drpUser.DataSource = dsInstalledUser.Tables[0];
            drpUser.DataValueField = "Id";
            drpUser.DataTextField = "FirstName";
            drpUser.DataBind();
            drpUser.Items.Insert(0, new ListItem("--All--", "0"));
            DataTable dtInstalledUsers = dsInstalledUser.Tables[0];
            Session["HighlightUsersForTypes"] = dtInstalledUsers;
            HighlightUsersForTypes(dtInstalledUsers, drpUser);
        }

        private void HighlightUsersForTypes(DataTable dtUsers, DropDownList ddlUsers)
        {
            if (dtUsers != null && dtUsers.Rows.Count > 0)
            {
                var rows = dtUsers.AsEnumerable();

                //get all users comma seperated ids with Active status
                String ActivatedUsers = String.Join(",", (from r in rows where (r.Field<string>("GroupNumber") == "Group1") select r.Field<Int32>("Id").ToString()));

                // for each userid find it into user dropdown list and apply red color to it.
                foreach (String user in ActivatedUsers.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    ListItem item;

                    if (ddlUsers != null)
                    {
                        item = ddlUsers.Items.FindByValue(user);
                        item.Attributes.Add("style", "color:red;");
                    }
                }

                //get all users comma seperated ids with interviewdate status
                String InterviewdateUsers = String.Join(",", (from r in rows where (r.Field<string>("GroupNumber") == "Group2") select r.Field<Int32>("Id").ToString()));

                // for each userid find it into user dropdown list and apply red color to it.
                foreach (String user in InterviewdateUsers.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    ListItem item;

                    if (ddlUsers != null)
                    {
                        item = ddlUsers.Items.FindByValue(user);
                        item.Attributes.Add("style", "color:blue;");
                    }
                }

                //get all users comma seperated ids with deactive status
                String DeactivatedUsers = String.Join(",", (from r in rows where (r.Field<string>("GroupNumber") == "Group3") select r.Field<Int32>("Id").ToString()));

                // for each userid find it into user dropdown list and apply red color to it.
                foreach (String user in DeactivatedUsers.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    ListItem item;

                    if (ddlUsers != null)
                    {
                        item = ddlUsers.Items.FindByValue(user);
                        item.Attributes.Add("style", "color:grey;");
                    }
                }

                ////get all users comma seperated ids with interviewdate status
                //String InstallProspectUsers = String.Join(",", (from r in rows where (r.Field<string>("GroupNumber") == "Group3") select r.Field<Int32>("Id").ToString()));

                //// for each userid find it into user dropdown list and apply red color to it.
                //foreach (String user in InstallProspectUsers.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                //{
                //    ListItem item;

                //    if (ddlUsers != null)
                //    {
                //        item = ddlUsers.Items.FindByValue(user);
                //        item.Attributes.Add("style", "color:green;");
                //    }
                //}

                ////get all users comma seperated ids with interviewdate status
                //String OfferMadeInterviewDateUsers = String.Join(",", (from r in rows where (r.Field<string>("GroupNumber") == "Group4") select r.Field<Int32>("Id").ToString()));

                //// for each userid find it into user dropdown list and apply red color to it.
                //foreach (String user in OfferMadeInterviewDateUsers.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                //{
                //    ListItem item;

                //    if (ddlUsers != null)
                //    {
                //        item = ddlUsers.Items.FindByValue(user);
                //        item.Attributes.Add("style", "color:red;");
                //    }
                //}
            }
        }

        protected void DownloadFile(object sender, EventArgs e)
        {
            string filePath = (sender as LinkButton).CommandArgument;
            Response.ContentType = ContentType;
            Response.AppendHeader("Content-Disposition", "attachment; filename=" + Path.GetFileName(filePath));
            Response.WriteFile(filePath);
            Response.End();
        }

        private string GetId(string Desig, string UserStatus)
        {
            string LastInt = "";
            DataTable dtId;
            string SalesId = string.Empty;
            dtId = InstallUserBLL.Instance.GetMaxSalesId(Desig);
            if (dtId.Rows.Count > 0)
            {
                SalesId = Convert.ToString(dtId.Rows[0][0]);
            }
            if ((Desig == "Admin" || Desig == "Office Manager" || Desig == "Recruiter") && (UserStatus != "Deactive"))
            {
                if (SalesId != "")
                {
                    LastInt = SalesId.Substring(SalesId.Length - 4);
                    SalesId = "ADM000" + Convert.ToString(Convert.ToInt32(LastInt) + 1);
                }
                else
                {
                    SalesId = "ADM0001";
                }
            }
            else if ((Desig == "Admin" || Desig == "Office Manager" || Desig == "Recruiter") && (UserStatus == "Deactive") && (SalesId != ""))
            {
                SalesId = SalesId + "-X";
            }
            else if ((Desig == "Jr. Sales" || Desig == "Jr Project Manager" || Desig == "Sales Manager" || Desig == "Sr. Sales") && (UserStatus != "Deactive"))
            {
                if (SalesId != "")
                {
                    LastInt = SalesId.Substring(SalesId.Length - 4);
                    SalesId = "SLE000" + Convert.ToString(Convert.ToInt32(LastInt) + 1);
                }
                else
                {
                    SalesId = "SLE0001";
                }
            }
            else if ((Desig == "Jr. Sales" || Desig == "Jr Project Manager" || Desig == "Sales Manager" || Desig == "Sr. Sales") && (UserStatus == "Deactive") && (SalesId != ""))
            {
                SalesId = SalesId + "-X";
            }

            return SalesId;
        }

        private DataTable FillValueFromFiles(string FilePath, string Extension)
        {
            DataTable dtExcel = new DataTable();
            ExcelPackage package = new ExcelPackage();

            switch (Extension)
            {
                case ".xls":
                case ".xlsx":

                    package = new ExcelPackage(BulkProspectUploader.FileContent);
                    dtExcel = ReadExcelFile(package);

                    break;

                case ".csv":
                    dtExcel = ReadCsvFile(FilePath);
                    break;
            }

            return dtExcel;
        }

        public DataTable GetImportIntsallUsersDataTable()
        {
            DataTable objDataTable = new DataTable();

            objDataTable.Columns.Add("Designation", typeof(string));
            objDataTable.Columns.Add("Status", typeof(string));
            objDataTable.Columns.Add("DateSourced", typeof(string));
            objDataTable.Columns.Add("Source", typeof(string));
            objDataTable.Columns.Add("FirstName", typeof(string));
            objDataTable.Columns.Add("LastName", typeof(string));
            objDataTable.Columns.Add("Email", typeof(string));
            objDataTable.Columns.Add("Phone1", typeof(string));
            objDataTable.Columns.Add("Phone1Type", typeof(string));
            objDataTable.Columns.Add("Phone2", typeof(string));
            objDataTable.Columns.Add("Phone2Type", typeof(string));
            objDataTable.Columns.Add("Address", typeof(string));
            objDataTable.Columns.Add("Zip", typeof(string));
            objDataTable.Columns.Add("State", typeof(string));
            objDataTable.Columns.Add("City", typeof(string));
            objDataTable.Columns.Add("Suit_Apt_Room", typeof(string));
            objDataTable.Columns.Add("Notes", typeof(string));

            return objDataTable;
        }

        public DataTable ReadExcelFile(ExcelPackage package)
        {
            DataTable dtUser = GetImportIntsallUsersDataTable();
            try
            {
                ExcelWorksheet workSheet = package.Workbook.Worksheets.First();

                // excluding 1st row containing heading.
                for (var rowNumber = 2; rowNumber <= workSheet.Dimension.End.Row; rowNumber++)
                {
                    var row = workSheet.Cells[rowNumber, 1, rowNumber, dtUser.Columns.Count];
                    var drUser = dtUser.NewRow();
                    int intFilledColumnCount = 0;
                    foreach (var cell in row)
                    {
                        drUser[cell.Start.Column - 1] = cell.Text;
                        if (!string.IsNullOrEmpty(cell.Text))
                        {
                            intFilledColumnCount++;
                        }
                    }

                    // stop looping, if end of file.
                    // we have default value 'Applicant' in designation column, so, it will never be blank. So, we used -1 in below condition.
                    if (intFilledColumnCount <= 1)
                    {
                        break;
                    }

                    dtUser.Rows.Add(drUser);
                }

            }
            catch (Exception ex)
            {
                UtilityBAL.AddException("EditUser-ReadExcelFile", Session["loginid"] == null ? "" : Session["loginid"].ToString(), ex.Message, ex.StackTrace);
                return null;

            }

            return dtUser;
        }

        public DataTable ReadCsvFile(string FileSaveWithPath)
        {
            /// Ref Site :http://www.c-sharpcorner.com/blogs/read-csv-file-into-data-table1

            DataTable dtUser = GetImportIntsallUsersDataTable();
            string Fulltext;

            //string FileSaveWithPath = Server.MapPath("\\Files\\Import" + System.DateTime.Now.ToString("ddMMyyyy_hhmmss") + ".csv");
            //FileUpload1.SaveAs(FileSaveWithPath);
            using (StreamReader sr = new StreamReader(FileSaveWithPath))
            {
                while (!sr.EndOfStream)
                {
                    Fulltext = sr.ReadToEnd().ToString(); //read full file text  
                    string[] rows = Fulltext.Split('\n'); //split full file text into rows  
                    for (int i = 0; i < rows.Count() - 1; i++)
                    {
                        string[] rowValues = rows[i].Split(','); //split each row with comma to get individual values  
                        {
                            // excluding 1st row containing heading.
                            if (i > 0)
                            {
                                DataRow drUser = dtUser.NewRow();
                                for (int k = 0; k < rowValues.Count(); k++)
                                {
                                    if ((k == rowValues.Count() - 1) && (rowValues[k].IndexOf("\r") > 0)) //CSV last col value many have "/r"
                                        // Remove /r from value                                         
                                        drUser[k] = rowValues[k].ToString().Replace("\r", "");
                                    else
                                        drUser[k] = rowValues[k].ToString();

                                }
                                dtUser.Rows.Add(drUser); //add other rows  
                            }
                        }
                    }
                }
            }

            return dtUser;
        }

        private DataTable FillValueFromFiles_old(string FilePath, string Extension, string FileName)
        {
            string conStr = "";
            switch (Extension)
            {
                case ".xls": //Excel 97-03
                    conStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FilePath + ";Extended Properties='Excel 8.0;HDR=No;IMEX=1'";
                    //conStr = ConfigurationManager.ConnectionStrings["Excel03ConString"]
                    //         .ConnectionString;
                    break;

                case ".xlsx": //Excel 07
                    conStr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FilePath + ";Extended Properties=Excel 12.0;";
                    //conStr = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + FilePath + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'";
                    //conStr = ConfigurationManager.ConnectionStrings["Excel07ConString"]
                    //          .ConnectionString;
                    break;

                case ".csv":
                    FilePath = FilePath.Substring(0, FilePath.LastIndexOf(FileName)); // Removing file name. 
                    conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + FilePath + ";Extended Properties=\"Text;HDR=Yes;FORMAT=Delimited\"";

                    break;
            }
            conStr = String.Format(conStr, FilePath);
            OleDbConnection connExcel = new OleDbConnection(conStr);
            OleDbCommand cmdExcel = new OleDbCommand();
            OleDbDataAdapter oda = new OleDbDataAdapter();
            DataTable dtExcel = new DataTable();
            cmdExcel.Connection = connExcel;
            string IdGenerated = "";
            //Get the name of First Sheet
            connExcel.Open();
            DataTable dtExcelSchema;
            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
            connExcel.Close();

            //Read Data from First Sheet
            connExcel.Open();
            cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";
            oda.SelectCommand = cmdExcel;
            oda.Fill(dtExcel);
            connExcel.Close();

            return dtExcel;
        }

        private bool ValidateUploadedData(DataTable dtUser)
        {
            DataTable dtUserError = GetImportIntsallUsersDataTable();
            string strMessage = string.Empty;

            bool IsValid = true;
            if (dtUser == null)
            {
                strMessage = "Kindly validate uploaded Data / File.";
                IsValid = false;
            }
            else if (dtUser.Rows.Count <= 0)
            {
                strMessage = "No data found to uploade , Kindly check the uploaded file";
                IsValid = false;
            }
            else
            {
                // Validate data -- Mobile no , email is entered or so...
                for (int i = 0; i < dtUser.Rows.Count; i++)
                {
                    bool blRowValid = true;

                    if (string.IsNullOrEmpty(Convert.ToString(dtUser.Rows[i]["Designation"])))
                    {
                        strMessage = "Kindly enter Designition value for all the records";

                        blRowValid =
                        IsValid = false;
                    }

                    if (string.IsNullOrEmpty(Convert.ToString(dtUser.Rows[i]["Status"])))
                    {
                        strMessage = "Kindly enter Status value for all the records";

                        blRowValid =
                        IsValid = false;
                    }
                    if (string.IsNullOrEmpty(Convert.ToString(dtUser.Rows[i]["Source"])))
                    {
                        strMessage = "Kindly enter Source value for all the records";

                        blRowValid =
                        IsValid = false;
                    }
                    if (string.IsNullOrEmpty(Convert.ToString(dtUser.Rows[i]["FirstName"])))
                    {
                        strMessage = "Kindly enter First Name value for all the records";

                        blRowValid =
                        IsValid = false;
                    }
                    if (string.IsNullOrEmpty(Convert.ToString(dtUser.Rows[i]["LastName"])))
                    {
                        strMessage = "Kindly enter Last Name value for all the records";

                        blRowValid =
                        IsValid = false;
                    }
                    if (string.IsNullOrEmpty(Convert.ToString(dtUser.Rows[i]["Email"])))
                    {
                        strMessage = "Kindly enter Email value for all the records";

                        blRowValid =
                        IsValid = false;
                    }
                    if (string.IsNullOrEmpty(Convert.ToString(dtUser.Rows[i]["Phone1"])))
                    {
                        strMessage = "Kindly enter Primary Contact Phone value for all the records";

                        blRowValid =
                        IsValid = false;
                    }
                    else if (string.IsNullOrEmpty(Convert.ToString(dtUser.Rows[i]["Phone1Type"])))
                    {
                        strMessage = "Kindly enter Phone Type value for all the records";

                        blRowValid =
                        IsValid = false;
                    }
                    else if (string.IsNullOrEmpty(Convert.ToString(dtUser.Rows[i]["Zip"])))
                    {
                        strMessage = "Kindly enter Zip value for all the records";

                        blRowValid =
                        IsValid = false;
                    }

                    if (!blRowValid)
                    {
                        dtUserError.Rows.Add(dtUser.Rows[i].ItemArray);
                    }
                }
            }

            if (!IsValid)
            {
                grdBulkUploadUserErrors.DataSource = dtUserError;
                grdBulkUploadUserErrors.DataBind();

                ScriptManager.RegisterStartupScript
                    (
                        this,
                        this.GetType(),
                        "ShowPopup_divBulkUploadUserErrors",
                        string.Concat(
                                        "ShowPopupWithTitle('#", divBulkUploadUserErrors.ClientID, "','Information');"
                                     ),
                        true
                    );
            }

            return IsValid;
        }

        public XmlDocument GetIntsallUsersXmlDoc(DataTable dtExcel)
        {
            string helper = "";
            bool IsValid = true;

            List<Designation> lstDesignation = DesignationBLL.Instance.GetAllDesignation();
            DataTable dtSource = InstallUserBLL.Instance.GetSource().Tables[0];

            List<user1> lstIntsallUsers = new List<user1>();
            user1 objIntsallUser = null;

            for (int i = 0; i < dtExcel.Rows.Count; i++)
            {
                try
                {
                    objIntsallUser = new user1();

                    #region Old - Field Description

                    //0 ID #: ---	1 *Designitions:--	2 status:	-- 3 Date Sourced: 	-- 4 *First Name*  	-- 5 *Last Name	-- 6 * Source	-- 7 *Primary contact phone #:(3-3-4)
                    //8 *phone type:(drop down: Cell Phone #, House Phone #, Work Phone #, Alt #)	-- 9 secondary contact phone #(3-3-4)	-- 10 phone type:(drop down: Cell Phone #, House Phone #, Work Phone #, Alt #)
                    //11 *Company Name	-- 12 *Primary Trade 	-- 13 SecondaryTrade* (list as many secondary… 1 primary)	
                    //14 *Home Address  	-- 15 Zip  	-- 16 State  17 City 	 -- 18 Suite/Apt/Room(If applicable)   
                    //19 *Secondary Address	 -- 20 Zip  -- 21 State -- 22 City 	-- 23 Suite/Apt/Room(If applicable)
                    //24 Are you currently employed? 	-- 25 Reason for leaving your current employer/position  -- 26 Have you ever applied or worked here before? 
                    //27 How many full time positions have you had in the past 5 years?	 -- 28 Can you tell me a little about any sales or construction industry experience you have?
                    //29 No FELONY or DUI charges?  -- 30 Will you be able to pass a drug test and background check?  -- 31  What are your salary requirements for this position?
                    //32 If selected for position, when will you be available to start?

                    #endregion

                    #region Values From Excel

                    Designation objDesignation = lstDesignation.FirstOrDefault(d => d.DesignationName
                                                                                     .ToUpper()
                                                                                    .Trim() == dtExcel.Rows[i]["Designation"]
                                                                                                      .ToString()
                                                                                                      .ToUpper()
                                                                                                      .Trim()
                                                                              );
                    if (objDesignation != null)
                    {
                        objIntsallUser.DesignationId = objDesignation.ID;
                        objIntsallUser.Designation = objDesignation.ID.ToString();
                    }
                    else
                    {
                        objIntsallUser.DesignationId = 0;
                        objIntsallUser.Designation = string.Empty;
                    }

                    JGConstant.InstallUserStatus objInstallUserStatus;
                    if (Enum.TryParse<JGConstant.InstallUserStatus>(dtExcel.Rows[i]["Status"].ToString().Trim(), true, out objInstallUserStatus))
                    {
                        objIntsallUser.status = Convert.ToByte(objInstallUserStatus).ToString();
                    }
                    else
                    {
                        objIntsallUser.status = string.Empty;
                    }

                    //objuser.DateSourced = DateTime.Now.ToString();
                    //objuser.Source = Convert.ToString(Session["Username"]);
                    objIntsallUser.DateSourced = dtExcel.Rows[i]["DateSourced"].ToString().Trim();
                    objIntsallUser.Source = dtExcel.Rows[i]["Source"].ToString().Trim();

                    objIntsallUser.SourceId = (from DataRow drSource in dtSource.Rows
                                               where (string)drSource["Source"] == objIntsallUser.Source
                                               select (int)drSource["Id"]).FirstOrDefault();

                    objIntsallUser.firstname = dtExcel.Rows[i]["FirstName"].ToString().Trim();
                    objIntsallUser.lastname = dtExcel.Rows[i]["LastName"].ToString().Trim();

                    objIntsallUser.Email = dtExcel.Rows[i]["Email"].ToString().Trim();
                    objIntsallUser.phone = dtExcel.Rows[i]["Phone1"].ToString().Trim();
                    objIntsallUser.phonetype = dtExcel.Rows[i]["Phone1Type"].ToString().Trim();
                    objIntsallUser.Phone2 = dtExcel.Rows[i]["Phone2"].ToString().Trim();
                    objIntsallUser.Phone2Type = dtExcel.Rows[i]["Phone2Type"].ToString().Trim();

                    objIntsallUser.address = dtExcel.Rows[i]["Address"].ToString().Trim();
                    objIntsallUser.zip = dtExcel.Rows[i]["Zip"].ToString().Trim();
                    objIntsallUser.state = dtExcel.Rows[i]["State"].ToString().Trim();
                    objIntsallUser.city = dtExcel.Rows[i]["City"].ToString().Trim();
                    objIntsallUser.SuiteAptRoom = dtExcel.Rows[i]["Suit_Apt_Room"].ToString().Trim();

                    objIntsallUser.Notes = dtExcel.Rows[i]["Notes"].ToString().Trim();

                    #endregion

                    #region Other Values

                    objIntsallUser.Email2 = string.Empty;

                    objIntsallUser.CompanyName = string.Empty;

                    objIntsallUser.SourceUser = Convert.ToString(Session["userid"]);

                    if (helper == "yes" || helper == "true")
                    {
                        objIntsallUser.CurrentEmployement = true;
                        objIntsallUser.PrevApply = true;
                        objIntsallUser.FELONY = true;
                        objIntsallUser.DrugTest = true;
                    }
                    else if (helper == "no" || helper == "false")
                    {
                        objIntsallUser.CurrentEmployement = false;
                        objIntsallUser.PrevApply = false;
                        objIntsallUser.FELONY = false;
                        objIntsallUser.DrugTest = false;
                    }

                    objIntsallUser.UserType = "SalesUser";

                    #endregion

                    if (
                        string.IsNullOrEmpty(objIntsallUser.Designation) ||
                        string.IsNullOrEmpty(objIntsallUser.status) ||
                        string.IsNullOrEmpty(objIntsallUser.Source) ||
                        string.IsNullOrEmpty(objIntsallUser.firstname) ||
                        string.IsNullOrEmpty(objIntsallUser.lastname) ||
                        string.IsNullOrEmpty(objIntsallUser.Email) ||
                        string.IsNullOrEmpty(objIntsallUser.phone) ||
                        string.IsNullOrEmpty(objIntsallUser.phonetype) ||
                        string.IsNullOrEmpty(objIntsallUser.zip)
                       )
                    {
                        IsValid = false;
                        break;
                    }
                    else
                    {
                        lstIntsallUsers.Add(objIntsallUser);
                    }

                    #region Commented - CheckDuplicate

                    //DataSet dsCheckDuplicate = InstallUserBLL.Instance.CheckInstallUser(dtExcel.Rows[i][5].ToString().Trim(), dtExcel.Rows[i][3].ToString().Trim());

                    //if (dsCheckDuplicate.Tables[0].Rows.Count == 0)
                    //{
                    //    IdGenerated = GetId(dtExcel.Rows[i][9].ToString().Trim(), dtExcel.Rows[i][10].ToString().Trim());
                    //    objuser.InstallId = IdGenerated;
                    //    DataSet ds = InstallUserBLL.Instance.CheckSource(Convert.ToString(Session["Username"]));
                    //    if (ds.Tables[0].Rows.Count > 0)
                    //    {
                    //        //do nothing
                    //    }
                    //    else
                    //    {
                    //        DataSet dsadd = InstallUserBLL.Instance.AddSource(Convert.ToString(Session["Username"]));
                    //    }
                    //    //objuser.DateSourced = Convert.ToString(dtExcel.Rows[i][9].ToString());
                    //    objuser.Notes = dtExcel.Rows[i][8].ToString().Trim();
                    //    bool result = InstallUserBLL.Instance.AddUser(objuser);
                    //    count += Convert.ToInt32(result);
                    //}

                    #endregion
                }
                catch (Exception ex)
                {
                    UtilityBAL.AddException("EditUser-CreateUserObjectXml", Session["loginid"] == null ? "" : Session["loginid"].ToString(), ex.Message, ex.StackTrace);
                    continue;
                }
            }

            #region Commented - CheckDuplicate

            ////check duplicacy of data in sheet itself
            //var duplicate = from c in list.AsEnumerable()
            //                group c by c.Email into grp
            //                where grp.Count() > 1
            //                select grp.Key;
            //if (duplicate.ToList().Count > 0)
            //{
            //    IsValid = false;
            //    ScriptManager.RegisterStartupScript(this, this.GetType(), "AlertBox", "alert('Upload file contains data error or matching data exists, please check and upload again');", true);
            //    return;
            //} 

            #endregion

            XmlDocument xmlDoc = null;

            if (IsValid)
            {
                xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(Serialize(lstIntsallUsers));

                if (xmlDoc.FirstChild.NodeType == XmlNodeType.XmlDeclaration)
                {
                    xmlDoc.RemoveChild(xmlDoc.FirstChild);
                }
            }
            else
            {
                CommonFunction.ShowAlertFromPage(this, "Upload file contains data error or matching data exists, please check and upload again.");
            }

            return xmlDoc;
        }

        public void CreateDuplicateUserObjectXml(DataTable dt, out XmlDocument xmlDoc)
        {
            List<user1> list = new List<user1>();
            string helper = "";
            user1 objuser = null;
            xmlDoc = new XmlDocument();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                try
                {
                    objuser = new user1();

                    #region BindUserObject

                    //helper = dt.Rows[i]["Id"].ToString().Trim();
                    //objuser.Id = helper == "" ? 0 : Convert.ToInt32(helper);

                    objuser.Email = dt.Rows[i]["Email"].ToString().Trim();
                    objuser.Email2 = dt.Rows[i]["Email2"].ToString().Trim();
                    objuser.Designation = dt.Rows[i]["Designation"].ToString().Trim();
                    objuser.status = dt.Rows[i]["status"].ToString().Trim();
                    objuser.DateSourced = dt.Rows[i]["DateSourced"].ToString().Trim();
                    objuser.firstname = dt.Rows[i]["firstname"].ToString().Trim();
                    objuser.lastname = dt.Rows[i]["lastname"].ToString().Trim();
                    objuser.SourceUser = Convert.ToString(Session["userid"]);
                    objuser.Source = Convert.ToString(Session["Username"]);

                    objuser.phone = dt.Rows[i]["phone"].ToString().Trim();
                    //objuser.phonetype = dt.Rows[i]["phonetype"].ToString().Trim();
                    objuser.Phone2 = dt.Rows[i]["Phone2"].ToString().Trim();
                    //objuser.Phone2Type = dt.Rows[i]["Phone2Type"].ToString().Trim();

                    objuser.CompanyName = dt.Rows[i]["CompanyName"].ToString().Trim();

                    helper = dt.Rows[i]["PrimeryTradeId"].ToString().Trim();
                    objuser.PrimeryTradeId = helper == "" ? 0 : Convert.ToInt32(helper);

                    helper = dt.Rows[i]["SecondoryTradeId"].ToString().Trim();
                    objuser.SecondoryTradeId = helper == "" ? 0 : Convert.ToInt32(helper);

                    //objuser.address = dt.Rows[i]["address"].ToString().Trim();
                    //objuser.zip = dt.Rows[i]["zip"].ToString().Trim();
                    //objuser.state = dt.Rows[i]["state"].ToString().Trim();
                    //objuser.city = dt.Rows[i]["city"].ToString().Trim();
                    //objuser.SuiteAptRoom = dt.Rows[i]["SuiteAptRoom"].ToString().Trim();

                    //objuser.Address2 = dt.Rows[i]["Address2"].ToString().Trim();
                    //objuser.Zip2 = dt.Rows[i]["Zip2"].ToString().Trim();
                    //objuser.State2 = dt.Rows[i]["State2"].ToString().Trim();
                    //objuser.City2 = dt.Rows[i]["City2"].ToString().Trim();
                    //objuser.SuiteAptRoom2 = dt.Rows[i]["SuiteAptRoom2"].ToString().Trim();

                    //helper = dt.Rows[i]["CurrentEmployement"].ToString().Trim().ToLower();

                    //if (helper == "yes" || helper == "true")
                    //    objuser.CurrentEmployement = true;
                    //else if (helper == "no" || helper == "false")
                    //    objuser.CurrentEmployement = false;

                    //objuser.LeavingReason = dt.Rows[i]["LeavingReason"].ToString().Trim();

                    //helper = dt.Rows[i]["PrevApply"].ToString().Trim().ToLower();

                    //if (helper == "yes" || helper == "true")
                    //    objuser.PrevApply = true;
                    //else if (helper == "no" || helper == "false")
                    //    objuser.PrevApply = false;

                    //helper = dt.Rows[i]["FullTimePosition"].ToString().Trim();
                    //objuser.FullTimePosition = helper == "" ? 0 : Convert.ToInt32(helper);
                    //objuser.SalesExperience = dt.Rows[i]["SalesExperience"].ToString().Trim();

                    //helper = dt.Rows[i]["FELONY"].ToString().Trim().ToLower();

                    //if (helper == "yes" || helper == "true")
                    //    objuser.FELONY = true;
                    //else if (helper == "no" || helper == "false")
                    //    objuser.FELONY = false;

                    //helper = dt.Rows[i]["DrugTest"].ToString().Trim().ToLower();

                    //if (helper == "yes" || helper == "true")
                    //    objuser.DrugTest = true;
                    //else if (helper == "no" || helper == "false")
                    //    objuser.DrugTest = false;

                    //objuser.SalaryReq = dt.Rows[i]["SalaryReq"].ToString().Trim();
                    //objuser.Avialability = dt.Rows[i]["Avialability"].ToString().Trim();

                    #endregion

                    list.Add(objuser);
                }
                catch (Exception ex)
                {
                    continue;
                }
            }

            xmlDoc.LoadXml(Serialize(list));

            if (xmlDoc.FirstChild.NodeType == XmlNodeType.XmlDeclaration)
                xmlDoc.RemoveChild(xmlDoc.FirstChild);
        }

        private void ImportIntsallUsers(DataTable dtExcel)
        {
            XmlDocument xmlDoc = GetIntsallUsersXmlDoc(dtExcel);

            if (xmlDoc != null)
            {
                DataSet dsImportResult = null;

                if (xmlDoc.OuterXml != "")
                {
                    dsImportResult = InstallUserBLL.Instance.BulkIntsallUser(xmlDoc.InnerXml);
                }

                pnlAddNewUser.Visible = false;
                pnlDuplicate.Visible = false;

                #region '-- Process Import Result --'

                // duplicate users
                if (dsImportResult != null && dsImportResult.Tables.Count > 0 && dsImportResult.Tables[0].Rows.Count > 0)
                {

                    int RowCount = (from DataRow ReturnDr in dsImportResult.Tables[0].Rows
                                    where (string)ReturnDr["ActionTaken"] != "I"
                                    select (string)ReturnDr["Email"]).Count();

                    if (RowCount > 0) // if found any row not Inserted than
                    {
                        DataTable DuplicateRecords = (from DataRow ReturnDr in dsImportResult.Tables[0].Rows
                                                      where (string)ReturnDr["ActionTaken"] != "I"
                                                      select ReturnDr).CopyToDataTable();

                        Session["DuplicateUsers"] = DuplicateRecords;

                        listDuplicateUsers.DataSource = DuplicateRecords;
                        listDuplicateUsers.DataBind();

                        lblDuplicateCount.Text = "<h1>Duplicate Records : (" + RowCount.ToString() + ")</h1>";

                        pnlDuplicate.Visible = true;
                    }

                    RowCount = (from DataRow ReturnDr in dsImportResult.Tables[0].Rows
                                where (string)ReturnDr["ActionTaken"] == "I"
                                select (string)ReturnDr["Email"]).Count();

                    if (RowCount > 0) // if row Inserted / Added
                    {
                        DataTable InsertedRecords = (from DataRow ReturnDr in dsImportResult.Tables[0].Rows
                                                     where (string)ReturnDr["ActionTaken"] == "I"
                                                     select ReturnDr).CopyToDataTable();

                        lstNewUserAdd.DataSource = InsertedRecords;
                        lstNewUserAdd.DataBind();

                        lblNewRecordAddedCount.Text = "<h1> New Record Added : (" + RowCount.ToString() + ")</h1>";

                        pnlAddNewUser.Visible = true;
                        //Session["DuplicateUsers"] = ds.Tables[0];
                        //listDuplicateUsers.DataSource = ds.Tables[0];
                        //listDuplicateUsers.DataBind();
                    }

                    ScriptManager.RegisterStartupScript(this, this.GetType(), "overlay", "OverlayPopupUploadBulk();", true);
                }
                else
                {
                    //ScriptManager.RegisterStartupScript(this, this.GetType(), "overlay", "alert('All records has been added successfully!');window.location ='EditUser.aspx';", true);
                    UcStatusPopUp.ucPopUpMsg = "Kindly validate uploaded Data / File.";
                    UcStatusPopUp.changeText();
                    ScriptManager.RegisterStartupScript(this, this.GetType(), "Overlay", "showStatusChangePopUp();", true);
                }

                #endregion
            }
        }

        public static string Serialize(object dataToSerialize)
        {
            if (dataToSerialize == null) return null;

            using (StringWriter stringwriter = new System.IO.StringWriter())
            {
                var serializer = new XmlSerializer(dataToSerialize.GetType());
                serializer.Serialize(stringwriter, dataToSerialize, null);

                return stringwriter.ToString();
            }
        }

        public static T Deserialize<T>(string xmlText)
        {
            if (String.IsNullOrWhiteSpace(xmlText)) return default(T);

            using (StringReader stringReader = new System.IO.StringReader(xmlText))
            {
                var serializer = new XmlSerializer(typeof(T));
                return (T)serializer.Deserialize(stringReader);
            }
        }

        /// <summary>
        /// Public method Get value on the bease of Selected status and ID
        /// uses in ucStaffLoginAlert , EditUsers
        /// </summary>
        /// <param name="SelectedStatus"></param>
        /// <param name="Id"></param>
        /// <returns></returns>
        public bool CheckRequiredFields(string SelectedStatus, int Id)
        {
            DataSet dsNew = new DataSet();
            dsNew = InstallUserBLL.Instance.getuserdetails(Id);
            if (dsNew.Tables.Count > 0)
            {
                if (dsNew.Tables[0].Rows.Count > 0)
                {
                    if (SelectedStatus == "Applicant")
                    {
                        if (Convert.ToString(dsNew.Tables[0].Rows[0][1]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][2]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][3]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][8]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][38]) == "")
                        {
                            //ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('Status cannot be changed to Applicant as required fields for it are not filled.')", true);
                            return false;
                        }
                    }
                    else if (SelectedStatus == "OfferMade" || SelectedStatus == "Offer Made")
                    {
                        //if (Convert.ToString(dsNew.Tables[0].Rows[0][1]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][2]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][4]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][5]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][11]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][12]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][13]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][3]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][8]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][38]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][44]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][46]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][48]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][50]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][100]) == "")
                        if (Convert.ToString(dsNew.Tables[0].Rows[0]["Email"]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0]["Password"]) == "")
                        {
                            txtEmail.Text = Convert.ToString(dsNew.Tables[0].Rows[0]["Email"]);
                            //ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('Status cannot be changed to Offer Made as required fields for it are not filled.')", true);
                            return false;
                        }
                    }
                    else if (SelectedStatus == "Active")
                    {
                        if (Convert.ToString(dsNew.Tables[0].Rows[0][1]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][2]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][3]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][4]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][5]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][7]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][9]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][11]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][12]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][13]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][17]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][16]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][17]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][8]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][18]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][19]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][20]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][35]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][38]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][39]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][44]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][46]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][48]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][50]) == "" || Convert.ToString(dsNew.Tables[0].Rows[0][100]) == "")
                        {
                            //ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('Status cannot be changed to Offer Made as required fields for it are not filled.')", true); 
                            return false;
                        }
                    }
                }
            }
            return true;
        }

        private void SendEmail(string emailId, string FName, string LName, string status, string Reason, string Designition, int DesignitionId, string HireDate, string EmpType, string PayRates, HTMLTemplates objHTMLTemplateType, List<Attachment> Attachments = null, string strManager = "")
        {
            DesignationHTMLTemplate objHTMLTemplate = HTMLTemplateBLL.Instance.GetDesignationHTMLTemplate(objHTMLTemplateType, DesignitionId.ToString());

            //DataSet ds = AdminBLL.Instance.GetEmailTemplate(Designition, htmlTempID);// AdminBLL.Instance.FetchContractTemplate(104);
            //if (ds == null)
            //{
            //    ds = AdminBLL.Instance.GetEmailTemplate("Admin", htmlTempID);
            //}
            //else if (ds.Tables[0].Rows.Count == 0)
            //{
            //    ds = AdminBLL.Instance.GetEmailTemplate("Admin", htmlTempID);
            //}

            //string strHeader = ds.Tables[0].Rows[0]["HTMLHeader"].ToString(); //GetEmailHeader(status);
            //string strBody = ds.Tables[0].Rows[0]["HTMLBody"].ToString(); //GetEmailBody(status);
            //string strFooter = ds.Tables[0].Rows[0]["HTMLFooter"].ToString(); // GetFooter(status);
            //string strsubject = ds.Tables[0].Rows[0]["HTMLSubject"].ToString();

            string userName = ConfigurationManager.AppSettings["VendorCategoryUserName"].ToString();
            string password = ConfigurationManager.AppSettings["VendorCategoryPassword"].ToString();
            string fullname = FName + " " + LName;

            string strHeader = objHTMLTemplate.Header;
            string strBody = objHTMLTemplate.Body;
            string strFooter = objHTMLTemplate.Footer;
            string strsubject = objHTMLTemplate.Subject;

            strBody = strBody.Replace("#Email#", emailId).Replace("#email#", emailId);
            strBody = strBody.Replace("#FirstName#", FName);
            strBody = strBody.Replace("#LastName#", LName);
            strBody = strBody.Replace("#Name#", FName).Replace("#name#", FName);
            strBody = strBody.Replace("#Date#", dtInterviewDate.Text).Replace("#date#", dtInterviewDate.Text);
            strBody = strBody.Replace("#Time#", ddlInsteviewtime.SelectedValue).Replace("#time#", ddlInsteviewtime.SelectedValue);
            strBody = strBody.Replace("#Designation#", Designition).Replace("#designation#", Designition);

            strFooter = strFooter.Replace("#Name#", FName).Replace("#name#", FName);
            strFooter = strFooter.Replace("#Date#", dtInterviewDate.Text).Replace("#date#", dtInterviewDate.Text);
            strFooter = strFooter.Replace("#Time#", ddlInsteviewtime.SelectedValue).Replace("#time#", ddlInsteviewtime.SelectedValue);
            strFooter = strFooter.Replace("#Designation#", Designition).Replace("#designation#", Designition);

            strBody = strBody.Replace("Lbl Full name", fullname);
            strBody = strBody.Replace("LBL position", Designition);
            //strBody = strBody.Replace("lbl: start date", txtHireDate.Text);
            //strBody = strBody.Replace("($ rate","$"+ txtHireDate.Text);
            strBody = strBody.Replace("Reason", Reason);

            strBody = strBody.Replace("#manager#", strManager);

            strBody = strHeader + strBody + strFooter;

            //Hi #lblFName#, <br/><br/>You are requested to appear for an interview on #lblDate# - #lblTime#.<br/><br/>Regards,<br/>

            if (status == "OfferMade")
            {
                //TODO : commented code for missing directive using Word = Microsoft.Office.Interop.Word;
                //createForeMenForJobAcceptance(strBody, FName, LName, Designition, emailId, HireDate, EmpType, PayRates);
            }
            if (status == "Deactive")
            {
                //TODO : commented code for missing directive using Word = Microsoft.Office.Interop.Word;
                //CreateDeactivationAttachment(strBody, FName, LName, Designition, emailId, HireDate, EmpType, PayRates);
            }

            List<Attachment> lstAttachments = objHTMLTemplate.Attachments;

            //for (int i = 0; i < lstAttachments.Count; i++)
            //{
            //    string sourceDir = Server.MapPath(ds.Tables[1].Rows[i]["DocumentPath"].ToString());
            //    if (File.Exists(sourceDir))
            //    {
            //        Attachment attachment = new Attachment(sourceDir);
            //        attachment.Name = Path.GetFileName(sourceDir);
            //        lstAttachments.Add(attachment);
            //    }
            //}

            if (Attachments != null)
            {
                lstAttachments.AddRange(Attachments);
            }

            try
            {
                JG_Prospect.App_Code.CommonFunction.SendEmail(Designition, emailId, strsubject, strBody, lstAttachments);

                ScriptManager.RegisterStartupScript(this, this.GetType(), "UserMsg", "alert('An email notification has sent on " + emailId + ".');", true);
            }
            catch
            {
                ScriptManager.RegisterStartupScript(this, this.GetType(), "UserMsg", "alert('Error while sending email notification on " + emailId + ".');", true);
            }
        }

        private string GetFooter(string status)
        {
            string Footer = string.Empty;
            DataTable DtFooter;
            DtFooter = InstallUserBLL.Instance.getTemplate(status, "footer");
            if (DtFooter.Rows.Count > 0)
            {
                Footer = DtFooter.Rows[0][0].ToString();
            }
            return Footer;
        }

        private string GetEmailBody(string status)
        {
            string Body = string.Empty;
            DataTable DtBody;
            DtBody = InstallUserBLL.Instance.getTemplate(status, "Body");
            if (DtBody.Rows.Count > 0)
            {
                Body = DtBody.Rows[0][0].ToString();
            }
            return Body;
        }

        private string GetEmailHeader(string status)
        {
            string Header = string.Empty;
            DataTable DtHeader;
            DtHeader = InstallUserBLL.Instance.getTemplate(status, "Header");
            if (DtHeader.Rows.Count > 0)
            {
                Header = DtHeader.Rows[0][0].ToString();
            }
            return Header;
        }

        #region TODO : commented code for missing directive using Word = Microsoft.Office.Interop.Word;
        //private void FindAndReplace(Word.Application wordApp, object findText, object replaceText)
        //{
        //    object matchCase = true;
        //    object matchWholeWord = true;
        //    object matchWildCards = false;
        //    object matchSoundsLike = false;
        //    object matchAllWordForms = false;
        //    object forward = true;
        //    object format = false;
        //    object matchKashida = false;
        //    object matchDiacritics = false;
        //    object matchAlefHamza = false;
        //    object matchControl = false;
        //    object read_only = false;
        //    object visible = true;
        //    object replace = 2;
        //    object wrap = 1;
        //    wordApp.Selection.Find.Execute(ref findText, ref matchCase,
        //        ref matchWholeWord, ref matchWildCards, ref matchSoundsLike,
        //        ref matchAllWordForms, ref forward, ref wrap, ref format,
        //        ref replaceText, ref replace, ref matchKashida,
        //                ref matchDiacritics,
        //        ref matchAlefHamza, ref matchControl);
        //}

        //public void createForeMenForJobAcceptance(string str_Body, string FName, string LName, string Designition, string emailId, string HireDate, string EmpType, string PayRates)
        //{
        //    //copy sample file for Foreman Job Acceptance letter template
        //    string str_date = DateTime.Now.ToString().Replace("/", "");
        //    str_date = str_date.Replace(":", "");
        //    str_date = str_date.Replace("-", "");
        //    str_date = str_date.Replace(" ", "");
        //    string SourcePath = @"~/Sr_App/MailDocSample/ForemanJobAcceptancelettertemplate.docx";
        //    string TargetPath = @"~/Sr_App/MailDocument/" + str_date + FName + "ForemanJobAcceptanceletter.docx";
        //    System.IO.File.Copy(Server.MapPath(SourcePath), Server.MapPath(TargetPath), true);
        //    //modify word document
        //    object missing = System.Reflection.Missing.Value;
        //    Word.Application wordApp = new Word.Application();
        //    Word.Document aDoc = null;
        //    object Target = Server.MapPath(TargetPath);
        //    if (File.Exists(Server.MapPath(TargetPath)))
        //    {
        //        DateTime today = DateTime.Now;
        //        object readonlyNew = false;
        //        object isVisible = false;
        //        wordApp.Visible = false;
        //        FileInfo objFInfo = new FileInfo(Server.MapPath(TargetPath));
        //        objFInfo.IsReadOnly = false;
        //        aDoc = wordApp.Documents.Open(ref Target, ref missing, ref readonlyNew, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);
        //        aDoc.Activate();
        //        this.FindAndReplace(wordApp, "LBL Date", DateTime.Now.ToShortDateString());
        //        this.FindAndReplace(wordApp, "Lbl Full name", FName + " " + LName);
        //        this.FindAndReplace(wordApp, "LBL name", FName + " " + LName);
        //        this.FindAndReplace(wordApp, "LBL position", Designition);
        //        this.FindAndReplace(wordApp, "lbl fulltime", EmpType);
        //        this.FindAndReplace(wordApp, "lbl: start date", HireDate);
        //        this.FindAndReplace(wordApp, "$ rate", PayRates);
        //        this.FindAndReplace(wordApp, "lbl: next pay period", "");
        //        this.FindAndReplace(wordApp, "lbl: paycheck date", "");
        //        aDoc.SaveAs(ref Target, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
        //        aDoc.Close(ref missing, ref missing, ref missing);
        //    }
        //    using (MailMessage mm = new MailMessage("qat2015team@gmail.com", emailId))
        //    {
        //        try
        //        {
        //            mm.Subject = "Foreman Job Acceptance";
        //            mm.Body = str_Body;
        //            mm.Attachments.Add(new Attachment(Server.MapPath(TargetPath)));
        //            mm.IsBodyHtml = true;
        //            SmtpClient smtp = new SmtpClient();
        //            smtp.Host = "smtp.gmail.com";
        //            smtp.EnableSsl = true;
        //            NetworkCredential NetworkCred = new NetworkCredential("qat2015team@gmail.com", "q$7@wt%j*65ba#3M@9P6");
        //            smtp.UseDefaultCredentials = true;
        //            smtp.Credentials = NetworkCred;
        //            smtp.Port = 587;
        //            smtp.Send(mm);
        //        }
        //        catch (Exception ex)
        //        {
        //            ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('" + ex.Message + "')", true);
        //        }
        //        //ClientScript.RegisterStartupScript(GetType(), "alert", "alert('Email sent.');", true);
        //    }
        //}

        //public void CreateDeactivationAttachment(string MailBody, string FName, string LName, string Designition, string emailId, string HireDate, string EmpType, string PayRates)
        //{
        //    string str_date = DateTime.Now.ToString().Replace("/", "");
        //    str_date = str_date.Replace(":", "");
        //    str_date = str_date.Replace("-", "");
        //    str_date = str_date.Replace(" ", "");
        //    string SourcePath = @"~/Sr_App/MailDocSample/DeactivationMail.doc";
        //    string TargetPath = @"~/Sr_App/MailDocument/" + str_date + FName + "DeactivationMail.doc";
        //    System.IO.File.Copy(Server.MapPath(SourcePath), Server.MapPath(TargetPath), true);
        //    //modify word document
        //    object missing = System.Reflection.Missing.Value;
        //    Word.Application wordApp = new Word.Application();
        //    Word.Document aDoc = null;
        //    object Target = Server.MapPath(TargetPath);
        //    if (File.Exists(Server.MapPath(TargetPath)))
        //    {
        //        DateTime today = DateTime.Now;
        //        object readonlyNew = false;
        //        object isVisible = false;
        //        wordApp.Visible = false;
        //        FileInfo objFInfo = new FileInfo(Server.MapPath(TargetPath));
        //        objFInfo.IsReadOnly = false;
        //        aDoc = wordApp.Documents.Open(ref Target, ref missing, ref readonlyNew, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);
        //        aDoc.Activate();
        //        this.FindAndReplace(wordApp, "name", FName + " " + LName);
        //        this.FindAndReplace(wordApp, "HireDate", HireDate);
        //        this.FindAndReplace(wordApp, "full time or part  time", EmpType);
        //        this.FindAndReplace(wordApp, "HourlyRate", PayRates);
        //        this.FindAndReplace(wordApp, "WorkingStatus", "No");
        //        this.FindAndReplace(wordApp, "LastWorkingDay", DateTime.Now.ToShortDateString());
        //        aDoc.SaveAs(ref Target, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
        //        aDoc.Close(ref missing, ref missing, ref missing);
        //    }
        //    using (MailMessage mm = new MailMessage("qat2015team@gmail.com", emailId))
        //    {
        //        try
        //        {
        //            mm.Subject = "Deactivation";
        //            mm.Body = MailBody;
        //            mm.Attachments.Add(new Attachment(Server.MapPath(TargetPath)));
        //            mm.IsBodyHtml = true;
        //            SmtpClient smtp = new SmtpClient();
        //            smtp.Host = "smtp.gmail.com";
        //            smtp.EnableSsl = true;
        //            NetworkCredential NetworkCred = new NetworkCredential("qat2015team@gmail.com", "q$7@wt%j*65ba#3M@9P6");
        //            smtp.UseDefaultCredentials = true;
        //            smtp.Credentials = NetworkCred;
        //            smtp.Port = 587;
        //            smtp.Send(mm);
        //        }
        //        catch (Exception ex)
        //        {
        //            ScriptManager.RegisterStartupScript(this, this.GetType(), "alert", "alert('" + ex.Message + "')", true);
        //        }
        //    }
        //}
        #endregion
        public List<string> GetTimeIntervals()
        {
            List<string> timeIntervals = new List<string>();
            TimeSpan startTime = new TimeSpan(0, 0, 0);
            DateTime startDate = new DateTime(DateTime.MinValue.Ticks); // Date to be used to get shortTime format.
            for (int i = 0; i < 48; i++)
            {
                int minutesToBeAdded = 30 * i;      // Increasing minutes by 30 minutes interval
                TimeSpan timeToBeAdded = new TimeSpan(0, minutesToBeAdded, 0);
                TimeSpan t = startTime.Add(timeToBeAdded);
                DateTime result = startDate + t;
                timeIntervals.Add(result.ToShortTimeString());      // Use Date.ToShortTimeString() method to get the desired format                
            }
            return timeIntervals;
        }

        /// <summary>
        /// Fill ddl for User Recruter 
        /// Also call from ucStaffLogin , EditUser
        /// </summary>
        private void LoadUsersByRecruiterDesgination(DropDownList ddlUsers)
        {
            ddlUsers.SelectedIndex = -1;
            ddlUsers.Items.Clear();

            DataSet dsUsers = TaskGeneratorBLL.Instance.GetInstallUsers(2, "Admin,Admin Recruiter,Office Manager,Recruiter,ITLead,");
            if (dsUsers != null && dsUsers.Tables.Count > 0)
            {
                DataView dvUsers = dsUsers.Tables[0].DefaultView;
                dvUsers.RowFilter = string.Format(
                                                    "[Status] IN ('{0}','{1}')",
                                                    Convert.ToByte(JGConstant.InstallUserStatus.Active).ToString(),
                                                    Convert.ToByte(JGConstant.InstallUserStatus.OfferMade).ToString()
                                                );
                dvUsers.Sort = "[Status] ASC";

                DataTable dtUsers = dvUsers.ToTable();

                for (int i = 0; i < dtUsers.Rows.Count; i++)
                {
                    DataRow objUser = dtUsers.Rows[i];
                    ddlUsers.Items.Add(new ListItem(objUser["FristName"].ToString(), objUser["Id"].ToString()));
                    if (objUser["Status"].ToString() == Convert.ToByte(JGConstant.InstallUserStatus.Active).ToString())
                    {
                        ddlUsers.Items[i].Attributes.Add("style", "color: red;");
                    }
                }
                //ddlUsers.DataSource = dvUsers.ToTable();
                //ddlUsers.DataTextField = "FristName";
                //ddlUsers.DataValueField = "Id";
                //ddlUsers.DataBind();
            }
            ddlUsers.Items.Insert(0, new ListItem("--All--", "0"));
        }

        private void FillTechTaskDropDown(DropDownList ddlTechTask, DropDownList ddlTechSubTask, int intDesignationId)
        {
            ddlTechTask.ClearSelection();
            ddlTechTask.Items.Clear();

            ddlTechSubTask.ClearSelection();
            ddlTechSubTask.Items.Clear();

            DataSet dsTechTask;

            if (intDesignationId == 0)
            {
                dsTechTask = TaskGeneratorBLL.Instance.GetAllActiveTechTask();
            }
            else
            {
                dsTechTask = TaskGeneratorBLL.Instance.GetAllActiveTechTaskForDesignationID(intDesignationId);
            }

            if (dsTechTask != null & dsTechTask.Tables.Count > 0)
            {
                DataTable dtTechTask = dsTechTask.Tables[0];
                //dtTechTask.Columns.Add("TitleWithLink");
                //for (int iCurrentRow = 0; iCurrentRow < dtTechTask.Rows.Count; iCurrentRow++)
                //{
                //    dtTechTask.Rows[iCurrentRow]["TitleWithLink"] = dtTechTask.Rows[iCurrentRow]["Title"] + " - <a href=TaskGenerator.aspx?TaskId=" + dtTechTask.Rows[iCurrentRow]["TaskId"] + ">" + dtTechTask.Rows[iCurrentRow]["TaskId"] +"</a>";
                //}

                ddlTechTask.DataSource = dtTechTask;
                ddlTechTask.DataTextField = "Title";
                ddlTechTask.DataValueField = "TaskId";
                ddlTechTask.DataBind();
            }
            ddlTechTask.Items.Insert(0, new ListItem("--select--", "0"));
            ddlTechTask.SelectedValue = "0";

            ddlTechSubTask.Items.Insert(0, new ListItem("--select--", "0"));
            ddlTechSubTask.SelectedValue = "0";
        }

        private void FillSubTaskDropDown(long intTaskId, DropDownList ddlTechSubTask)
        {
            DataSet dsSubTasks = TaskGeneratorBLL.Instance.GetTaskHierarchy(intTaskId, CommonFunction.CheckAdminAndItLeadMode());

            DataTable dtLevel1SubTasks = null;
            DataTable dtLevel2SubTasks = null;
            DataTable dtLevel3SubTasks = null;

            if (dsSubTasks != null && dsSubTasks.Tables.Count > 0 && dsSubTasks.Tables[0].Rows.Count > 0)
            {
                DataView dvSubTasks = dsSubTasks.Tables[0].DefaultView;

                dvSubTasks.RowFilter = string.Format("TaskId <> {0} AND TaskLevel = 1", intTaskId);
                dtLevel1SubTasks = dvSubTasks.ToTable();

                foreach (DataRow drSubTask1 in dtLevel1SubTasks.Rows)
                {
                    string strTaskId = Convert.ToString(drSubTask1["TaskId"]);
                    string strTitle = Convert.ToString(drSubTask1["Title"]);

                    ddlTechSubTask.Items.Add(new ListItem(strTitle, strTaskId));

                    dvSubTasks.RowFilter = string.Format("ParentTaskId = {0} AND TaskLevel = 2", strTaskId);
                    dtLevel2SubTasks = dvSubTasks.ToTable();

                    foreach (DataRow drSubTask2 in dtLevel2SubTasks.Rows)
                    {
                        strTaskId = Convert.ToString(drSubTask2["TaskId"]);
                        strTitle = "|--" + Convert.ToString(drSubTask2["Title"]);

                        ddlTechSubTask.Items.Add(new ListItem(strTitle, strTaskId));

                        dvSubTasks.RowFilter = string.Format("ParentTaskId = {0} AND TaskLevel = 3", strTaskId);
                        dtLevel3SubTasks = dvSubTasks.ToTable();

                        foreach (DataRow drSubTask3 in dtLevel3SubTasks.Rows)
                        {
                            strTaskId = Convert.ToString(drSubTask3["TaskId"]);
                            strTitle = "|----" + Convert.ToString(drSubTask3["Title"]);

                            ddlTechSubTask.Items.Add(new ListItem(strTitle, strTaskId));
                        }
                    }
                }
            }
            else
            {
                ddlTechSubTask.DataSource = null;
                ddlTechSubTask.DataTextField = "Title";
                ddlTechSubTask.DataValueField = "TaskId";
                ddlTechSubTask.DataBind();
            }
        }

        //private void BindGrid()
        //{
        //    GetSalesUsersStaticticsAndData();

        //    DataTable dt = (DataTable)(Session["UserGridData"]);
        //    EnumerableRowCollection<DataRow> query = null;
        //    int iSelectedDesignationID = 0;
        //    int iAddedByUserID = 0;
        //    int iSourceID = 0;
        //    if (ddlDesignation.SelectedIndex > 0)
        //    {
        //        iSelectedDesignationID = string.IsNullOrEmpty(ddlDesignation.SelectedValue) ? 0 : Convert.ToInt32(ddlDesignation.SelectedValue);
        //    }
        //    //if (drpUser.SelectedIndex > 0)
        //    //{
        //    //    iAddedByUserID = string.IsNullOrEmpty(drpUser.SelectedValue) ? 0 : Convert.ToInt32(drpUser.SelectedValue);
        //    //}
        //    //if (ddlSource.SelectedIndex > 0)
        //    //{
        //    //    iSourceID = string.IsNullOrEmpty(ddlSource.SelectedValue) ? 0 : Convert.ToInt32(ddlSource.SelectedValue);
        //    //}
        //    if ((ddlUserStatus.SelectedIndex != 0 || ddlDesignation.SelectedIndex != 0 || drpUser.SelectedIndex != 0 || ddlSource.SelectedIndex != 0)
        //        && dt != null)
        //    {
        //        string Status = ddlUserStatus.SelectedItem.Value;
        //        query = from userdata in dt.AsEnumerable()
        //                where (userdata.Field<string>("Status") == Status || ddlUserStatus.SelectedIndex == 0)
        //                && (userdata.Field<Int32?>("DesignationID") == iSelectedDesignationID || ddlDesignation.SelectedIndex == 0)
        //                && (userdata.Field<Int32?>("AddedById") == Convert.ToInt32(drpUser.SelectedValue) || drpUser.SelectedIndex == 0)
        //                && (userdata.Field<string>("Source") == ddlSource.SelectedValue || ddlSource.SelectedIndex == 0)
        //                select userdata;
        //        if (query.Count() > 0)
        //        {
        //            dt = query.CopyToDataTable();
        //        }
        //        else
        //            dt = null;
        //    }
        //    //grdUsers.DataSource = dt;
        //    //grdUsers.DataBind();

        //    BindUsers(dt);
        //}

        private void GetSalesUsersStaticticsAndData(bool blResetGrid = false)
        {
            if (blResetGrid)
            {
                grdUsers.PageIndex = 0;
                //this.SalesUserSortDirection = SortDirection.Descending;
                //this.SalesUserSortExpression = "CreatedDateTime";
            }

            DateTime? dtFromDate = null;
            DateTime? dtToDate = null;
            if (!chkAllDates.Checked)
            {
                dtFromDate = Convert.ToDateTime(txtfrmdate.Text, JG_Prospect.Common.JGConstant.CULTURE);
                dtToDate = Convert.ToDateTime(txtTodate.Text, JG_Prospect.Common.JGConstant.CULTURE);
            }
            string strUserStatus = string.Empty;
            if (dtFromDate < dtToDate || (dtFromDate == null && dtToDate == null))
            {
                string strSortExpression = this.SalesUserSortExpression + " " + (this.SalesUserSortDirection == SortDirection.Ascending ? "ASC" : "DESC");

                DataSet dsSalesUserData = InstallUserBLL.Instance.GetSalesUsersStaticticsAndData
                                                        (
                                                            txtSearch.Text.Trim(),
                                                            ddlUserStatus.SelectedValue,
                                                            Convert.ToInt32(ddlDesignation.SelectedValue),
                                                            Convert.ToInt32(ddlSource.SelectedValue),
                                                            dtFromDate,
                                                            dtToDate,
                                                            Convert.ToInt32(drpUser.SelectedValue),
                                                            grdUsers.PageIndex,
                                                            grdUsers.PageSize,
                                                            strSortExpression
                                                        );
                if (dsSalesUserData != null)
                {
                    DataTable dtSalesUser_Statictics_Status = dsSalesUserData.Tables[0];
                    DataTable dtSalesUser_Statictics_AddedBy = dsSalesUserData.Tables[1];
                    DataTable dtSalesUser_Statictics_Designation = dsSalesUserData.Tables[2];
                    DataTable dtSalesUser_Statictics_Source = dsSalesUserData.Tables[3];
                    DataTable dtSalesUser_Grid = dsSalesUserData.Tables[4];

                    #region OrderStatus Column

                    string usertype = Session["usertype"].ToString().ToLower();

                    if (dtSalesUser_Grid.Columns["OrderStatus"] == null)
                    {
                        dtSalesUser_Grid.Columns.Add("OrderStatus");
                        foreach (DataRow dr in dtSalesUser_Grid.Rows)
                        {
                            dr["OrderStatus"] = dr["Status"].ToString().Replace(" ", "");
                        }
                        //int st = 0;

                        //if (usertype == "jg account" || usertype == "sales manager" || usertype == "office manager" || usertype == "recruiter")
                        //{
                        //    foreach (DataRow dr in dtSalesUser_Grid.Rows)
                        //    {
                        //        st = (int)((OrderStatus1)Enum.Parse(typeof(OrderStatus1), dr["Status"].ToString().Replace(" ", "")));
                        //        dr["OrderStatus"] = st.ToString();
                        //    }
                        //}
                        //else if (usertype == "admin" || usertype == "jr. sales" || usertype == "project manager")
                        //{
                        //    foreach (DataRow dr in dtSalesUser_Grid.Rows)
                        //    {
                        //        st = (int)((OrderStatus2)Enum.Parse(typeof(OrderStatus2), dr["Status"].ToString().Replace(" ", "")));
                        //        dr["OrderStatus"] = st.ToString();
                        //    }
                        //}
                    }

                    #endregion

                    #region Statictics

                    if (dtSalesUser_Statictics_Status.Rows.Count > 0)
                    {
                        List<HrData> lstHrData = new List<HrData>();
                        foreach (DataRow row in dtSalesUser_Statictics_Status.Rows)
                        {
                            HrData hrdata = new HrData();
                            hrdata.status = row["Status"].ToString();
                            hrdata.count = row["Count"].ToString();
                            lstHrData.Add(hrdata);
                        }

                        var rowOfferMade = lstHrData.Where(r => r.status == Convert.ToByte(JGConstant.InstallUserStatus.OfferMade).ToString()).FirstOrDefault();
                        if (rowOfferMade != null)
                        {
                            string count = rowOfferMade.count;
                            lbljoboffercount.Text = count;
                        }
                        else
                        {
                            lbljoboffercount.Text = "0";
                        }
                        var rowActive = lstHrData.Where(r => r.status == Convert.ToByte(JGConstant.InstallUserStatus.Active).ToString()).FirstOrDefault();
                        if (rowActive != null)
                        {
                            string count = rowActive.count;
                            lblActiveCount.Text = count;
                        }
                        else
                        {
                            lblActiveCount.Text = "0";
                        }
                        var rowRejected = lstHrData.Where(r => r.status == Convert.ToByte(JGConstant.InstallUserStatus.Rejected).ToString()).FirstOrDefault();
                        if (rowRejected != null)
                        {
                            string count = rowRejected.count;
                            lblRejectedCount.Text = count;
                        }
                        else
                        {
                            lblRejectedCount.Text = "0";
                        }
                        var rowDeactive = lstHrData.Where(r => r.status == Convert.ToByte(JGConstant.InstallUserStatus.Deactive).ToString()).FirstOrDefault();
                        if (rowDeactive != null)
                        {
                            string count = rowDeactive.count;
                            lblDeactivatedCount.Text = count;
                        }
                        else
                        {
                            lblDeactivatedCount.Text = "0";
                        }
                        var rowInstallProspect = lstHrData.Where(r => r.status == Convert.ToByte(JGConstant.InstallUserStatus.InstallProspect).ToString()).FirstOrDefault();
                        if (rowInstallProspect != null)
                        {
                            string count = rowInstallProspect.count;
                            lblInstallProspectCount.Text = count;
                        }
                        else
                        {
                            lblInstallProspectCount.Text = "0";
                        }
                        var rowPhoneScreened = lstHrData.Where(r => r.status == Convert.ToByte(JGConstant.InstallUserStatus.PhoneScreened).ToString()).FirstOrDefault();
                        if (rowPhoneScreened != null)
                        {
                            string count = rowPhoneScreened.count;
                            lblPhoneVideoScreenedCount.Text = count;
                        }
                        else
                        {
                            lblPhoneVideoScreenedCount.Text = "0";
                        }
                        var rowInterviewDate = lstHrData.Where(r => r.status == Convert.ToByte(JGConstant.InstallUserStatus.InterviewDate).ToString()).FirstOrDefault();
                        if (rowInterviewDate != null)
                        {
                            string count = rowInterviewDate.count;
                            lblInterviewDateCount.Text = count;
                        }
                        else
                        {
                            lblInterviewDateCount.Text = "0";
                        }
                        var rowApplicant = lstHrData.Where(r => r.status == Convert.ToByte(JGConstant.InstallUserStatus.Applicant).ToString()).FirstOrDefault();
                        string Applicantcount = "0";
                        if (rowApplicant != null)
                        {
                            Applicantcount = rowApplicant.count;

                        }
                        else
                        {
                            Applicantcount = "0";
                        }

                        lblNewApplicantsCount.Text = Convert.ToDouble(Applicantcount).ToString();
                        // Ratio Calculation
                        lblAppInterviewRatio.Text = Convert.ToString(Convert.ToDouble(lblInterviewDateCount.Text) / Convert.ToDouble(Applicantcount));
                        //lblAppHireRatio.Text = Convert.ToString(Convert.ToDouble(lblActiveCount.Text) / Convert.ToDouble(Applicantcount) );
                        //lblJobOfferHireRatio.Text = Convert.ToString(Convert.ToDouble(lblActive.Text) / Convert.ToDouble(lblInterviewDateCount.Text));
                        if (lblInterviewDateCount.Text != "0")
                        {
                            lblInterviewActiveRatio.Text = Convert.ToString(Convert.ToDouble(lblActiveCount.Text) / Convert.ToDouble(lblInterviewDateCount.Text));
                        }
                        else
                        {
                            lblInterviewActiveRatio.Text = "0";
                        }
                        if (lbljoboffercount.Text != "0")
                        {
                            lblJobOfferActiveRatio.Text = Convert.ToString(Convert.ToDouble(lblActiveCount.Text) / Convert.ToDouble(lbljoboffercount.Text));
                        }
                        else
                        {
                            lblJobOfferActiveRatio.Text = "0";
                        }
                        if (lblActiveCount.Text != "0")
                        {
                            lblActiveDeactiveRatio.Text = Convert.ToString(Convert.ToDouble(lblDeactivatedCount.Text) / Convert.ToDouble(lblActiveCount.Text));
                        }
                        else
                        {
                            lblActiveDeactiveRatio.Text = "0";
                        }

                        BindPieChart(lstHrData);
                    }
                    else
                    {
                        lbljoboffercount.Text = "0";
                        lblActiveCount.Text = "0";
                        lblRejectedCount.Text = "0";
                        lblDeactivatedCount.Text = "0";
                        lblInstallProspectCount.Text = "0";
                        lblPhoneVideoScreenedCount.Text = "0";
                        lblInterviewDateCount.Text = "0";
                        lblAppInterviewRatio.Text = "0";
                        //  lblAppHireRatio.Text = "0";
                    }
                    #endregion

                    if (dtSalesUser_Grid.Rows.Count > 0)
                    {
                        //Session["UserGridData"] = dtSalesUser_Grid;
                        //BindUsers(dtSalesUser_Grid);

                        grdUsers.DataSource = dtSalesUser_Grid;
                        grdUsers.VirtualItemCount = Convert.ToInt32(dsSalesUserData.Tables[5].Rows[0]["TotalRecordCount"]);
                        grdUsers.DataBind();
                        grdUsers.UseAccessibleHeader = true;
                        grdUsers.HeaderRow.TableSection = TableRowSection.TableHeader;
                        BindUsersCount(dtSalesUser_Statictics_AddedBy, dtSalesUser_Statictics_Designation, dtSalesUser_Statictics_Source);
                    }
                    else
                    {
                        //Session["UserGridData"] = null;
                        grdUsers.DataSource = null;
                        grdUsers.DataBind();
                    }

                    upUsers.Update();
                }
            }
            else
            {
                ScriptManager.RegisterStartupScript(this, this.GetType(), "AlertBox", "alert('ToDate must be greater than FromDate');", true);
            }
        }

        private void BindPieChart(List<HrData> lstHrData)
        {
            string[] x = new string[lstHrData.Count()];
            int[] y = new int[lstHrData.Count()];

            for (int i = 0; i < lstHrData.Count(); i++)
            {
                JGConstant.InstallUserStatus status = (JGConstant.InstallUserStatus)Convert.ToInt32(lstHrData[i].status.ToString());
                x[i] = status.ToString();
                y[i] = Convert.ToInt32(lstHrData[i].count);
            }

            Chart1.Series[0].Points.DataBindXY(x, y);
            Chart1.Series[0].ChartType = SeriesChartType.Pie;
            Chart1.ChartAreas["ChartArea1"].Area3DStyle.Enable3D = true;
            Chart1.Legends[0].Enabled = true;
        }

        //private void BindPieChart(DataTable dtgridData)
        //{
        //    DataTable dt = dtgridData;

        //    var query = from row in dt.AsEnumerable()
        //                group row by row.Field<string>("status") into st
        //                orderby st.Key
        //                select new
        //                {
        //                    Name = st.Key,
        //                    Total = st.Count()
        //                };

        //    DataTable newItems = new DataTable();
        //    newItems.Columns.Add("Name");
        //    newItems.Columns.Add("Total");

        //    foreach (var item in query)
        //    {
        //        DataRow newRow = newItems.NewRow();
        //        newRow["Name"] = item.Name;
        //        newRow["Total"] = item.Total;

        //        newItems.Rows.Add(newRow);
        //    }

        //    string[] x = new string[query.Count()];
        //    int[] y = new int[query.Count()];

        //    for (int i = 0; i < query.Count(); i++)
        //    {

        //        x[i] = newItems.Rows[i]["Name"].ToString();
        //        y[i] = Convert.ToInt32(newItems.Rows[i]["Total"]);
        //    }

        //    Chart1.Series[0].Points.DataBindXY(x, y);
        //    Chart1.Series[0].ChartType = SeriesChartType.Pie;
        //    Chart1.ChartAreas["ChartArea1"].Area3DStyle.Enable3D = true;
        //    Chart1.Legends[0].Enabled = true;
        //}

        private void BindUsersCount(DataTable dtAddedBy, DataTable dtDesignation, DataTable dtSource)
        {
            //var addedBy = from row in dt.AsEnumerable()
            //              group row by row.Field<string>("AddedBy") into st
            //              orderby st.Key
            //              select new
            //              {
            //                  AddedBy = st.Key,
            //                  Count = st.Count()
            //              };

            listAddedBy.DataSource = dtAddedBy;
            listAddedBy.DataBind();

            //var desig = from row in dt.AsEnumerable()
            //            group row by row.Field<string>("Designation") into st
            //            orderby st.Key
            //            select new
            //            {
            //                Designation = st.Key,
            //                Count = st.Count()
            //            };

            listDesignation.DataSource = dtDesignation;
            listDesignation.DataBind();

            //var source = from row in dt.AsEnumerable()
            //             group row by row.Field<string>("Source") into st
            //             orderby st.Key
            //             select new
            //             {
            //                 Source = st.Key,
            //                 Count = st.Count()
            //             };

            listSource.DataSource = dtSource;
            listSource.DataBind();
        }

        //private void bindPayPeriod(DataSet dsCurrentPeriod)
        //{
        //    DataSet ds = UserBLL.Instance.getallperiod();

        //    if (ds.Tables[0].Rows.Count > 0)
        //    {
        //        drpPayPeriod.Items.Insert(0, new ListItem("Select", "0"));
        //        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
        //        {
        //            DataRow dr = ds.Tables[0].Rows[i];
        //            drpPayPeriod.Items.Add(new ListItem(dr["Periodname"].ToString(), dr["Id"].ToString()));
        //        }
        //        drpPayPeriod.SelectedValue = dsCurrentPeriod.Tables[0].Rows[0]["Id"].ToString();
        //        txtfrmdate.Text = Convert.ToDateTime(dsCurrentPeriod.Tables[0].Rows[0]["FromDate"].ToString()).ToString("MM/dd/yyyy");
        //        txtTodate.Text = Convert.ToDateTime(dsCurrentPeriod.Tables[0].Rows[0]["ToDate"].ToString()).ToString("MM/dd/yyyy");
        //    }
        //    else
        //    {
        //        drpPayPeriod.DataSource = null;
        //        drpPayPeriod.DataBind();
        //    }

        //}

        //protected void drpPayPeriod_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    if (drpPayPeriod.SelectedIndex != -1)
        //    {
        //        DataSet ds = UserBLL.Instance.getperioddetails(Convert.ToInt16(drpPayPeriod.SelectedValue));
        //        if (ds.Tables[0].Rows.Count > 0)
        //        {
        //            txtfrmdate.Text = Convert.ToDateTime(ds.Tables[0].Rows[0]["FromDate"].ToString()).ToString("MM/dd/yyyy");
        //            txtTodate.Text = Convert.ToDateTime(ds.Tables[0].Rows[0]["ToDate"].ToString()).ToString("MM/dd/yyyy");
        //        }
        //    }
        //}

        //private string GetId(string UserType, string UserStatus)
        //{
        //    DataTable dtId;
        //    string installId = string.Empty;

        //    string newId = string.Empty;
        //    dtId = InstallUserBLL.Instance.getMaxId(UserType, UserStatus);
        //    if (dtId.Rows.Count > 0)
        //    {
        //        installId = Convert.ToString(dtId.Rows[0][0]);
        //    }
        //    if ((UserType == "ForeMan" || UserType == "Installer") && (UserStatus == "Applicant" || UserStatus == "InterviewDate" || UserStatus == "OfferMade" || UserStatus == "PhoneScreened" || UserStatus == "Rejected"))
        //    {
        //        if (installId != "")
        //        {
        //            newId = installId.Substring(10);
        //            newId = Convert.ToString(Convert.ToUInt32(newId) + 1);
        //            installId = installId.Remove(10);
        //            installId = installId + newId;
        //        }
        //        else
        //        {
        //            installId = "P-OPP-00001";
        //        }
        //    }
        //    else if ((UserType == "ForeMan" || UserType == "Installer") && (UserStatus == "Active"))
        //    {
        //        if (installId != "")
        //        {
        //            newId = installId.Substring(8);
        //            newId = Convert.ToString(Convert.ToUInt32(newId) + 1);
        //            installId = installId.Remove(8);
        //            installId = installId + newId;
        //        }
        //        else
        //        {
        //            installId = "OPP-00001";
        //        }
        //    }
        //    else if ((UserType == "ForeMan" || UserType == "Installer") && (UserStatus == "Deactive"))
        //    {
        //        if (installId != "")
        //        {
        //            newId = installId.Substring(10);
        //            newId = Convert.ToString(Convert.ToUInt32(newId) + 1);
        //            installId = installId.Remove(10);
        //            installId = installId + newId;
        //        }
        //        else
        //        {
        //            installId = "X-OPP-00001";
        //        }
        //    }
        //    else if ((UserType == "SubContractor") && (UserStatus == "Applicant" || UserStatus == "InterviewDate" || UserStatus == "OfferMade" || UserStatus == "PhoneScreened" || UserStatus == "Rejected"))
        //    {
        //        if (installId != "")
        //        {
        //            newId = installId.Substring(8);
        //            newId = Convert.ToString(Convert.ToUInt32(newId) + 1);
        //            installId = installId.Remove(8);
        //            installId = installId + newId;
        //        }
        //        else
        //        {
        //            installId = "P-SC-00001";
        //        }
        //    }
        //    else if ((UserType == "SubContractor") && (UserStatus == "Active"))
        //    {
        //        if (installId != "")
        //        {
        //            newId = installId.Substring(6);
        //            newId = Convert.ToString(Convert.ToUInt32(newId) + 1);
        //            installId = installId.Remove(6);
        //            installId = installId + newId;
        //        }
        //        else
        //        {
        //            installId = "SC-00001";
        //        }
        //    }
        //    else if ((UserType == "SubContractor") && (UserStatus == "Deactive"))
        //    {
        //        if (installId != "")
        //        {
        //            newId = installId.Substring(8);
        //            newId = Convert.ToString(Convert.ToUInt32(newId) + 1);
        //            installId = installId.Remove(8);
        //            installId = installId + newId;
        //        }
        //        else
        //        {
        //            installId = "X-SC-00001";
        //        }
        //    }
        //    Session["installId"] = installId;
        //    return installId;
        //}

        private void LoadEmailContentToSentToUser(GridViewRow objUserRow)
        {
            DesignationHTMLTemplate objHTMLTemplate = HTMLTemplateBLL.Instance.GetDesignationHTMLTemplate(HTMLTemplates.InterviewDateAutoEmail, JGSession.DesignationId.ToString());

            txtEmailSubject.Text = objHTMLTemplate.Subject;
            txtEmailBody.Text = string.Concat(
                //objHTMLTemplate.Designation + " --- ",
                                                objHTMLTemplate.Header,
                                                objHTMLTemplate.Body,
                                                objHTMLTemplate.Footer
                                            );

            string strEmail = ((LinkButton)objUserRow.FindControl("lbtnEmail")).Text;
            string strFirstName = ((Label)objUserRow.FindControl("lblFirstName")).Text;
            string strLastName = ((Label)objUserRow.FindControl("lblLastName")).Text;
            string strDesignation = ((Label)objUserRow.FindControl("lblDesignation")).Text;

            string strBody = txtEmailBody.Text;
            strBody = strBody.Replace("#Email#", strEmail).Replace("#email#", strEmail);
            strBody = strBody.Replace("#FirstName#", strFirstName);
            strBody = strBody.Replace("#LastName#", strLastName);
            strBody = strBody.Replace("#Name#", strFirstName).Replace("#name#", strFirstName);
            strBody = strBody.Replace("#Designation#", strDesignation).Replace("#designation#", strDesignation);

            txtEmailBody.Text = strBody;

            txtEmailCustomMessage.Text = string.Empty;

            hdnEmailTo.Value =
            lblEmailTo.Text = strEmail;

            upSendEmailToUser.Update();

            ScriptManager.RegisterStartupScript
                                (
                                    this,
                                    this.GetType(),
                                    "ShowPopup_divSendEmailToUser",
                                    string.Concat(
                                                    "SetCKEditor('", txtEmailBody.ClientID, "');",
                                                    "SetCKEditor('", txtEmailCustomMessage.ClientID, "');",
                                                    "ShowPopupWithTitle('#", divSendEmailToUser.ClientID, "','Send Email');"
                                                 ),
                                    true
                                );
        }

        public string grdBulkUploadUserErrors_GetCellText(object strInput)
        {
            if (!string.IsNullOrEmpty(Convert.ToString(strInput)))
            {
                return strInput.ToString();
            }
            else
            {
                return "<div style=\"color: blue;font-weight:bold;text-align: center;font-size: 20px;\">x</div>";
            }
        }

        private void BookmarkUnStarInstallUser(int UserId, int BookmarkById, string action)
        {
            InstallUserBLL.Instance.BookmarkUnStarInstallUser(UserId, BookmarkById, action);
        }

        /// <summary>
        /// Add a GitHub user as Collaborator in repo        
        /// </summary>
        /// <param name="gitUserName"></param>
        private async void AddUserAsGitcollaborator(string gitUserName)
        {
            try
            {
                var reponame = ConfigurationManager.AppSettings["GitRepoName"].ToString();
                var adminname = ConfigurationManager.AppSettings["GitRepoAdminName"].ToString();
                var adminloginId = ConfigurationManager.AppSettings["GitRepoAdminLoginId"].ToString();
                var adminpassword = ConfigurationManager.AppSettings["GitRepoAdminPassword"].ToString();
                var client = new Octokit.GitHubClient(new Octokit.ProductHeaderValue("my-cool-app"));
                var basicAuth = new Octokit.Credentials(adminloginId, adminpassword);
                client.Credentials = basicAuth;
                await client.Repository.Collaborator.Add(adminname, reponame, gitUserName);
            }
            catch (Exception ex)
            {
                //Log exception 
                Console.WriteLine("{0} Exception caught.", ex);
            }
        }

        #region 'Assigned Task ToUser'

        private void AssignedTaskToUser(int intEditId, DropDownList ddlTechTask, DropDownList ddlTechSubTask)
        {
            string ApplicantId = intEditId.ToString();

            //If dropdown has any value then assigned it to user else. return 
            if (ddlTechTask.Items.Count > 0)
            {
                // save (insert / delete) assigned users.
                //bool isSuccessful = TaskGeneratorBLL.Instance.SaveTaskAssignedUsers(Convert.ToUInt64(ddlTechTask.SelectedValue), Session["EditId"].ToString());

                // save assigned user a TASK.
                bool isSuccessful = TaskGeneratorBLL.Instance.SaveTaskAssignedToMultipleUsers(Convert.ToUInt64(ddlTechSubTask.SelectedValue), ApplicantId);

                // Change task status to assigned = 3.
                if (isSuccessful)
                    UpdateTaskStatus(Convert.ToInt32(ddlTechSubTask.SelectedValue), Convert.ToUInt16(JGConstant.TaskStatus.Assigned));

                if (ddlTechTask.SelectedValue != "" || ddlTechTask.SelectedValue != "0")
                    SendEmailToAssignedUsers(ApplicantId, ddlTechTask.SelectedValue, ddlTechSubTask.SelectedValue, ddlTechTask.SelectedItem.Text);
            }
        }

        private void UpdateTaskStatus(Int32 taskId, UInt16 Status)
        {
            Task task = new Task();
            task.TaskId = taskId;
            task.Status = Status;

            int result = TaskGeneratorBLL.Instance.UpdateTaskStatus(task);    // save task master details

            //String AlertMsg;

            //if (result > 0)
            //{
            //    AlertMsg = "Status changed successfully!";
            //}
            //else
            //{
            //    AlertMsg = "Status change was not successfull, Please try again later.";
            //}
        }

        private void SendEmailToAssignedUsers(string strInstallUserIDs, string strTaskId, string strSubTaskId, string strTaskTitle)
        {
            try
            {
                //string strHTMLTemplateName = "Task Generator Auto Email";
                //DataSet dsEmailTemplate = AdminBLL.Instance.GetEmailTemplate(strHTMLTemplateName, 108);

                DesignationHTMLTemplate objHTMLTemplate = HTMLTemplateBLL.Instance.GetDesignationHTMLTemplate(HTMLTemplates.Task_Generator_Auto_Email, JGSession.DesignationId.ToString());

                foreach (string userID in strInstallUserIDs.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    DataSet dsUser = TaskGeneratorBLL.Instance.GetInstallUserDetails(Convert.ToInt32(userID));

                    string emailId = dsUser.Tables[0].Rows[0]["Email"].ToString();
                    string FName = dsUser.Tables[0].Rows[0]["FristName"].ToString();
                    string LName = dsUser.Tables[0].Rows[0]["LastName"].ToString();
                    string fullname = FName + " " + LName;

                    //string strHeader = dsEmailTemplate.Tables[0].Rows[0]["HTMLHeader"].ToString();
                    //string strBody = dsEmailTemplate.Tables[0].Rows[0]["HTMLBody"].ToString();
                    //string strFooter = dsEmailTemplate.Tables[0].Rows[0]["HTMLFooter"].ToString();
                    //string strsubject = dsEmailTemplate.Tables[0].Rows[0]["HTMLSubject"].ToString();
                    string strHeader = objHTMLTemplate.Header;
                    string strBody = objHTMLTemplate.Body;
                    string strFooter = objHTMLTemplate.Footer;
                    string strsubject = objHTMLTemplate.Subject;

                    strsubject = strsubject.Replace("#ID#", strTaskId);
                    strsubject = strsubject.Replace("#TaskTitleID#", strTaskTitle);
                    strsubject = strsubject.Replace("#TaskTitle#", strTaskTitle);

                    strBody = strBody.Replace("#ID#", strTaskId);
                    strBody = strBody.Replace("#TaskTitleID#", strTaskTitle);
                    strBody = strBody.Replace("#TaskTitle#", strTaskTitle);
                    strBody = strBody.Replace("#Fname#", fullname);
                    strBody = strBody.Replace("#email#", emailId);
                    strBody = strBody.Replace("#Designation(s)#", ddlDesignationForTask.SelectedItem != null ? ddlDesignationForTask.SelectedItem.Text : "");
                    strBody = strBody.Replace("#TaskLink#", string.Format(
                                                                            "{0}?TaskId={1}&hstid={2}",
                                                                            string.Concat(
                                                                                            Request.Url.Scheme,
                                                                                            Uri.SchemeDelimiter,
                                                                                            Request.Url.Host.Split('?')[0],
                                                                                            "/Sr_App/TaskGenerator.aspx"
                                                                                         ),
                                                                            strTaskId,
                                                                            strSubTaskId
                                                                        )
                                            );

                    strBody = strHeader + strBody + strFooter;

                    string strHTMLTemplateName = "Task Generator Auto Email";
                    DataSet dsEmailTemplate = AdminBLL.Instance.GetEmailTemplate(strHTMLTemplateName, 108);
                    List<Attachment> lstAttachments = new List<Attachment>();
                    // your remote SMTP server IP.
                    for (int i = 0; i < dsEmailTemplate.Tables[1].Rows.Count; i++)
                    {
                        string sourceDir = Server.MapPath(dsEmailTemplate.Tables[1].Rows[i]["DocumentPath"].ToString());
                        if (File.Exists(sourceDir))
                        {
                            Attachment attachment = new Attachment(sourceDir);
                            attachment.Name = Path.GetFileName(sourceDir);
                            lstAttachments.Add(attachment);
                        }
                    }

                    CommonFunction.SendEmail(HTMLTemplates.Task_Generator_Auto_Email.ToString(), emailId, strsubject, strBody, lstAttachments);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("{0} Exception caught.", ex);
            }
        }

        #endregion

        #region '--WebMethods--'

        [System.Web.Services.WebMethod]
        public static List<Task> GetTasksForDesignation(string designationID)
        {
            List<Task> taskList = new List<Task>();
            DataSet dsTechTaskForDesignation;

            if (!string.IsNullOrEmpty(designationID) && !designationID.ToLower().Contains("all"))
            {
                int iDesignationID = Convert.ToInt32(designationID);
                dsTechTaskForDesignation = TaskGeneratorBLL.Instance.GetAllActiveTechTaskForDesignationID(iDesignationID);
                if (dsTechTaskForDesignation != null & dsTechTaskForDesignation.Tables.Count > 0)
                {
                    //taskJSON = JsonConvert.SerializeObject(dsTechTaskForDesignation.Tables[0]);
                    if (dsTechTaskForDesignation.Tables[0].Rows.Count > 0)
                    {
                        for (int iCurrentRow = 0; iCurrentRow < dsTechTaskForDesignation.Tables[0].Rows.Count; iCurrentRow++)
                        {
                            taskList.Add(new Task
                            {
                                TaskId = Convert.ToInt32(dsTechTaskForDesignation.Tables[0].Rows[iCurrentRow]["TaskId"]),
                                Title = dsTechTaskForDesignation.Tables[0].Rows[iCurrentRow]["Title"].ToString()
                            });
                        }
                    }
                }
            }
            return taskList;
        }

        #endregion

        #endregion
    }
}