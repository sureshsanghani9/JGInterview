<%@ Page Title="" Language="C#" MasterPageFile="~/Sr_App/SR_app.Master" AutoEventWireup="true" ValidateRequest="false"
    CodeBehind="EditUser.aspx.cs" Inherits="JG_Prospect.EditUser" MaintainScrollPositionOnPostback="true" Async="true" %>

<%@ Register Assembly="System.Web.DataVisualization, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" Namespace="System.Web.UI.DataVisualization.Charting" TagPrefix="asp" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>
<%@ Register Assembly="CKEditor.NET" Namespace="CKEditor.NET" TagPrefix="CKEditor" %>
<%@ Register Src="~/UserControl/ucStatusChangePopup.ascx" TagPrefix="ucStatusChange" TagName="PoPup" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">

    <script src="../js/Custom/JgPopUp.js" type="text/javascript"></script>
    <link type="text/css" href="../css/flags24.css" rel="Stylesheet" />
    <style type="text/css">
        /*Grid add Container START*/
        .GrdContainer {
            width: 100%;
            border: 1px solid #d3d3d3;
        }

            .GrdContainer div {
                width: 100%;
            }

            .GrdContainer .GrdHeader {
                background-color: #d3d3d3;
                padding: 2px;
                cursor: pointer;
                font-weight: bold;
            }

            .GrdContainer .GrdContent {
                display: none;
                padding: 5px;
                height: 160px;
            }

        .GrdContent ul li span {
            width: 100% !important;
        }

        .GrdContent ul li {
            width: 80%;
            padding-top: 10px;
        }

            .GrdContent ul li span label {
                width: 75%;
                float: left;
                padding-top: 0px;
            }

            .GrdContent ul li span input {
                width: 20% !important;
                float: left;
            }

            .GrdContent ul li select, .GrdContent ul li input {
                width: 85% !important;
            }

        .GrdBtnAdd {
            margin-top: 12px;
            height: 30px;
            background: url(img/main-header-bg.png) repeat-x;
            color: #fff;
        }

        /*Grid add Container END */
        .PrimaryPhone {
            cursor: pointer;
        }

        .GrdPrimaryEmail {
            text-decoration: underline;
            cursor: pointer;
            color: blue;
            line-height: 20px;
            width: 150px;
            overflow: hidden;
            white-space: nowrap;
            text-overflow: ellipsis;
        }
        /*.GrdPrimaryEmail:hover {
                overflow: visible;
                white-space: normal;
                width: auto;
                position: absolute;
                background-color: #FFF;
            }*/
        .ddChild li {
            text-align: left;
            margin: 0 !important;
            width: auto !important;
            border-bottom: none !important;
        }

        .grd-lblPrimaryPhone img {
            float: left;
        }

        .grd-lblPrimaryPhone {
            width: 135px !important;
            padding-top: 8px;
        }

        .user-zip {
            padding-left: 50px;
            margin-left: 70px;
        }

        .SearchLoad {
            position: absolute;
            display: block;
            margin-top: 116px;
            margin-left: 153px;
        }

        .wordBreak {
            word-wrap: break-word;
        }

        .black_overlay {
            display: none;
            position: fixed;
            top: 0%;
            left: 0%;
            width: 100%;
            height: 100%;
            background-color: black;
            z-index: 1001;
            -moz-opacity: 0.8;
            opacity: .80;
            filter: alpha(opacity=80);
            overflow-y: hidden;
        }
        /*#327FB5*/
        .white_content {
            display: none;
            position: absolute;
            top: 10%;
            left: 20%;
            width: 60%;
            min-height: 10%;
            padding: 0 16px 16px 16px;
            border: 10px solid #000000;
            background-color: white;
            z-index: 1002;
            overflow: auto;
        }

        .close {
            position: absolute;
            top: 35px;
            right: 30px;
            transition: all 200ms;
            font-size: 30px;
            font-weight: bold;
            text-decoration: none;
            color: #333;
        }

        .HeaderFreez {
            position: absolute;
            /*top: expression(this.offsetParent.scrollTop);*/
            z-index: 10;
            margin-top: -53px;
        }

        .grdUserMain {
            /*margin-top: 50px;*/
        }

            .grdUserMain tr td {
                padding: 10px 8px 12px 4px !important;
            }

        .txtSearch {
            width: 135px;
            padding: 5px;
            border-radius: 5px 0 0 5px;
            color: #666;
            font-size: 14px;
        }

        .btnSearc {
            width: 100px;
            border-radius: 0 5px 5px 0;
            line-height: 28px;
            background: #A33E3F;
            color: #fff;
            cursor: pointer;
        }

        .btnBookmark {
            border: 0 !important;
            padding: 2px !important;
            float: right;
            position: relative;
        }
    </style>
    <script type="text/javascript">


        function ConfirmDelete() {
            var Ok = confirm('Are you sure you want to Delete this User?');
            if (Ok)
                return true;
            else
                return false;
        }

        function ClosePopup() {
            document.getElementById('light').style.display = 'none';
            document.getElementById('fade').style.display = 'none';
        }

        function overlay() {
            document.getElementById('light').style.display = 'block';
            document.getElementById('fade').style.display = 'block';
        }


        function ClosePopupInterviewDate() {
            document.getElementById('interviewDatelite').style.display = 'none';
            document.getElementById('interviewDatefade').style.display = 'none';
        }

        function overlayInterviewDate() {

            document.getElementById('interviewDatelite').style.display = 'block';
            document.getElementById('interviewDatefade').style.display = 'block';
            //$('#interviewDatelite').focus();
            $("html, body").animate({ scrollTop: 0 }, "slow");
        }

        function ClosePopupOfferMade() {
            document.getElementById('DivOfferMade').style.display = 'none';
            document.getElementById('DivOfferMadefade').style.display = 'none';
        }

        function OverlayPopupOfferMade() {
            document.getElementById('DivOfferMade').style.display = 'block';
            document.getElementById('DivOfferMadefade').style.display = 'block';
            $("html, body").animate({ scrollTop: 0 }, "slow");
        }

        function ClosePopupUploadBulk() {
            document.getElementById('lightUploadBulk').style.display = 'none';
            document.getElementById('fadeUploadBulk').style.display = 'none';
        }

        function OverlayPopupUploadBulk() {
            document.getElementById('lightUploadBulk').style.display = 'block';
            document.getElementById('fadeUploadBulk').style.display = 'block';
            $("html, body").animate({ scrollTop: 0 }, "slow");
        }

        function CloseAddUserPopUp() {
            document.getElementById('lightUploadBulk').style.display = 'none';
            document.getElementById('fadeUploadBulk').style.display = 'none';
        }

        //var validFilesTypes = ["xls", "xlsx", "csv"];
        var validFilesTypes = ["xlsx", "csv"];
        function ValidateFile() {
            var file = document.getElementById("<%=BulkProspectUploader.ClientID%>");
            var label = document.getElementById("<%=Label1.ClientID%>");
            var path = file.value;
            var ext = path.substring(path.lastIndexOf(".") + 1, path.length).toLowerCase();
            var isValidFile = false;
            for (var i = 0; i < validFilesTypes.length; i++) {
                if (ext == validFilesTypes[i]) {
                    isValidFile = true;
                    break;
                }
            }
            if (!isValidFile) {
                alert('Select file of type csv or xlsx ');
                //label.style.color = "red";
                //label.innerHTML = "Invalid File. Please upload a File with" +

                // " extension:\n\n" + validFilesTypes.join(", ");

            }
            return isValidFile;
        }

    </script>
    <script>
        function pageLoad() {


            $(document).ready(function () {


                $(".GrdHeader").click(function () {

                    $header = $(this);
                    //getting the next element
                    $content = $header.next();
                    //open up the content needed - toggle the slide- if visible, slide up, if not slidedown.
                    $content.slideToggle(500, function () {
                        //execute this after slideToggle is done
                        //change text of header based on visibility of content div
                        $header.text(function () {
                            //change text based on condition
                            return $content.is(":visible") ? "Click To Collapse" : "Click To Add Phone /Email";
                        });
                    });
                });


                $('.PrimaryPhone').click(function () {
                    showCustomPopUp("\\CommingSoon.aspx", "Primary Phone");
                });
                $('.GrdPrimaryEmail').click(function () {
                    //showCustomPopUp("\\CommingSoon.aspx", "Primary Email");
                });
            });

        }
    </script>
    <style type="text/css">
        .modalBackground {
            background-color: Black;
            filter: alpha(opacity=90);
            opacity: 0.8;
        }

        .modalPopup {
            background-color: #FFFFFF;
            border-width: 3px;
            border-style: solid;
            border-color: black;
            padding-top: 10px;
            padding-left: 2px;
            width: 129px;
            height: 173px;
        }

        table.select_period_table {
        }

            table.select_period_table tr td {
                width: 50% !important;
            }

                table.select_period_table tr td label {
                    display: block !important;
                    width: 100% !important;
                }

                table.select_period_table tr td input {
                    width: 100% !important;
                    box-sizing: border-box !important;
                }

        table.tblshowhrdata {
            width: 100%;
            border: 1px solid #ddd;
            background: #fff;
            border-collapse: collapse;
        }

        .tblPieChart td.head {
            color: white;
            font-weight: bold;
            text-align: center;
            height: 15px;
            background: #A33E3F url(../img/line.png) bottom repeat-x;
            padding: 10px 0px;
            width: 32%;
            line-height: 15px;
            min-height: 5px;
            vertical-align: top;
        }

        .scrollCls {
            height: 300px !important;
            overflow-y: scroll;
        }
        /*.scrollCls table tbody {
            display: block;
            height: 300px;
            overflow-y: scroll;
        }*/
    </style>
    <link href="../Styles/dd.css" rel="stylesheet" />
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <link href="../css/dropzone/css/basic.css" rel="stylesheet" />
    <link href="../css/dropzone/css/dropzone.css" rel="stylesheet" />
    <script type="text/javascript" src="../js/dropzone.js"></script>
    <script src="../ckeditor/ckeditor.js"></script>
    <div class="right_panel">
        <!-- appointment tabs section start -->
        <ul class="appointment_tab">
            <li><a href="HRReports.aspx">HR Reports</a></li>
            <li><a href="InstallCreateUser.aspx">Create Install User</a></li>
            <li><a href="EditInstallUser.aspx">Edit Install User</a></li>
            <li><a href="CreateSalesUser.aspx">Create Sales-Admin_IT User</a></li>
            <li><a href="EditUser.aspx">Edit Sales-Admin_IT User</a></li>
        </ul>
        <h1>Edit User</h1>
        <div class="form_panel">
            <asp:UpdatePanel ID="upSalesUserStatictics" runat="server">
                <ContentTemplate>
                    <span>
                        <asp:Label ID="lblmsg" runat="server" Visible="false"></asp:Label>
                    </span>
                    <table style="width: 100%; background-color: #fff;" class="tblPieChart">
                        <tr>
                            <td style="width: 50%; padding: 0px;">
                                <asp:Chart ID="Chart1" runat="server" Height="320px" Width="415px">
                                    <Titles>
                                        <asp:Title ShadowOffset="3" Name="Items" />
                                    </Titles>
                                    <Legends>
                                        <asp:Legend Alignment="Center" Docking="Bottom" IsTextAutoFit="False" Name="Default" LegendStyle="Table" />
                                    </Legends>
                                    <Series>
                                        <asp:Series Name="Default" />
                                    </Series>
                                    <ChartAreas>
                                        <asp:ChartArea Name="ChartArea1" BorderWidth="0" />
                                    </ChartAreas>
                                </asp:Chart>
                            </td>
                            <td style="width: 50%; padding: 0px;">
                                <div class="scrollCls">
                                    <table style="height: inherit;">
                                        <tr>
                                            <td class="head">Added By</td>
                                            <td class="head">Designation</td>
                                            <td class="head">Source</td>
                                        </tr>
                                        <tr>
                                            <td style="padding: 0px;">
                                                <table>
                                                    <asp:ListView ID="listAddedBy" runat="server">
                                                        <ItemTemplate>
                                                            <tr>
                                                                <td><span><%#(Eval("AddedBy") == null || Eval("AddedBy") == "" )? "No Name" : Eval("AddedBy")%></span></td>
                                                                <td><span><%#Eval("Count")%></span></td>
                                                            </tr>
                                                        </ItemTemplate>
                                                    </asp:ListView>
                                                </table>
                                            </td>
                                            <td style="padding: 0px;">
                                                <table>
                                                    <asp:ListView ID="listDesignation" runat="server">
                                                        <ItemTemplate>
                                                            <tr>
                                                                <td><span><%#(Eval("Designation") == null || Eval("Designation") == "" )? "No Designation" : Eval("Designation")%></span></td>
                                                                <td><span><%#Eval("Count")%></span></td>
                                                            </tr>
                                                        </ItemTemplate>
                                                    </asp:ListView>
                                                </table>
                                            </td>
                                            <td style="padding: 0px;">
                                                <table>
                                                    <asp:ListView ID="listSource" runat="server">
                                                        <ItemTemplate>
                                                            <tr>
                                                                <td><span><%#(Eval("Source") == null || Eval("Source") == "" )? "No Name" : Eval("Source")%></span></td>
                                                                <td><span><%#Eval("Count")%></span></td>
                                                            </tr>
                                                        </ItemTemplate>
                                                    </asp:ListView>
                                                </table>
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                            </td>
                        </tr>
                    </table>
                    <br />
                    <br />
                    <div class="showhrdata">
                        <table class="tblshowhrdata">
                            <tr>
                                <td>
                                    <asp:Label ID="lbljoboffer" runat="server">New "Job Offers" Submissions</asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lbljoboffercount" runat="server" Text="0"></asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lblInterviewDate" runat="server">New "Interview Date"</asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lblInterviewDateCount" runat="server" Text="0"></asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lblActive" runat="server">New active</asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lblActiveCount" runat="server" Text="0"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:Label ID="lblPhoneVideoScreened" runat="server">New "Phone/Video Screened"</asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lblPhoneVideoScreenedCount" runat="server" Text="0"></asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lblRejected" runat="server">New "Rejected"</asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lblRejectedCount" runat="server" Text="0"></asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lblDeactivated" runat="server">New "Deactivated"</asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lblDeactivatedCount" runat="server" Text="0"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:Label ID="lblNewApplicants" runat="server">New "Applicants"</asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lblNewApplicantsCount" runat="server" Text="0"></asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lblInstallProspect" runat="server">New "Prospect Referrals"</asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lblInstallProspectCount" runat="server" Text="0"></asp:Label>
                                </td>

                                <td>
                                    <asp:Label ID="lblAppInterview" runat="server">Applicant/interview ratio</asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lblAppInterviewRatio" runat="server" Text="0"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:Label ID="lblInterviewActive" runat="server">Interview/Active ratio</asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lblInterviewActiveRatio" runat="server" Text="0"></asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lblJobOfferActive" runat="server">Offer Made/Active ratio</asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lblJobOfferActiveRatio" runat="server" Text="0"></asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lblActiveDeactive" runat="server">Active/Deactive Ratio</asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lblActiveDeactiveRatio" runat="server" Text="0"></asp:Label>
                                </td>
                                <%--<td>
                                    <asp:Label ID="lblInactive" runat="server">New Inactive</asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lblInactiveCount" runat="server" Text="0"></asp:Label>
                                </td><td>
                                    <asp:Label ID="lblAppHire" runat="server">Applicant/new hire ratio</asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lblAppHireRatio" runat="server" Text="0"></asp:Label>
                                </td><td>
                                    <asp:Label ID="lblJobOfferHire" runat="server">Job Offer/new hire ratio	</asp:Label>
                                </td>
                                <td>
                                    <asp:Label ID="lblJobOfferHireRatio" runat="server" Text="0"></asp:Label>
                                </td>--%>
                            </tr>
                        </table>
                    </div>
                </ContentTemplate>
            </asp:UpdatePanel>
            <br />
            <br />
            <asp:UpdatePanel ID="upFilter" runat="server">
                <ContentTemplate>
                    <table style="width: 100%;">
                        <tr style="background-color: #A33E3F; color: white; font-weight: bold; text-align: center; width: 100%;">
                            <td>
                                <asp:Label ID="lblUserStatus" Text="User Status" runat="server" /><span style="color: red">*</span></td>
                            <td>
                                <asp:Label ID="lblDesignation" Text="Designation" runat="server" /></td>
                            <td>
                                <asp:Label ID="lblAddedBy" Text="Added By" runat="server" /></td>
                            <td>
                                <asp:Label ID="lblSourceH" Text="Source" runat="server" /></td>
                            <td colspan="2">
                                <asp:Label ID="Label2" Text="Select Period" runat="server" /></td>
                        </tr>
                        <tr style="text-align: center; width: 100%">
                            <td style="text-align: center;">
                                <asp:DropDownList ID="ddlUserStatus" runat="server" Width="140px" AutoPostBack="true"
                                    OnSelectedIndexChanged="ddlFilter_SelectedIndexChanged" OnPreRender="ddlUserStatus_PreRender" />
                            </td>
                            <td>
                                <asp:DropDownList ID="ddlDesignation" runat="server" Width="140px"
                                    OnSelectedIndexChanged="ddlFilter_SelectedIndexChanged" AutoPostBack="true" />
                            </td>
                            <td>
                                <asp:DropDownList ID="drpUser" runat="server" Width="140px" OnSelectedIndexChanged="ddlFilter_SelectedIndexChanged" AutoPostBack="true"></asp:DropDownList></td>
                            <td>
                                <asp:DropDownList ID="ddlSource" runat="server" Width="140px" OnSelectedIndexChanged="ddlFilter_SelectedIndexChanged" AutoPostBack="true"></asp:DropDownList></td>
                            <td>
                                <asp:CheckBox ID="chkAllDates" runat="server" Checked="true" Text="All" OnCheckedChanged="chkAllDates_CheckedChanged" AutoPostBack="true" /></td>
                            <td>
                                <asp:Label ID="Label3" Text="From :*" runat="server" />
                                <asp:TextBox ID="txtfrmdate" runat="server" TabIndex="2" CssClass="date"
                                    onkeypress="return false" MaxLength="10" AutoPostBack="true"
                                    Style="width: 80px;" OnTextChanged="txtfrmdate_TextChanged" Enabled="false"></asp:TextBox>
                                <cc1:CalendarExtender ID="calExtendFromDate" runat="server" TargetControlID="txtfrmdate">
                                </cc1:CalendarExtender>
                                <asp:Label ID="Label4" Text="To :*" runat="server" />
                                <asp:TextBox ID="txtTodate" CssClass="date" onkeypress="return false"
                                    MaxLength="10" runat="server" TabIndex="3" AutoPostBack="true"
                                    Style="width: 80px;" OnTextChanged="txtTodate_TextChanged" Enabled="false"></asp:TextBox>
                                <cc1:CalendarExtender ID="CalendarExtender2" runat="server" TargetControlID="txtTodate">
                                </cc1:CalendarExtender>
                                <br />
                                <asp:RequiredFieldValidator ID="requirefrmdate" ControlToValidate="txtfrmdate"
                                    runat="server" ErrorMessage=" Select From date" ForeColor="Red" ValidationGroup="show">
                                </asp:RequiredFieldValidator><asp:RequiredFieldValidator ID="Requiretodate" ControlToValidate="txtTodate"
                                    runat="server" ErrorMessage=" Select To date" ForeColor="Red" ValidationGroup="show">
                                </asp:RequiredFieldValidator>
                                <br />
                            </td>
                        </tr>
                    </table>
                </ContentTemplate>
            </asp:UpdatePanel>
            <div style="width: auto; border: 1px solid #ccc; padding: 3px;">
                <asp:UpdatePanel ID="upUsers" runat="server" UpdateMode="Conditional">
                    <ContentTemplate>

                        <div style="float: left; padding-top: 10px; /*margin-bottom: -40px; */">

                            <asp:TextBox ID="txtSearch" runat="server" CssClass="textbox" placeholder="search users" MaxLength="15" />
                            <asp:Button ID="btnSearchGridData" runat="server" Text="Search" Style="display: none;" class="btnSearc" OnClick="btnSearchGridData_Click" />

                            Number of Records: 
                            <asp:DropDownList ID="ddlPageSize_grdUsers" runat="server" AutoPostBack="true"
                                OnSelectedIndexChanged="ddlPageSize_grdUsers_SelectedIndexChanged">
                                <asp:ListItem Text="10" Value="10" />
                                <asp:ListItem Selected="True" Text="20" Value="20" />
                                <asp:ListItem Text="30" Value="30" />
                                <asp:ListItem Text="40" Value="40" />
                                <asp:ListItem Text="50" Value="50" />
                            </asp:DropDownList>
                        </div>


                        <asp:GridView ID="grdUsers" OnPreRender="grdUsers_PreRender" runat="server" CssClass="scroll" Width="100%" EmptyDataText="No Data"
                            AutoGenerateColumns="False" DataKeyNames="Id,DesignationID" AllowSorting="true" AllowPaging="true" AllowCustomPaging="true" PageSize="20"
                            OnRowDataBound="grdUsers_RowDataBound" OnRowCommand="grdUsers_RowCommand" OnSorting="grdUsers_Sorting"
                            OnPageIndexChanging="grdUsers_PageIndexChanging">
                            <PagerSettings Mode="NumericFirstLast" NextPageText="Next" PreviousPageText="Previous" Position="TopAndBottom" />
                            <PagerStyle HorizontalAlign="Right" CssClass="pagination-ys" />
                            <Columns>

                                <asp:TemplateField HeaderText="Action" HeaderStyle-HorizontalAlign="Center" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="5%" ItemStyle-Width="5%" SortExpression="BookmarkedInstallUserId">
                                    <ItemTemplate>
                                        <asp:CheckBox ID="chkSelected" runat="server" />
                                        <br />
                                        <asp:LinkButton ID="lbltest" Text="Edit" CommandName="EditSalesUser" runat="server"
                                            CommandArgument='<%#Eval("Id")%>'></asp:LinkButton>
                                        <br />
                                        <asp:LinkButton ID="lnkDeactivate" Text="Deactivate" CommandName="DeactivateSalesUser" runat="server" OnClientClick="return confirm('Are you sure you want to deactivate this user?')"
                                            CommandArgument='<%#Eval("Id")%>'></asp:LinkButton>
                                        <br />
                                        <asp:LinkButton ID="lnkDelete" Text="Delete" CommandName="DeleteSalesUser" runat="server" OnClientClick="return confirm('Are you sure you want to delete this user?')"
                                            CommandArgument='<%#Eval("Id")%>'></asp:LinkButton>
                                        <br />
                                        <asp:ImageButton ID="imgBookmark" runat="server" CssClass="btnBookmark" BorderWidth="0"/>
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField ShowHeader="True" HeaderText="Id# <br /> Designation" HeaderStyle-HorizontalAlign="Center" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="10%" ItemStyle-Width="10%" ControlStyle-ForeColor="Black"
                                    Visible="true" SortExpression="Designation">
                                    <EditItemTemplate>
                                        <asp:TextBox ID="txtid" runat="server" MaxLength="30" Text='<%#Eval("Id")%>'></asp:TextBox>
                                    </EditItemTemplate>
                                    <ItemTemplate>
                                        <asp:Label ID="lblid" Visible="false" runat="server" Text='<%#Eval("Id")%>'></asp:Label>
                                        <asp:LinkButton ID="lnkID" Text='<%#Eval("UserInstallId")%>' CommandName="EditSalesUser" runat="server"
                                            CommandArgument='<%#Eval("Id")%>'></asp:LinkButton>
                                        <br />
                                        <asp:Label ID="lblDesignation" runat="server" Text='<%#Eval("Designation")%>'></asp:Label>
                                    </ItemTemplate>
                                    <ControlStyle ForeColor="Black" />
                                    <ControlStyle ForeColor="Black" />
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:TemplateField>

                                <asp:TemplateField ShowHeader="True" HeaderText="Install Id" Visible="false" SortExpression="Id" ControlStyle-ForeColor="Black"
                                    ItemStyle-HorizontalAlign="Center">
                                    <EditItemTemplate>
                                        <asp:TextBox ID="txtInstallid" runat="server" MaxLength="30" Text='<%#Eval("InstallId")%>'></asp:TextBox>
                                    </EditItemTemplate>
                                    <ItemTemplate>
                                        <asp:Label ID="lblInstallid" runat="server" Text='<%#Eval("InstallId")%>'></asp:Label>
                                    </ItemTemplate>
                                    <ControlStyle ForeColor="Black" />
                                    <ControlStyle ForeColor="Black" />
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Picture" Visible="false" SortExpression="picture">
                                    <ItemTemplate>
                                        <asp:LinkButton ID="lbtnPicture" Text="Picture" CommandName="ShowPicture" runat="server"
                                            CommandArgument='<%#Eval("picture")%>'></asp:LinkButton>
                                    </ItemTemplate>
                                </asp:TemplateField>

                                <asp:TemplateField ShowHeader="True" HeaderText="First Name<br />Last Name" HeaderStyle-HorizontalAlign="Center" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="15%" ItemStyle-Width="15%" SortExpression="FristName" ControlStyle-ForeColor="Black">
                                    <EditItemTemplate>
                                        <asp:TextBox ID="txtFirstName" runat="server" MaxLength="30" Text='<%#Eval("FristName")%>'></asp:TextBox>
                                    </EditItemTemplate>
                                    <ItemTemplate>
                                        <asp:Label ID="lblFirstName" runat="server" Text='<%#Eval("FristName")%>'></asp:Label>
                                        <br />
                                        <asp:Label ID="lblLastName" runat="server" Text='<%# Eval("Lastname") %>'></asp:Label>
                                    </ItemTemplate>
                                    <ControlStyle ForeColor="Black" />
                                    <ControlStyle ForeColor="Black" />
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="Last name" Visible="false" SortExpression="Lastname" ItemStyle-HorizontalAlign="Center">
                                    <EditItemTemplate>
                                        <asp:TextBox ID="txtlastname" runat="server" Text='<%# Bind("Lastname") %>'></asp:TextBox>
                                    </EditItemTemplate>
                                    <ItemTemplate>
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Designation" Visible="false" SortExpression="Designation" ItemStyle-HorizontalAlign="Center">
                                    <EditItemTemplate>
                                        <asp:TextBox ID="txtDesignation" runat="server" Text='<%#Eval("Designation")%>'></asp:TextBox>
                                    </EditItemTemplate>
                                    <ItemTemplate>
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="Status" ItemStyle-HorizontalAlign="Center" HeaderStyle-HorizontalAlign="Center" HeaderStyle-Width="20%" ItemStyle-Width="20%" SortExpression="Status">
                                    <ItemTemplate>
                                        <asp:HiddenField ID="lblStatus" runat="server" Value='<%#Eval("Status")%>'></asp:HiddenField>
                                        <asp:HiddenField ID="lblOrderStatus" runat="server" Value='<%#(Eval("OrderStatus") == null || Eval("OrderStatus") == "") ? -99: Eval("OrderStatus")%>'></asp:HiddenField>
                                        <asp:DropDownList ID="ddlStatus" CssClass="grd-status" Style="width: 95%;" AutoPostBack="true" OnSelectedIndexChanged="grdUsers_ddlStatus_SelectedIndexChanged" runat="server" OnPreRender="ddlUserStatus_PreRender">
                                            <%--<asp:ListItem Text="Referral applicant" Value="ReferralApplicant"></asp:ListItem>
                                            <asp:ListItem Text="Applicant" Value="Applicant"></asp:ListItem>
                                            <asp:ListItem Text="Phone/Video Screened" Value="PhoneScreened"></asp:ListItem>
                                            <asp:ListItem Text="Rejected" Value="Rejected"></asp:ListItem>
                                            <asp:ListItem Text="Interview Date" Value="InterviewDate"></asp:ListItem>
                                            <asp:ListItem Text="Offer Made" Value="OfferMade"></asp:ListItem>
                                            <asp:ListItem Text="Active" Value="Active"></asp:ListItem>
                                            <asp:ListItem Text="Deactive" Value="Deactive"></asp:ListItem>
                                            <asp:ListItem Text="Install Prospect" Value="Install Prospect"></asp:ListItem>--%>
                                        </asp:DropDownList><br />
                                        <asp:Label ID="lblRejectDetail" runat="server" Text='<%#Eval("RejectDetail") %>'></asp:Label>
                                        <br />
                                        <span><%#string.IsNullOrEmpty(Eval("InterviewDetail").ToString())?"":Eval("InterviewDetail").ToString().Split(' ')[0]%></span>&nbsp<span style="color: red"><%#string.IsNullOrEmpty(Eval("InterviewDetail").ToString())?"":Eval("InterviewDetail").ToString().Remove(0, Eval("InterviewDetail").ToString().IndexOf(' ') + 1)%></span>&nbsp<span><%#string.IsNullOrEmpty(Eval("InterviewDetail").ToString())?"":"(EST)"%></span><asp:Label ID="lblInterviewDetail" runat="server" Visible="false" Text='<%#Eval("InterviewDetail") %>'></asp:Label>
                                        <asp:HyperLink ID="hypTechTask" runat="server" Visible="false" />
                                    </ItemTemplate>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Source<br/>Added By<br/>Added On" HeaderStyle-Width="15%" ItemStyle-Width="15%" SortExpression="Source" ItemStyle-HorizontalAlign="Center" HeaderStyle-HorizontalAlign="Center">
                                    <ItemTemplate>
                                        <asp:Label ID="lblSource" runat="server" Text='<%#Eval("Source")%>'></asp:Label>
                                        <br />
                                        <span><%#Eval("AddedBy")%></span>
                                        <a href='<%#(string.IsNullOrEmpty(Eval("AddedById").ToString()))?"#":"ViewSalesUser.aspx?id="+ Eval("AddedById")%>'><%#(string.IsNullOrEmpty(Eval("AddedByUserInstallId").ToString()))?"":"- "+ Eval("AddedByUserInstallId")%></a>
                                        <%--<asp:LinkButton ID="lnkAddedByUserInstallId" Text='<%#(string.IsNullOrEmpty(Eval("AddedByUserInstallId").ToString()))?"":"-"+ Eval("AddedByUserInstallId")%>' CommandName="EditAddedByUserInstall" runat="server"
                                            CommandArgument='<%#(string.IsNullOrEmpty(Eval("AddedById").ToString()))?"":Eval("AddedById")%>' Enabled='<%#(string.IsNullOrEmpty(Eval("AddedByUserInstallId").ToString()))?false:true%>'></asp:LinkButton>--%>
                                        <br />
                                        <span><%#String.Format("{0:M/d/yyyy}", Eval("CreatedDateTime"))%></span>&nbsp<span style="color: red"><%#String.Format("{0:hh:mm:ss tt}", Eval("CreatedDateTime"))%></span>&nbsp<span>(EST)</span>
                                    </ItemTemplate>

                                </asp:TemplateField>

                                <asp:TemplateField HeaderText="Added On" Visible="false" SortExpression="CreatedDateTime" HeaderStyle-HorizontalAlign="Center" ItemStyle-HorizontalAlign="Center">
                                    <ItemTemplate>
                                    </ItemTemplate>

                                </asp:TemplateField>


                                <asp:TemplateField HeaderText="Email<br/>Phone Type - Phone" HeaderStyle-Width="20%" ItemStyle-Width="20%" ItemStyle-HorizontalAlign="Center" SortExpression="Phone">
                                    <ItemTemplate>
                                        <%-- ControlStyle-CssClass="wordBreak" <asp:Label ID="lblPhone" runat="server" Text='<%# Bind("Phone") %>'></asp:Label>--%>
                                        <%--onclick="<%# "javascript:grdUsers_Email_OnClick(this,'" + Eval("Email") + "');"%>"--%>
                                        <div class="GrdPrimaryEmail">
                                            <asp:LinkButton ID="lbtnEmail" runat="server" Text='<%# Eval("Email") %>' ToolTip='<%# Eval("Email") %>'
                                                CommandName="send-email" CommandArgument='<%# Container.DataItemIndex %>' />
                                        </div>
                                        <asp:Label ID="lblPrimaryPhone" CssClass="grd-lblPrimaryPhone" runat="server" Text='<%# Eval("PrimaryPhone") %>'></asp:Label>
                                        <div class="GrdContainer" style="width: 90%">
                                            <div class="GrdHeader">
                                                <span>Click To Add Phone /Email</span>
                                            </div>
                                            <div class="GrdContent">
                                                <ul style="padding-left: 0px;">
                                                    <li>
                                                        <asp:CheckBox ID="chkIsPrimaryPhone" Text=" Is Primary contact" runat="server"></asp:CheckBox></li>
                                                    <li>
                                                        <asp:DropDownList ID="ddlContactType" runat="server">
                                                            <asp:ListItem Text="Home Phone"></asp:ListItem>
                                                            <asp:ListItem Text="Office Phone"></asp:ListItem>
                                                            <asp:ListItem Text="Alt Phone"></asp:ListItem>
                                                            <asp:ListItem Text="Email"></asp:ListItem>
                                                        </asp:DropDownList></li>
                                                    <li>
                                                        <asp:TextBox ID="txtNewContact" runat="server"></asp:TextBox></li>
                                                    <li>
                                                        <asp:Button ID="btnAddPhone" CssClass="GrdBtnAdd" runat="server" Text="Add" CommandName="AddNewContact" CommandArgument='<%# Eval("Id") %>'></asp:Button></li>
                                                </ul>
                                            </div>
                                        </div>

                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:TemplateField HeaderText="Country-Zip<br/>Type-Apptitude Test %<br/>Resume Attachment" HeaderStyle-Width="15%" ItemStyle-Width="15%" ItemStyle-HorizontalAlign="Center" SortExpression="Zip" ControlStyle-CssClass="wordBreak">
                                    <ItemTemplate>
                                        <div style='<%# string.IsNullOrEmpty(Eval("CountryCode").ToString()) == true ? "": "background-image:url(img/flags24.png);background-repeat:no-repeat;float:left;height:22px;width:24px;margin-top:-5px;" %>' class='<%#Eval("CountryCode").ToString().ToLower()%>'>
                                        </div>
                                        <%--<span><%# Eval("Zip") %></span>--%>
                                        <asp:Label ID="lblZip" runat="server" Text='<%# " - "+ Eval("Zip") %>'></asp:Label>

                                        <br />
                                        <br />
                                        <span><%# (Eval("EmpType").ToString() =="0")?"Not Selected -":Eval("EmpType") +" -" %></span>
                                        <span><%#(string.IsNullOrEmpty(Eval("Aggregate").ToString()))?"N/A":string.Format("{0:#,##}",Eval("Aggregate"))+ "%" %></span>

                                        <br />
                                        <a href='<%# Eval("Resumepath") %>' id="aReasumePath" runat="server" target="_blank"><%# System.IO.Path.GetFileName(Eval("Resumepath").ToString()) %></a>
                                        <%--<span><%# Eval("EmpType") %></span> <span> - <span><%#(string.IsNullOrEmpty(Eval("Aggregate").ToString()))?"N/A":string.Format("{0:#,##}",Eval("Aggregate"))+ "%" %></span>--%>
                                    </ItemTemplate>
                                </asp:TemplateField>

                            </Columns>
                        </asp:GridView>
                    </ContentTemplate>
                    <Triggers>
                        <asp:AsyncPostBackTrigger ControlID="lbtnDeactivateSelected" EventName="Click" />
                        <asp:AsyncPostBackTrigger ControlID="lbtnDeleteSelected" EventName="Click" />
                        <asp:AsyncPostBackTrigger ControlID="lbtnChangeStatusForSelected" EventName="Click" />
                    </Triggers>
                </asp:UpdatePanel>
            </div>
            <table style="width: 100%">
                <tr style="width: 100%">
                    <td>
                        <asp:LinkButton ID="lnkDownload" Text="Download Sample Excel Format For Bulk Upload" CommandArgument='../UserFile/SalesSample.xlsx' runat="server" OnClick="DownloadFile"></asp:LinkButton>
                        <%--<br />
                        <br />
                        <asp:LinkButton ID="lnkDownloadCSV" Text="Download Sample CSV Format For Bulk Upload" CommandArgument='../UserFile/SalesSample.csv' runat="server" OnClick="DownloadFile"></asp:LinkButton>--%>
                    </td>
                    <td>
                        <div style="float: left;">
                            <div id="divBulkUploadFile" class="dropzone work-file" data-hidden="<%=hdnBulkUploadFile.ClientID%>"
                                data-accepted-files=".csv,.xlsx" data-upload-path-code="1">
                                <div class="fallback">
                                    <input name="WorkFile" type="file" />
                                    <input type="submit" value="UploadWorkFile" />
                                </div>
                            </div>
                            <div id="divBulkUploadFilePreview" class="dropzone-previews work-file-previews">
                            </div>
                        </div>
                        <div class="btn_sec" style="float: left;">
                            <asp:Button ID="btnUploadNew" runat="server" Text="Upload" OnClick="btnUploadNew_Click" CssClass="ui-button" Style="padding: 0px 10px 0px 10px!important;" />
                        </div>
                        <div class="hide">
                            <input id="hdnBulkUploadFile" runat="server" type="hidden" />
                            <asp:Button ID="btnUpload" runat="server" Text="Upload" OnClientClick="return ValidateFile()" OnClick="btnUpload_Click" />

                            <label>Upload Prospects using xlsx file:
                                <asp:FileUpload ID="BulkProspectUploader" runat="server" /></label>
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator3" ControlToValidate="BulkProspectUploader" runat="server" ErrorMessage="Select file to import data." ValidationGroup="BulkImport"></asp:RequiredFieldValidator>
                        </div>
                    </td>
                    <td align="right">
                        <asp:LinkButton ID="lbtnDeactivateSelected" runat="server" Text="Deactivate Selected"
                            OnClientClick="return confirm('Are you sure you want to deactivate selected users?')"
                            OnClick="lbtnDeactivateSelected_Click" />
                        <br />
                        <br />
                        <asp:LinkButton ID="lbtnDeleteSelected" runat="server" Text="Delete Selected"
                            OnClientClick="return confirm('Are you sure you want to delete selected users?')"
                            OnClick="lbtnDeleteSelected_Click" />
                        <br />
                        <br />
                        <asp:LinkButton ID="lbtnChangeStatusForSelected" runat="server" Text="Change Status For Selected" OnClick="lbtnChangeStatusForSelected_Click" />
                    </td>
                </tr>
            </table>
            <br />
            <br />
            <div class="btn_sec">
                <asp:Button ID="btnExport" runat="server" Text="Export" OnClick="btnExport_Click" /><br />
                <br />
                <asp:Label ID="Label1" runat="server" />
            </div>
        </div>
    </div>
    <%--Modal Popup Stars--%>
    <div id="divModalPopups">
        <%--<asp:UpdatePanel ID="updatepanel1" runat="server">
                        <ContentTemplate>--%>
        <asp:Button ID="Button1" Style="display: none;" runat="server" Text="Button" />
        <cc1:ModalPopupExtender ID="mp1" runat="server" PopupControlID="Panel1" TargetControlID="Button1"
            CancelControlID="btnClose" BackgroundCssClass="modalBackground">
        </cc1:ModalPopupExtender>
        <asp:Panel ID="Panel1" runat="server" CssClass="modalPopup" align="center" Style="display: none">
            <asp:Image ID="img_InstallerImage" runat="server" Height="150px" Width="118px" />
            <br />
            <asp:Button ID="btnClose" runat="server" Text="Close" />
        </asp:Panel>

        <asp:Panel ID="panelPopup" runat="server">
            <div id="light" class="white_content">
                <h3>Reason
                </h3>
                <a href="javascript:void(0)" onclick="document.getElementById('light').style.display='none';document.getElementById('fade').style.display='none'">Close</a>
                <table width="100%" style="border: Solid 3px #b04547; width: 100%; height: 70%"
                    cellpadding="0" cellspacing="0">
                    <tr>
                        <td align="center" colspan="2" style="height: 15px;">
                            <asp:TextBox ID="txtReason" runat="server" placeholder="Enter Reason" TextMode="MultiLine"></asp:TextBox>
                            <asp:RequiredFieldValidator ID="reqReason" runat="server" ErrorMessage="Enter reason" ControlToValidate="txtReason" ValidationGroup="Reason"></asp:RequiredFieldValidator>
                        </td>
                    </tr>
                    <tr>
                        <td align="center">
                            <asp:Button ID="btnSaveReason" runat="server" BackColor="#327FB5" ForeColor="White" Height="32px"
                                Style="height: 26px; font-weight: 700; line-height: 1em;" Text="Save" Width="100px" ValidationGroup="Reason"
                                TabIndex="119" OnClick="btnSaveReason_Click" />
                            <%--<asp:Button ID="Button2" runat="server" OnClick="" />--%>
                        </td>
                    </tr>
                </table>
            </div>
        </asp:Panel>
        <div id="fade" class="black_overlay">
        </div>
        <ucStatusChange:PoPup ID="UcStatusPopUp" runat="server"></ucStatusChange:PoPup>
        <asp:Panel ID="panel2" runat="server">
            <div id="interviewDatelite" class="white_content" style="height: auto;">
                <h3>Interview Details
                </h3>
                <%--<a href="javascript:void(0)" onclick="">Close</a>--%>
                <asp:UpdatePanel runat="server" UpdateMode="Always">
                    <ContentTemplate>
                        <table width="100%" style="border: Solid 3px #b04547; width: 100%; height: 300px;"
                            cellpadding="0" cellspacing="0">
                            <tr>
                                <td colspan="3" align="center">Name:
                                    <asp:Label ID="lblName_InterviewDetails" runat="server" /></td>
                            </tr>
                            <tr>
                                <td align="center" style="height: 15px;">Date :
                        <asp:TextBox ID="dtInterviewDate" placeholder="Select Date" runat="server" ClientIDMode="Static" onkeypress="return false" TabIndex="104" Width="127px"></asp:TextBox>
                                    <cc1:CalendarExtender ID="CalendarExtender1" TargetControlID="dtInterviewDate" Format="MM/dd/yyyy" runat="server"></cc1:CalendarExtender>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ErrorMessage="Select Date" ControlToValidate="dtInterviewDate" ValidationGroup="InterviewDate"></asp:RequiredFieldValidator>
                                </td>
                                <td align="center"></td>
                                <td>Time :
                            <asp:DropDownList ID="ddlInsteviewtime" runat="server" TabIndex="105" Width="112px"></asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td align="right">Recruiter</td>
                                <td>: </td>
                                <td align="left">
                                    <asp:DropDownList ID="ddlUsers" runat="server" />
                                    <asp:RequiredFieldValidator ID="rfvddlUsers" runat="server" ErrorMessage="Select Recruiter" ControlToValidate="ddlUsers"
                                        ValidationGroup="InterviewDate" InitialValue="0" />
                                </td>
                            </tr>
                            <tr>
                                <td align="right">Designation</td>
                                <td>: </td>
                                <td align="left">
                                    <asp:DropDownList ID="ddlDesignationForTask" runat="server" Width="140px" EnableViewState="true" AutoPostBack="true" OnSelectedIndexChanged="ddlDesignationForTask_SelectedIndexChanged"></asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td align="right">Task</td>
                                <td>: </td>
                                <td align="left">
                                    <asp:DropDownList ID="ddlTechTask" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlTechTask_SelectedIndexChanged" />
                                    <asp:RequiredFieldValidator ID="rfvTechTask" runat="server" ControlToValidate="ddlTechTask" ErrorMessage="Select Tech Task"
                                        InitialValue="0" ValidationGroup="InterviewDate" />
                                </td>
                            </tr>
                            <tr>
                                <td align="right">Sub Task</td>
                                <td>: </td>
                                <td align="left">
                                    <asp:DropDownList ID="ddlTechSubTask" runat="server" />
                                    <asp:RequiredFieldValidator ID="rfvTechSubTask" runat="server" ControlToValidate="ddlTechSubTask" ErrorMessage="Select Sub Task"
                                        InitialValue="0" ValidationGroup="InterviewDate" />
                                </td>
                            </tr>
                            <tr>
                                <td align="center" colspan="3">
                                    <asp:Button ID="btnSaveInterview" runat="server" BackColor="#327FB5" ForeColor="White" Height="32px"
                                        Style="height: 26px; font-weight: 700; line-height: 1em;" Text="OK" Width="100px" ValidationGroup="InterviewDate"
                                        TabIndex="119" OnClick="btnSaveInterview_Click" />
                                    <asp:Button ID="btnCancelInterview" runat="server" Text="Cancel" OnClick="btnCancelInterview_Click" Width="100px"
                                        Style="height: 26px; font-weight: 700; line-height: 1em;"
                                        OnClientClick="javascript:document.getElementById('interviewDatelite').style.display='none';document.getElementById('interviewDatefade').style.display='none'" />
                                </td>
                            </tr>
                        </table>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </div>
        </asp:Panel>

        <asp:Panel ID="panel3" runat="server">
            <div id="litePassword" class="white_content">
                <h3>Password
                </h3>
                <a href="javascript:void(0)" onclick="document.getElementById('litePassword').style.display='none';document.getElementById('fadePassword').style.display='none'">Close</a>
                <table width="100%" style="border: Solid 3px #b04547; width: 100%; height: 70%"
                    cellpadding="0" cellspacing="0">
                    <tr>
                        <td align="center" style="height: 54px; width: 200px;">Enter Password To Change Status
                        </td>
                        <td align="center" style="height: 54px;">
                            <asp:TextBox ID="txtPassword" runat="server" placeholder="Enter Password" TextMode="Password"></asp:TextBox>
                            <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ErrorMessage="Enter Password" ControlToValidate="txtPassword" ValidationGroup="Password"></asp:RequiredFieldValidator>
                        </td>
                    </tr>
                    <tr>
                        <td align="center" colspan="2" style="height: 54px;">
                            <asp:Button ID="btnChangeStatus" runat="server" BackColor="#327FB5" ForeColor="White" Height="32px"
                                Style="height: 26px; font-weight: 700; line-height: 1em;" Text="Save" Width="100px" ValidationGroup="Password"
                                TabIndex="119" OnClick="btnChangeStatus_Click" />
                            <%--<asp:Button ID="Button2" runat="server" OnClick="" />--%>
                        </td>
                    </tr>
                </table>
            </div>
        </asp:Panel>
        <div id="fadePassword" class="black_overlay">
        </div>

        <asp:Panel ID="panel4" runat="server">
            <div id="DivOfferMade" class="white_content" style="height: auto;">
                <h3>Offer Made Details</h3>
                <asp:UpdatePanel runat="server" UpdateMode="Always">
                    <ContentTemplate>
                        <asp:HiddenField ID="hdnFirstName" runat="server" />
                        <asp:HiddenField ID="hdnLastName" runat="server" />
                        <table width="100%" style="border: Solid 3px #b04547; width: 100%; height: 300px;"
                            cellpadding="0" cellspacing="0">
                            <tr>
                                <td align="right">Name:
                                    <asp:Label ID="lblName_OfferMade" runat="server" /></td>
                                <td>Designation:
                                    <asp:Label ID="lblDesignation_OfferMade" runat="server" /></td>
                            </tr>
                            <tr>
                                <td align="right" style="height: 15px;">
                                    <br />
                                    <label>
                                        Email<span><asp:Label ID="lblReqEmail" Text="*" runat="server" ForeColor="Red"></asp:Label></span></label>
                                </td>
                                <td>
                                    <asp:TextBox ID="txtEmail" runat="server" MaxLength="40" Width="242px"
                                        Enabled="false" ReadOnly="true"></asp:TextBox>
                                    <br />
                                    <asp:RequiredFieldValidator ID="rqEmail" Display="Dynamic" runat="server" ControlToValidate="txtEmail"
                                        ValidationGroup="OfferMade" ForeColor="Red" ErrorMessage="Please Enter Email"></asp:RequiredFieldValidator>
                                    <asp:RegularExpressionValidator ID="reEmail" ControlToValidate="txtEmail" ValidationExpression="\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"
                                        Display="Dynamic" runat="server" ForeColor="Red" ErrorMessage="Please Enter a valid Email"
                                        ValidationGroup="OfferMade">
                                    </asp:RegularExpressionValidator>
                                </td>
                            </tr>
                            <tr>
                                <td align="right" style="height: 15px;">
                                    <label>
                                        Password<asp:Label ID="lblPassReq" runat="server" Text="*" ForeColor="Red"></asp:Label></label>
                                </td>
                                <td>
                                    <asp:TextBox ID="txtPassword1" runat="server" TextMode="Password" MaxLength="30"
                                        autocomplete="off" Width="242px"></asp:TextBox>
                                    <br />
                                    <label>
                                    </label>
                                    <asp:RequiredFieldValidator ID="rqPass" runat="server" ControlToValidate="txtPassword1"
                                        ValidationGroup="OfferMade" ForeColor="Red" Display="Dynamic" ErrorMessage="Please Enter Password"></asp:RequiredFieldValidator><br />
                                </td>
                            </tr>
                            <tr>
                                <td align="right" style="height: 15px;">
                                    <label>
                                        Confirm Password<asp:Label ID="lblConfirmPass" runat="server" Text="*" ForeColor="Red"></asp:Label></label>
                                </td>
                                <td>
                                    <asp:TextBox ID="txtpassword2" runat="server" TextMode="Password" autocomplete="off"
                                        MaxLength="30" EnableViewState="false" AutoCompleteType="None" Width="242px"></asp:TextBox>
                                    <br />
                                    <label>
                                    </label>
                                    <asp:CompareValidator ID="password" runat="server" ControlToValidate="txtpassword2"
                                        Display="Dynamic" ControlToCompare="txtPassword1" ForeColor="Red" ErrorMessage="Password didn't matched"
                                        ValidationGroup="OfferMade">
                                    </asp:CompareValidator>
                                    <asp:RequiredFieldValidator ID="rqConPass" runat="server" ControlToValidate="txtpassword2"
                                        ForeColor="Red" ValidationGroup="OfferMade" ErrorMessage="Enter Confirm Password"></asp:RequiredFieldValidator>

                                </td>
                            </tr>
                            <tr>
                                <td align="center" colspan="2">
                                    <asp:Button ID="btnSaveOfferMade" runat="server" BackColor="#327FB5" ForeColor="White" Height="32px"
                                        Style="height: 26px; font-weight: 700; line-height: 1em;" Text="Save" Width="100px" ValidationGroup="OfferMade"
                                        TabIndex="119" OnClick="btnSaveOfferMade_Click" />
                                    <asp:Button ID="btnCancelOfferMade" runat="server" Text="Cancel" OnClick="btnCancelInterview_Click" Width="100px"
                                        Style="height: 26px; font-weight: 700; line-height: 1em;"
                                        OnClientClick="javascript:document.getElementById('DivOfferMade').style.display='none';document.getElementById('DivOfferMadefade').style.display='none'" />
                                </td>
                            </tr>
                        </table>
                    </ContentTemplate>
                </asp:UpdatePanel>
            </div>
        </asp:Panel>
        <div id="DivOfferMadefade" class="black_overlay">
        </div>

        <asp:Panel ID="pnlUploadBulk" runat="server">
            <style>
                kTab {
                    :;
                }

                {
                    x;
                }
                /* END EXT
            </style>
            <div id="lightUploadBulk" class="white_content" style="text-align: center">
                <a class="close" href="#" onclick="CloseAddUserPopUp()">&times;</a>

                <asp:Panel ID="pnlDuplicate" runat="server">
                    <asp:Label ID="lblDuplicateCount" runat="server"></asp:Label>

                    <div style="padding: 20px; margin: auto;">
                        <center>
                                <table width="60%" class="uploadBulkTab" cellpadding="0">
                                <tr style="background-color: #A33E3F; color: white; font-weight: bold; text-align: center; width: 100%;">
                                    <td><span>Full Name</span></td>
                                    <td><span>Email</span></td>
                                    <td><span>Phone</span></td>
                                    <td><span>status</span></td>
                                </tr>
                                <asp:ListView ID="listDuplicateUsers" runat="server">
                                    <ItemTemplate>
                                        <tr>
                                            <td><span><%#Eval("FirstName")%>&nbsp;<%#Eval("LastName")%></span></td>
                                            <td><span><%#Eval("Email")%></span></td>
                                            <td><span><%#Eval("phone")%></span></td>
                                            <td><span><%#Eval("status")%></span></td>
                                        </tr>
                                    </ItemTemplate>
                                </asp:ListView>
                            </table>
                            </center>
                    </div>
                    <div style="padding: 20px; margin: auto;">
                        Email or Phone number of above users already exists, do you want to update the existing record?
                    </div>
                    <div style="padding: 10px; margin: auto;">
                        <asp:Button ID="btnYesEdit" runat="server" BackColor="#bb0000" ForeColor="White" Height="32px"
                            Style="height: 26px; font-weight: 700; line-height: 1em;" Text="Yes" Width="100px"
                            ValidationGroup="IndiCred" TabIndex="119" OnClick="btnYesEdit_Click" />
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <asp:Button ID="btnNoEdit" runat="server" BackColor="#bb0000" ForeColor="White" Height="32px"
                            Style="height: 26px; font-weight: 700; line-height: 1em;" Text="No" Width="100px"
                            ValidationGroup="IndiCred" TabIndex="119" OnClick="btnNoEdit_Click" />
                    </div>

                    <hr />
                    <br />
                </asp:Panel>

                <asp:Panel ID="pnlAddNewUser" runat="server">
                    <asp:Label ID="lblNewRecordAddedCount" runat="server"></asp:Label>
                    <center>
                    <table width="60%" class="uploadBulkTab" cellpadding="0" style="margin-top:20px;">
                        <tr style="background-color: #A33E3F; color: white; font-weight: bold; text-align: center; width: 100%;">
                            <td><span>Full Name</span></td>
                            <td><span>Email</span></td>
                            <td><span>Phone</span></td>
                            <td><span>status</span></td>
                        </tr>
                        <asp:ListView ID="lstNewUserAdd" runat="server">
                            <ItemTemplate>
                                <tr>
                                    <td><span><%#Eval("FirstName")%>&nbsp;<%#Eval("LastName")%></span></td>
                                    <td><span><%#Eval("Email")%></span></td>
                                    <td><span><%#Eval("phone")%></span></td>
                                    <td><span><%#Eval("status")%></span></td>
                                </tr>
                            </ItemTemplate>
                        </asp:ListView>
                    </table>
                    </center>
                </asp:Panel>

            </div>
        </asp:Panel>
        <div id="fadeUploadBulk" class="black_overlay">
        </div>
        <div id="interviewDatefade" class="black_overlay">
        </div>
    </div>
    <%--Modal Popup Ends--%>
    <%--Popup Stars--%>
    <div class="hide">
        <div id="divBulkUploadUserErrors" runat="server" title="Information" data-width="900px">
            <div style="padding: 5px 10px;">
                Below records contain empty values for mandatory fields. Please update cells marked by <span style="color: blue; font-weight: bold; text-align: center; font-size: 20px;">x</span> below in your file and upload again. If you see several empty rows at the end of the records, please delete those empty lines from your file.
            </div>
            <div style="max-height: 500px; height: 500px; overflow: auto;">
                <asp:GridView ID="grdBulkUploadUserErrors" runat="server" AutoGenerateColumns="false"
                    CssClass="table" Width="100%" CellSpacing="0" CellPadding="0" GridLines="Vertical">
                    <EmptyDataRowStyle ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle CssClass="trHeader " />
                    <RowStyle CssClass="FirstRow" BorderStyle="Solid" />
                    <AlternatingRowStyle CssClass="AlternateRow " />
                    <Columns>
                        <asp:TemplateField HeaderText="FirstName*" HeaderStyle-Width="75" ItemStyle-Width="75">
                            <ItemTemplate>
                                <%#grdBulkUploadUserErrors_GetCellText(Eval("FirstName")) %>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="FirstName*" HeaderStyle-Width="75" ItemStyle-Width="75">
                            <ItemTemplate>
                                <%#grdBulkUploadUserErrors_GetCellText(Eval("LastName")) %>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Email*" HeaderStyle-Width="90" ItemStyle-Width="90">
                            <ItemTemplate>
                                <%#grdBulkUploadUserErrors_GetCellText(Eval("Email")) %>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Designation*" HeaderStyle-Width="75" ItemStyle-Width="75">
                            <ItemTemplate>
                                <%#grdBulkUploadUserErrors_GetCellText(Eval("Designation")) %>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Status*" HeaderStyle-Width="50" ItemStyle-Width="50">
                            <ItemTemplate>
                                <%#grdBulkUploadUserErrors_GetCellText(Eval("Status")) %>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Source*" HeaderStyle-Width="60" ItemStyle-Width="60">
                            <ItemTemplate>
                                <%#grdBulkUploadUserErrors_GetCellText(Eval("Source")) %>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Primary Contact Phone*" HeaderStyle-Width="90" ItemStyle-Width="90">
                            <ItemTemplate>
                                <%#grdBulkUploadUserErrors_GetCellText(Eval("Phone1")) %>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Phone Type*" HeaderStyle-Width="60" ItemStyle-Width="60">
                            <ItemTemplate>
                                <%#grdBulkUploadUserErrors_GetCellText(Eval("Phone1Type")) %>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Zip*" HeaderStyle-Width="50" ItemStyle-Width="50">
                            <ItemTemplate>
                                <%#grdBulkUploadUserErrors_GetCellText(Eval("Zip")) %>
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
            <br />
        </div>
        <%--Send Email To User Popup--%>
        <div id="divSendEmailToUser" runat="server" title="Send Email">
            <asp:UpdatePanel ID="upSendEmailToUser" runat="server" UpdateMode="Conditional">
                <ContentTemplate>
                    <asp:ValidationSummary ID="vsEmailToUser" runat="server" ValidationGroup="vgEmailToUser" ShowSummary="False" ShowMessageBox="True" />
                    <fieldset>
                        <legend>
                            <asp:Label ID="lblEmailTo" runat="server" /><asp:HiddenField ID="hdnEmailTo" runat="server" />
                        </legend>
                        <table cellspacing="3" cellpadding="3" width="100%">
                            <tr>
                                <td>Subject:<br />
                                    <asp:TextBox ID="txtEmailSubject" runat="server" CssClass="textbox" Width="90%" ReadOnly="true" />
                                    <asp:RequiredFieldValidator ID="rfvEmailSubject" ValidationGroup="vgEmailToUser"
                                        runat="server" ControlToValidate="txtEmailSubject" ForeColor="Red" ErrorMessage="Please enter email subject." Display="None" />
                                </td>
                            </tr>
                            <tr>
                                <td>Custom Message:<br />
                                    <asp:TextBox ID="txtEmailCustomMessage" runat="server" CssClass="textbox" TextMode="MultiLine" Width="90%" />
                                    <asp:RequiredFieldValidator ID="rfvEmailCustomMessage" ValidationGroup="vgEmailToUser" Display="None"
                                        runat="server" ControlToValidate="txtEmailCustomMessage" ForeColor="Red" ErrorMessage="Please enter custom message for email." />
                                </td>
                            </tr>
                            <tr>
                                <td>Email Body:<br />
                                    <asp:TextBox ID="txtEmailBody" runat="server" CssClass="textbox" TextMode="MultiLine" Rows="4" Width="90%" />
                                    <asp:RequiredFieldValidator ID="rfvEmailBody" ValidationGroup="vgEmailToUser"
                                        runat="server" ControlToValidate="txtEmailBody" ForeColor="Red" ErrorMessage="Please enter email body." Display="None" />
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <div class="btn_sec">
                                        <asp:Button ID="btnSendEmailToUser" runat="server" ValidationGroup="vgEmailToUser" OnClick="btnSendEmailToUser_Click"
                                            CssClass="ui-button" Text="Send" />
                                        <asp:Button ID="btnCancelSendEmailToUser" runat="server" OnClick="btnCancelSendEmailToUser_Click"
                                            Text="Cancel" />
                                    </div>
                                </td>
                            </tr>
                        </table>
                    </fieldset>
                </ContentTemplate>
            </asp:UpdatePanel>
        </div>

        <div id="divChangeStatusForSelected" runat="server" title="Change Status">
            <asp:UpdatePanel ID="upChangeStatusForSelected" runat="server" UpdateMode="Conditional">
                <ContentTemplate>
                    <asp:ValidationSummary ID="vsChangeStatus" runat="server" ValidationGroup="vgChangeStatus" ShowMessageBox="true" ShowSummary="false" />
                    <div>
                        Status:
                        <asp:DropDownList ID="ddlStatus_Popup" runat="server" AutoPostBack="true" OnSelectedIndexChanged="ddlStatus_Popup_SelectedIndexChanged" />
                        <asp:RequiredFieldValidator ID="rfvStatus_Popup" runat="server" ErrorMessage="Please select recruiter." InitialValue="0"
                            ControlToValidate="ddlStatus_Popup" ValidationGroup="vgChangeStatus" Display="None" />
                    </div>
                    <br />
                    <div id="divInterviewDate" runat="server" visible="false">
                        Recruiter:
                        <asp:DropDownList ID="ddlRecruiter_Popup" runat="server" />
                        <asp:RequiredFieldValidator ID="rfvRecruiter_Popup" runat="server" ErrorMessage="Please select recruiter." InitialValue="0"
                            ControlToValidate="ddlRecruiter_Popup" ValidationGroup="vgChangeStatus" Display="None" />
                    </div>
                    <br />
                    <asp:GridView ID="grdUsers_Popup" runat="server" AutoGenerateColumns="false" DataKeyNames="Id,DesignationID"
                        CssClass="table" Width="100%" CellSpacing="0" CellPadding="0" GridLines="Vertical"
                        OnRowDataBound="grdUsers_Popup_RowDataBound">
                        <EmptyDataRowStyle ForeColor="White" HorizontalAlign="Center" />
                        <HeaderStyle CssClass="trHeader " />
                        <RowStyle CssClass="FirstRow" BorderStyle="Solid" />
                        <AlternatingRowStyle CssClass="AlternateRow " />
                        <Columns>
                            <asp:TemplateField HeaderText="Name" ItemStyle-Width="150">
                                <ItemTemplate>
                                    <asp:Literal ID="ltrlFirstName" runat="server" Text='<%#Eval("FirstName") %>' />&nbsp;<asp:Literal ID="ltrlLastName" runat="server" Text='<%#Eval("LastName") %>' />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Designation" ItemStyle-Width="150">
                                <ItemTemplate>
                                    <asp:Literal ID="ltrlDesignation" runat="server" Text='<%#Eval("Designation") %>' />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Interview" Visible="false">
                                <ItemTemplate>
                                    <asp:TextBox ID="txtInterviewDate" placeholder="Select Date" runat="server"
                                        onkeypress="return false" Width="127px" Text='<%#Eval("InterviewDate")%>' />
                                    <cc1:CalendarExtender ID="ceInterviewDate" TargetControlID="txtInterviewDate" Format="MM/dd/yyyy" runat="server" />
                                    <asp:RequiredFieldValidator ID="rfvInterviewDate" runat="server" ErrorMessage="Please select interview date."
                                        ControlToValidate="txtInterviewDate" ValidationGroup="vgChangeStatus" Display="None" />
                                    <br />
                                    <asp:DropDownList ID="ddlInterviewTime" runat="server" Width="112px" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Tech Task" Visible="false" ItemStyle-Width="100">
                                <ItemTemplate>
                                    Tech Task:
                                    <asp:DropDownList ID="ddlTechTask" runat="server" Width="95" AutoPostBack="true" OnSelectedIndexChanged="grdUsers_Popup_ddlTechTask_SelectedIndexChanged" />
                                    <asp:RequiredFieldValidator ID="rfvTechTask" runat="server" ErrorMessage="Please select tech task." InitialValue="0"
                                        ControlToValidate="ddlTechTask" ValidationGroup="vgChangeStatus" Display="None" />
                                    <br />
                                    Sub Task:
                                    <asp:DropDownList ID="ddlTechSubTask" runat="server" Width="95" />
                                    <asp:RequiredFieldValidator ID="rfvTechSubTask" runat="server" ErrorMessage="Please select tech sub task." InitialValue="0"
                                        ControlToValidate="ddlTechSubTask" ValidationGroup="vgChangeStatus" Display="None" />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Reason" Visible="false">
                                <ItemTemplate>
                                    <asp:TextBox ID="txtReason" runat="server" TextMode="MultiLine" Rows="3" Width="95%" />
                                    <asp:RequiredFieldValidator ID="rfvReason" runat="server" ErrorMessage="Please enter reason for status change."
                                        ControlToValidate="txtReason" ValidationGroup="vgChangeStatus" Display="None" />
                                </ItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                    <br />
                    <div class="btn_sec">
                        <asp:Button ID="btnSaveStatusForSelected" runat="server" Text="Change Status" ValidationGroup="vgChangeStatus"
                            OnClick="btnSaveStatusForSelected_Click" />&nbsp;
                    <asp:Button ID="btnCancelChangeStatusForSelected" runat="server" Text="Cancel" OnClick="btnCancelChangeStatusForSelected_Click" />
                    </div>
                </ContentTemplate>
            </asp:UpdatePanel>
        </div>
    </div>
    <%--Popup Ends--%>

    <script src="../js/jquery.dd.min.js"></script>
    <script type="text/javascript">

        Dropzone.autoDiscover = false;

        var prmTaskGenerator = Sys.WebForms.PageRequestManager.getInstance();

        prmTaskGenerator.add_beginRequest(function () {
            DestroyCKEditors();
        });

        prmTaskGenerator.add_endRequest(function () {
            EditUser_Initialize();
        });

        $(document).ready(function () {
            EditUser_Initialize();
        });

        function EditUser_Initialize() {

            SetSalesUserAutoSuggestion();
            SetSalesUserAutoSuggestionUI();

            ApplyDropZone();

            try {
                $("#<%=ddlUserStatus.ClientID%>").msDropDown();
                $(".grd-status").msDropDown();
            } catch (e) {
                alert(e.message);
            }
        }

        var objBulkUploadFileDropzone;

        function ApplyDropZone() {
            //debugger;
            ////User's drag and drop file attachment related code

            //remove already attached dropzone.
            if (objBulkUploadFileDropzone) {
                objBulkUploadFileDropzone.destroy();
                objBulkUploadFileDropzone = null;
            }
            objBulkUploadFileDropzone = GetWorkFileDropzone("div#divBulkUploadFile", 'div#divBulkUploadFilePreview', '#<%= hdnBulkUploadFile.ClientID %>');
        }

        function SetSalesUserAutoSuggestion() {

            $("#<%=txtSearch.ClientID%>").catcomplete({
                delay: 500,
                source: function (request, response) {
                    $.ajax({
                        type: "POST",
                        url: "ajaxcalls.aspx/GetSalesUserAutoSuggestion",
                        dataType: "json",
                        contentType: "application/json; charset=utf-8",
                        data: JSON.stringify({ searchterm: request.term }),
                        success: function (data) {
                            // Handle 'no match' indicated by [ "" ] response
                            if (data.d) {

                                response(data.length === 1 && data[0].length === 0 ? [] : JSON.parse(data.d));
                            }
                            // remove loading spinner image.                                
                            $("#<%=txtSearch.ClientID%>").removeClass("ui-autocomplete-loading");
                        }
                    });
                },
                minLength: 2,
                select: function (event, ui) {
                    $("#<%=txtSearch.ClientID%>").val(ui.item.value);
                    //TriggerSearch();
                    $('#<%=btnSearchGridData.ClientID%>').click();
                }
            });
        }

        function SetSalesUserAutoSuggestionUI() {

            $.widget("custom.catcomplete", $.ui.autocomplete, {
                _create: function () {
                    this._super();
                    this.widget().menu("option", "items", "> :not(.ui-autocomplete-category)");
                },
                _renderMenu: function (ul, items) {
                    var that = this,
                      currentCategory = "";
                    $.each(items, function (index, item) {
                        var li;
                        if (item.Category != currentCategory) {
                            ul.append("<li class='ui-autocomplete-category'> Search " + item.Category + "</li>");
                            currentCategory = item.Category;
                        }
                        li = that._renderItemData(ul, item);
                        if (item.Category) {
                            li.attr("aria-label", item.Category + " : " + item.label);
                        }
                    });

                }
            });
        }

        function grdUsers_Email_OnClick(sender, email) {
            $('#<%=lblEmailTo.ClientID%>').html(email);
            $('#<%=hdnEmailTo.ClientID%>').val(email);
            <%--SetCKEditor('<%=txtEmailHeader.ClientID%>');
            SetCKEditor('<%=txtEmailBody.ClientID%>');
            SetCKEditor('<%=txtEmailFooter.ClientID%>');--%>
            ShowPopupWithTitle('#<%=divSendEmailToUser.ClientID%>', 'Send Email');
        }

    </script>
</asp:Content>
