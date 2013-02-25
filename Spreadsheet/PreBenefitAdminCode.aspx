<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="PreBenefitAdminCode.aspx.cs" Inherits="Spreadsheet.PreBenefitAdminCode" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        .auto-style1 {
            width: 402px;
        }
        .auto-style2 {
            width: 402px;
            height: 153px;
        }
        .auto-style3 {
            height: 153px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table style="width:100%;">
            <tr>
                <td class="auto-style2"></td>
                <td class="auto-style3">


        <asp:RadioButtonList ID="RadioButtonList_type" runat="server"></asp:RadioButtonList>

                </td>
                <td class="auto-style3"></td>
            </tr>
            <tr>
                <td class="auto-style1">&nbsp;</td>
                <td>
                    <asp:Button ID="Button_ok" runat="server" OnClick="Button_ok_Click" Text="OK" Width="67px" />
                </td>
                <td>&nbsp;</td>
            </tr>
            <tr>
                <td class="auto-style1">&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
            </tr>
        </table>
        <br />
        <br />
        <br />

    </div>
    </form>
</body>
</html>
