<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="PreBenefitAdminCode.aspx.cs" Inherits="Spreadsheet.PreBenefitAdminCode" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div style="text-align: center; margin-top: 50px">
        <h1>กรูณาเลือกประเภทความด้อยโอกาส</h1>
    </div>
    <div style="margin-left: 40%; margin-top: 20px;">
        <asp:RadioButtonList ID="RadioButtonList_type" runat="server"></asp:RadioButtonList>
    </div>
    <div style="margin-top: 20px; margin-left: 45%;"> 
        <asp:Button ID="Button_ok" runat="server" OnClick="Button_ok_Click" Text="OK" Width="67px" />
    </div>
    </form>
</body>
</html>
