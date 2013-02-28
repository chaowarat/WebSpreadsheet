<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ServiceMapping.aspx.cs" Inherits="Spreadsheet.ServiceMapping" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        #content{
            height:auto;
            width:85%;
            margin:80px;
        }
        #left{
            float:left;
        }
        #right{
            float:right;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div id="content">
        <h1 style="text-align: center; margin-top: -50px;">Service mapping page</h1>
        <div id="left">
            <asp:GridView ID="GridView_service" runat="server" AutoGenerateColumns="False">
                <Columns>
                    <asp:BoundField DataField="SVCCode" HeaderText="SVCCode" ItemStyle-Width="180px" />
                    <asp:BoundField DataField="SVCName" HeaderText="SVCName" ItemStyle-Width="350px" />
                </Columns>
            </asp:GridView>
        </div>
        <div id="right">
            <asp:GridView ID="GridView_mapping" runat="server" AutoGenerateColumns="False">
                <Columns>
                    <asp:TemplateField HeaderText="Select">
                    <ItemTemplate>
                        <asp:CheckBox ID="CheckBox_select" runat="server" />
                    </ItemTemplate>
                    </asp:TemplateField>
                    <asp:BoundField DataField="DisabilityCode" HeaderText="DisabilityCode" ItemStyle-Width="180px" />
                    <asp:BoundField DataField="Name" HeaderText="DisabilityName" ItemStyle-Width="350px" />
                </Columns>
            </asp:GridView>
            <div style="text-align: left; position: relative; margin-top: 20px;">
                <asp:Button ID="Button_ok" runat="server" OnClick="Button_ok_Click" Text="OK" Width="92px" />
            </div>
        </div>
    </div>
    </form>
</body>
</html>
