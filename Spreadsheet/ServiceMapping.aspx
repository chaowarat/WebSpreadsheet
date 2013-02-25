<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ServiceMapping.aspx.cs" Inherits="Spreadsheet.ServiceMapping" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        .auto-style1 {
            width: 166px;
        }
        .auto-style5 {
            width: 810px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <table style="width:100%;">
            <tr>
                <td class="auto-style1">&nbsp;</td>
                <td style="text-align: center" class="auto-style5">Service Mapping</td>
                <td>&nbsp;</td>
            </tr>
            <tr>
                <td class="auto-style1">&nbsp;</td>
                <td class="auto-style5">
                    <asp:GridView ID="GridView_mapping" runat="server" AutoGenerateColumns="False" OnRowDataBound="GridView_mapping_RowDataBound">
                        <Columns>
                            <asp:TemplateField HeaderText="Service Code">
                                <ItemTemplate>
                                    <asp:Label runat="server" Text="" ID="Label_service_id"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="Service Name">
                                <ItemTemplate>
                                    <asp:Label runat="server" Text="" ID="Label_service_name"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:BoundField DataField="DisabilityCode" HeaderText="DisabilityCode" />
                            <asp:BoundField DataField="Name" HeaderText="DisabilityName" />
                            <asp:TemplateField HeaderText="Select">
                            <ItemTemplate>
                                <asp:CheckBox ID="CheckBox_select" runat="server" />
                            </ItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                    <br />
                </td>
                <td>&nbsp;</td>
            </tr>
            <tr>
                <td class="auto-style1">&nbsp;</td>
                <td class="auto-style5" style="text-align: center">
                    <asp:Button ID="Button_ok" runat="server" OnClick="Button_ok_Click" Text="OK" Width="92px" />
                </td>
                <td>&nbsp;</td>
            </tr>
        </table>
    
    </div>
    </form>
</body>
</html>
