<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="BenefitAdminCode.aspx.cs" Inherits="Spreadsheet.BenefitAdminCode" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
    <head id="Head1"  runat="server">
        <title>Benefit Admin Spreadsheet</title>

        <META HTTP-EQUIV="CACHE-CONTROL" CONTENT="NO-CACHE">
        <META HTTP-EQUIV="EXPIRES" CONTENT="0">

        <!--Required-->
        <link rel="stylesheet" type="text/css" href="css/jquery.sheet.css" />
        <link rel="stylesheet" type="text/css" href="jquery-ui/theme/jquery-ui.min.css" />
        <script type="text/javascript" src="jquery/jquery-1.8.3.min.js">
        </script>
        <script type="text/javascript" src="plugins/globalize.js">
        </script>
        <script type="text/javascript" src="jquery/jquery.sheet.js">
        </script>
        <script type="text/javascript" src="parser/parser.js">
        </script>
        <!--/Required-->
		
		<!--Optional-->
			<!--MouseWheel-->
				<script type="text/javascript" src="plugins/jquery.mousewheel.min.js"></script>
			<!--/MouseWheel-->
            <!--Nearest-->
            <script type="text/javascript" src="plugins/jquery.nearest.min.js"></script>
            <!--/Nearest-->
			<!--jQuery UI-->
				<script type="text/javascript" src="jquery-ui/ui/jquery-ui.min.js"></script>
			<!--/jQuery UI-->
			<!--Raphaeljs - for charts-->
				<script type="text/javascript" src="plugins/raphael-min.js"></script>
				<script type="text/javascript" src="plugins/g.raphael-min.js"></script>
			<!--/Raphaeljs-->
			<!--ColorPicker-->
				<link rel="stylesheet" type="text/css" href="plugins/jquery.colorPicker.css" />
				<script type="text/javascript" src="plugins/jquery.colorPicker.min.js"></script>
			<!--/ColorPicker-->
			<!--Theme Switcher Tool-->
				<!--<script type="text/javascript" src="http://jqueryui.com/themeroller/themeswitchertool/"></script>-->
			<!--/Theme Switcher Tool-->
			<!--Elastic-->
				<script type="text/javascript" src="plugins/jquery.elastic.min.js"></script>
			<!--/Elastic-->
			<!--Advanced_Math-->
				<script type="text/javascript" src="plugins/jquery.sheet.advancedfn.js"></script>
			<!--/Advanced_Math-->
			<!--Finance-->
				<script type="text/javascript" src="plugins/jquery.sheet.financefn.js"></script>
			<!--/Finance-->
		<!--/Optional-->

        <link rel="stylesheet" type="text/css" href="css/page-style.css" />     
        <script type="text/javascript" src="jquery/JScript-header.js"></script>
        <style type="text/css">

            .style1
            {
                width: 131px;
            }
            .style2
            {
                width: 830px;
            }
            .style5
            {
                width: 227px;
                text-align:right;
            }
        </style>

    </head>

<body>
    <form id="form1" runat="server">
        <h1 id="header" class="ui-state-default">
        	<a>
        		Benefit Admin Code Spreadsheet
        	</a>
        </h1>
        <table bgcolor="White" style="width: 100%; height: 0px;" width="100%">
            <tr>
                <td class="style1">
                    <asp:RadioButtonList ID="RadioButtonList_type" runat="server" Width="200px" 
                        RepeatDirection="Horizontal">
                        <asp:ListItem Selected="True">Insert</asp:ListItem>
                        <asp:ListItem Value="CreateNew">Create New Sheet</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
                <td class="style2">

        <asp:FileUpload ID="FileUpload1" runat="server" style="right" BorderStyle="Groove" 
                        Height="21px" Width="220px"/>

        <asp:Button ID="btnUpload" runat="server" Height="21px" 
            onclick="btnUpload_Click" Text="Upload" Width="86px" />
                    &nbsp;<asp:Label ID="LabelFile" runat="server" Text="File sheet"></asp:Label>
                    &nbsp;<asp:DropDownList ID="DropDownListFrom" runat="server" Height="18px" 
                        Width="132px">
                    </asp:DropDownList>
&nbsp;To sheet
                    <asp:DropDownList ID="DropDownListTo" runat="server" Height="18px" 
                        Width="132px">
                        <asp:ListItem>Service</asp:ListItem>
                        <asp:ListItem>Activity</asp:ListItem>
                        <asp:ListItem>SubActivity</asp:ListItem>
                        <asp:ListItem>Material</asp:ListItem>
                    </asp:DropDownList>
                    <asp:Button ID="OK" runat="server" onclick="ButtonOK_Click" Text="OK" 
                        Width="86px" Height="21px" />
                    <br />

                </td>
                <td class="style5">
                    <asp:Label ID="Label_user" runat="server" Font-Size="Medium" Text="Label"></asp:Label>
                     <br />
                     <asp:LinkButton ID="LinkButton_signOut" runat="server" 
                     onclick="LinkButton_signOut_Click">Sign out</asp:LinkButton>
                </td>
            </tr>
        </table>
        &nbsp;<div 
            id="jQuerySheet0" class="jQuerySheet" style="height: 500px; width: 97%; left: 5px;">
        </div>

        <span id="inlineMenu" style="display: none;">
			<span>
                  
            <img id="saving" src="images/saving.gif" alt="Klematis" width="20" height="20" style="margin-bottom: -5px;" />
            
            <!-- Button save -->
            <input id="Button2" type="button" value="Save" onclick="sheetInstance.getXMLSource1(sheetInstance.HTMLtoPrettySource(jQuery(sheetInstance.exportSheet.xml(true))[0]));" />
            <input id="buttonHidden" type="hidden"  onclick="sheetInstance.s.fnOpen(htmlData); return false;" />
            &nbsp;&nbsp; 

			<a href="#" onclick="sheetInstance.controlFactory.addRow(); return false;" title="Insert Row After Selected">
					<img alt="Insert Row After Selected" src="images/sheet_row_add.png"/></a>
				<a href="#" onclick="sheetInstance.controlFactory.addRow(null, true); return false;" title="Insert Row Before Selected">
					<img alt="Insert Row Before Selected" src="images/sheet_row_add.png"/></a>
				<a href="#" onclick="sheetInstance.controlFactory.addRow(null, null, ':last'); return false;" title="Add Row At End">
					<img alt="Add Row" src="images/sheet_row_add.png"/></a>
				<a href="#" onclick="sheetInstance.controlFactory.addRowMulti(); return false;" title="Add Multi-Rows">
					<img alt="Add Multi-Rows" src="images/sheet_row_add_multi.png"/></a>
				<a href="#" onclick="sheetInstance.deleteRow(); return false;" title="Delete Row">
					<img alt="Delete Row" src="images/sheet_row_delete.png"/></a>
				<a href="#" onclick="sheetInstance.controlFactory.addColumn(); return false;" title="Insert Column After Selected">
					<img alt="Insert Column After Selected" src="images/sheet_col_add.png"/></a>
				<a href="#" onclick="sheetInstance.controlFactory.addColumn(null, true); return false;" title="Insert Column Before Selected">
					<img alt="Insert Column Before Selected" src="images/sheet_col_add.png"/></a>
				<a href="#" onclick="sheetInstance.controlFactory.addColumn(null, null, ':last'); return false;" title="Add Column At End">
					<img alt="Add Column At End" src="images/sheet_col_add.png"/></a>
				<a href="#" onclick="sheetInstance.controlFactory.addColumnMulti(); return false;" title="Insert Multi-Columns">
					<img alt="Add Multi-Columns" src="images/sheet_col_add_multi.png"/></a>
				<a href="#" onclick="sheetInstance.deleteColumn(); return false;" title="Delete Column">
					<img alt="Delete Column" src="images/sheet_col_delete.png"/></a>
				<a href="#" onclick="sheetInstance.getTdRange(null, sheetInstance.obj.formula().val()); return false;" title="Get Cell Range">
					<img alt="Get Cell Range" src="images/sheet_get_range.png"/></a>
				<a href="#" onclick="sheetInstance.s.fnSave(); return false;" title="Save Sheets">
					<img alt="Save Sheet" src="images/disk.png"/></a>
				<a href="#" onclick="sheetInstance.deleteSheet(); return false;" title="Delete Current Sheet">
					<img alt="Delete Current Sheet" src="images/table_delete.png"/></a>
				<a href="#" onclick="sheetInstance.calc(sheetInstance.i); return false;" title="Refresh Calculations">
					<img alt="Refresh Calculations" src="images/arrow_refresh.png"/></a>
				<a href="#" onclick="sheetInstance.cellFind(); return false;" title="Find">
					<img alt="Find" src="images/find.png"/></a>
				<a href="#" onclick="sheetInstance.cellStyleToggle('styleBold'); return false;" title="Bold">
					<img alt="Bold" src="images/text_bold.png"/></a>
				<a href="#" onclick="sheetInstance.cellStyleToggle('styleItalics'); return false;" title="Italic">
					<img alt="Italic" src="images/text_italic.png"/></a>
				<a href="#" onclick="sheetInstance.cellStyleToggle('styleUnderline', 'styleLineThrough'); return false;" title="Underline">
					<img alt="Underline" src="images/text_underline.png"/></a>
				<a href="#" onclick="sheetInstance.cellStyleToggle('styleLineThrough', 'styleUnderline'); return false;" title="Strikethrough">
					<img alt="Strikethrough" src="images/text_strikethrough.png"/></a>
				<a href="#" onclick="sheetInstance.cellStyleToggle('styleLeft', 'styleCenter styleRight'); return false;" title="Align Left">
					<img alt="Align Left" src="images/text_align_left.png"/></a>
				<a href="#" onclick="sheetInstance.cellStyleToggle('styleCenter', 'styleLeft styleRight'); return false;" title="Align Center">
					<img alt="Align Center" src="images/text_align_center.png"/></a>
				<a href="#" onclick="sheetInstance.cellStyleToggle('styleRight', 'styleLeft styleCenter'); return false;" title="Align Right">
					<img alt="Align Right" src="images/text_align_right.png"/></a>
				<a href="#" onclick="sheetInstance.fillUpOrDown(); return false;" title="Fill Down">
					<img alt="Fill Down" src="images/arrow_down.png"/></a>
				<a href="#" onclick="sheetInstance.fillUpOrDown(true); return false;" title="Fill Up">
					<img alt="Fill Up" src="images/arrow_up.png"/></a>
				<span class="colorPickers">
					<input title="Foreground color" class="colorPickerFont" style="background-image: url('images/palette.png') ! important; width: 16px; height: 16px;"/>
					<input title="Background Color" class="colorPickerCell" style="background-image: url('images/palette_bg.png') ! important; width: 16px; height: 16px;"/>
				</span>
				<a href="#" onclick="sheetInstance.obj.formula().val('=HYPERLINK(\'' + prompt('Enter Web Address', 'http://www.visop-dev.com/') + '\')').keydown(); return false;" title="HyperLink">
					<img alt="Web Link" src="images/page_link.png"/></a>
				<a href="#" onclick="sheetInstance.toggleFullScreen(); $('#lockedMenu').toggle(); return false;" title="Toggle Full Screen">
					<img alt="Web Link" src="images/arrow_out.png"/></a>
                </span>
		</span>

    </form>
</body>
</html>
