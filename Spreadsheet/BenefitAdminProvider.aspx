<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="BenefitAdminProvider.aspx.cs" Inherits="Spreadsheet.BenefitAdminProvider" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:fb="http://ogp.me/ns/fb#">
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
    <head id="Head1"  runat="server">
        <title>Benefit Admin Spreadsheet</title>
        <!--Required-->
        <link rel="stylesheet" type="text/css" href="css/page-style.css" />
        <link rel="stylesheet" type="text/css" href="css/jquery.sheet.css" />
        <link rel="stylesheet" type="text/css" href="jquery-ui/theme/jquery-ui.css" />
        <script type="text/javascript" src="jquery/jquery-1.5.2.js"></script>
        <script type="text/javascript" src="jquery/jquery.sheet.js"></script>
        <script type="text/javascript" src="jquery/parser.js"></script>
        <script type="text/javascript" src="jquery/JScript-header.js"></script>
		<script type="text/javascript" src="plugins/jquery.scrollTo-min.js"></script>
		<script type="text/javascript" src="jquery-ui/ui/jquery-ui.min.js"></script>
		<script type="text/javascript" src="plugins/raphael-min.js"></script>
		<script type="text/javascript" src="plugins/g.raphael-min.js"></script>
		<link rel="stylesheet" type="text/css" href="plugins/jquery.colorPicker.css" />
		<script type="text/javascript" src="plugins/jquery.colorPicker.min.js"></script>
		<script type="text/javascript" src="plugins/jquery.elastic.min.js"></script>
		<script type="text/javascript" src="jquery/jquery.sheet.advancedfn.js"></script>
		<script type="text/javascript" src="jquery/jquery.sheet.financefn.js"></script>


        <style type="text/css">
            .style1
            {
                width: 131px;
            }
            .style2
            {
                width: 828px;
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
        		Benefit Admin Provider Spreadsheet</a>
        </h1>
        
        <table bgcolor="White" style="width: 100%; height: 0px;" width="100%">
            <tr>
                <td class="style1">
                    &nbsp;</td>
                <td class="style5">
                    <asp:Label ID="Label_user" runat="server" Font-Size="Medium" Text="Label"></asp:Label>
                    <br />
                    <asp:LinkButton ID="LinkButton_signOut" runat="server" 
            onclick="LinkButton_signOut_Click">Sign out</asp:LinkButton>
                </td>
            </tr>
        </table>
        <div 
            id="jQuerySheet0" class="jQuerySheet" style="height: 500px; width: 99%; left: 5px;">
        </div>
        
        <span id="inlineMenu" style="display: none;">
			<span>

            <!-- Button save -->
            <input id="Button2" type="button" value="Save" onclick="sheetInstance.getXMLSource2(sheetInstance.HTMLtoPrettySource(jQuery(sheetInstance.exportSheet.xml(true))[0]));" />

            &nbsp;&nbsp; 
            <a href = "#" onclick="sheetInstance.print2(sheetInstance.HTMLtoPrettySource(jQuery(sheetInstance.exportSheet.xml(true))[0])); return false;">View XML</a>
            <a href = "#" onclick="sheetInstance.viewSource(true); return false;">View HTML</a>
            <a href = "#" onclick="sheetInstance.print3(JSON.stringify(sheetInstance.exportSheet.json())); return false">View Json</a>

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
				<img alt="Web Link" src="images/arrow_out.png"/></a><!--<a href="#" onclick="insertAt('jSheetControls_formula', '~np~text~'+'/np~');return false;" title="Non-parsed"><img alt="Non-parsed" src="images/noparse.png"/></a>-->
		    </span>
		</span>
        </form>
</body>
</html>
