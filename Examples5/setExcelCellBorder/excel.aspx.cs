using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aceoffix.ExcelWriter;
using System.Drawing;

public partial class setExcelCellBorder_excel : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Workbook wb = new Workbook();
        Sheet sheet = wb.OpenSheet("Sheet1");
        Aceoffix.ExcelWriter.Table table = sheet.OpenTable("A1:P200");
        table.Border.LineColor = Color.White;


        Border B13Border = sheet.OpenTable("B13:B15").Border;
        B13Border.Weight = XlBorderWeight.xlThin;
        B13Border.LineColor=Color.Blue;
        B13Border.BorderType=XlBorderType.xlAllEdges;

        Border B6Border = sheet.OpenTable("B6:B8").Border;
        B6Border.Weight=XlBorderWeight.xlThick;
        B6Border.LineColor=Color.Green;
        B6Border.LineStyle=XlBorderLineStyle.xlSlantDashDot;
        B6Border.BorderType=XlBorderType.xlAllEdges;

        Border B9Border = sheet.OpenTable("B9:B12").Border;
        B9Border.Weight=XlBorderWeight.xlThick;
        B9Border.LineColor=Color.DarkGray;
        B9Border.LineStyle=XlBorderLineStyle.xlSlantDashDot;
        B9Border.BorderType=XlBorderType.xlAllEdges;

        Aceoffix.ExcelWriter.Table titleTable = sheet.OpenTable("B4:D5");
        titleTable.Border.Weight = Aceoffix.ExcelWriter.XlBorderWeight.xlThick;
        titleTable.Border.LineColor = Color.Red;
        titleTable.Border.BorderType = Aceoffix.ExcelWriter.XlBorderType.xlAllEdges;
        Aceoffix.ExcelWriter.Table bodyTable2 = sheet.OpenTable("B6:D21");
        bodyTable2.Border.Weight = Aceoffix.ExcelWriter.XlBorderWeight.xlThick;
        bodyTable2.Border.LineColor = Color.Red;
        bodyTable2.Border.BorderType = Aceoffix.ExcelWriter.XlBorderType.xlAllEdges;
        
        AceoffixCtrl1.Bind(wb);
        AceoffixCtrl1.Menubar = false;
        AceoffixCtrl1.CustomToolbar=false;
        AceoffixCtrl1.SaveFilePage = "savefile.aspx";
        AceoffixCtrl1.OpenDocument("doc/test.xls", Aceoffix.OpenModeType.xlsNormalEdit, "John Scott");
 
    }
}