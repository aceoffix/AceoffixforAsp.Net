using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Drawing;

public partial class mergeExcelCell_excel : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Aceoffix.ExcelWriter.Workbook wb = new Aceoffix.ExcelWriter.Workbook();
        Aceoffix.ExcelWriter.Sheet sheet = wb.OpenSheet("Sheet1");
        sheet.OpenTable("B2:F2").Merge();

        Aceoffix.ExcelWriter.Cell cellB2 = sheet.OpenCell("B2");
        cellB2.Value = "Products for Sale";
        cellB2.HorizontalAlignment = Aceoffix.ExcelWriter.XlHAlign.xlHAlignCenter;
        cellB2.ForeColor = Color.Red;
        cellB2.Font.Size = 16;

        sheet.OpenTable("B4:B6").Merge();
        Aceoffix.ExcelWriter.Cell cellB4 = sheet.OpenCell("B4");
        cellB4.Value = "Product A:";
        cellB4.HorizontalAlignment = Aceoffix.ExcelWriter.XlHAlign.xlHAlignCenter;
        cellB4.VerticalAlignment = Aceoffix.ExcelWriter.XlVAlign.xlVAlignCenter;
        cellB4.ForeColor = Color.Red;

        sheet.OpenTable("B7:B9").Merge();
        Aceoffix.ExcelWriter.Cell cellB7 = sheet.OpenCell("B7");
        cellB7.Value = "Product B:";
        cellB7.HorizontalAlignment = Aceoffix.ExcelWriter.XlHAlign.xlHAlignCenter;
        cellB7.VerticalAlignment = Aceoffix.ExcelWriter.XlVAlign.xlVAlignCenter;
        cellB7.ForeColor = Color.Red;

        AceoffixCtrl1.Bind(wb);
        AceoffixCtrl1.Menubar = false;
        AceoffixCtrl1.CustomToolbar = false;
        AceoffixCtrl1.OpenDocument("doc/test.xls", Aceoffix.OpenModeType.xlsNormalEdit, "John Scott");

    }
}