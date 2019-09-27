using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class excelFillCell_excel : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Aceoffix.ExcelWriter.Workbook wb = new Aceoffix.ExcelWriter.Workbook();
        Aceoffix.ExcelWriter.Sheet sheet = wb.OpenSheet("Sheet1");
        Aceoffix.ExcelWriter.Cell cellB4 = sheet.OpenCell("B4");
        cellB4.Value = "Jan";
        Aceoffix.ExcelWriter.Cell cellD2 = sheet.OpenCell("D2");
        cellD2.Value = "Sales Report (2015)";
        Aceoffix.ExcelWriter.Cell cellF14 = sheet.OpenCell("F14");
        cellF14.Value = "100%";
      
        AceoffixCtrl1.Bind(wb);

        AceoffixCtrl1.Menubar = false;
        AceoffixCtrl1.CustomToolbar = false;
        AceoffixCtrl1.OpenDocument("doc/test.xls", Aceoffix.OpenModeType.xlsNormalEdit, "John Scott");

    }
}