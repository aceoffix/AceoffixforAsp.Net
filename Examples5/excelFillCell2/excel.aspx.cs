using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aceoffix.ExcelWriter;
using System.Drawing;

public partial class excelFillCell2_excel : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Workbook wb = new Workbook();
        Sheet sheet = wb.OpenSheet("Sheet1");
        Aceoffix.ExcelWriter.Cell cellB4 = sheet.OpenCell("B4");
        cellB4.Value = "Jan";
        cellB4.ForeColor = Color.Red;
        Aceoffix.ExcelWriter.Cell cellD2 = sheet.OpenCell("D2");
        cellD2.Value = "Sales Report (2015)";
        cellD2.ForeColor = Color.Green;
        Aceoffix.ExcelWriter.Cell cellF14 = sheet.OpenCell("F14");
        cellF14.Value = "100%";
        cellF14.ForeColor = Color.FromArgb(255, 64, 0);
       
        AceoffixCtrl1.Bind(wb);
      
        AceoffixCtrl1.CustomToolbar=false;

        AceoffixCtrl1.Menubar = false;
        AceoffixCtrl1.SaveDataPage = "SaveData.aspx";
        AceoffixCtrl1.OpenDocument("doc/test.xls", Aceoffix.OpenModeType.xlsSubmitForm, "John Scott");

    }
}