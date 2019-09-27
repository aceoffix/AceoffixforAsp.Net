using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class excel_excelsubmit : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    protected void AceoffixCtrl1_Load(object sender, EventArgs e)
    {
        Aceoffix.ExcelWriter.Workbook wb = new Aceoffix.ExcelWriter.Workbook();
        Aceoffix.ExcelWriter.Sheet sheet1 = wb.OpenSheet("Purchase Order");

        System.Random rd = new Random(System.DateTime.Now.Millisecond);
        sheet1.OpenCell("D8").Value = "XYZ-11-" + (10001 + rd.Next(10000)).ToString();
        sheet1.OpenCell("D8").ReadOnly = true;

        Aceoffix.ExcelWriter.Cell cellDate = sheet1.OpenCell("I6");
        string dt = DateTime.Now.ToString();
        cellDate.Value = dt;
        sheet1.OpenCell("J31").Value = dt;

        sheet1.OpenCell("D15").ReadOnly = false; // Enabled edit

        Aceoffix.ExcelWriter.Table table1 = sheet1.OpenTable("B18:J23");
        table1.ReadOnly = false; // Enabled edit

        sheet1.OpenCell("J24").ReadOnly = true;
        sheet1.OpenCell("J26").ReadOnly = true;
        sheet1.OpenCell("J28").ReadOnly = true;

        sheet1.OpenCell("G31").Value = "John Scott";

        AceoffixCtrl1.Caption = "Excel spreadsheet used as an input form.";
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        // Create custom toolbar
        AceoffixCtrl1.AddCustomToolButton("Save", "SaveDocument()", 1);
        AceoffixCtrl1.AddCustomToolButton("-", "", 0);
        AceoffixCtrl1.AddCustomToolButton("Print", "ShowPrintDlg()", 6);
        AceoffixCtrl1.AddCustomToolButton("-", "", 0);
        AceoffixCtrl1.AddCustomToolButton("Full-screen Switch", "SwitchFullScreen()", 4);

        AceoffixCtrl1.Bind(wb);
        AceoffixCtrl1.SaveDataPage = "saveexceldata.aspx";
        AceoffixCtrl1.OpenDocument("../doc/xls-submit.xls", Aceoffix.OpenModeType.xlsSubmitForm, "John Scott");
    }
}
