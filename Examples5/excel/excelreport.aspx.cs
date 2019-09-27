using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class excel_excelreport : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    protected void AceoffixCtrl1_Load(object sender, EventArgs e)
    {
        Aceoffix.ExcelWriter.Workbook wb = new Aceoffix.ExcelWriter.Workbook();
        Aceoffix.ExcelWriter.Sheet sheet1 = wb.OpenSheet("Sheet1");
        Aceoffix.ExcelWriter.Table table1 = sheet1.OpenTable("C5:E16");

        //Output data into the table.
        System.Random rd = new Random(System.DateTime.Now.Millisecond);
        for (int j = 0; j < 12; j++)
        {
            for (int i = 0; i < table1.DataFields.Count; i++)
            {
                int iValue = (rd.Next(25) + 10) * 1000;
                table1.DataFields[i].Value = iValue.ToString(); // You can set the data from the database.
                if (iValue > 30 * 1000) // If the value is greater than 30*1000, then the cell will display alert style to warn the user.
                {
                    table1.DataFields[i].BackColor = System.Drawing.Color.Red;
                    table1.DataFields[i].ForeColor = System.Drawing.Color.Yellow;
                }
            }
            table1.NextRow();
        }
        table1.Close();

        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        AceoffixCtrl1.Caption = "Subject: Usage of Excel Table [Double click the title bar to enter/exit full screen]";
        AceoffixCtrl1.Bind(wb);

        // Create custom toolbar
        AceoffixCtrl1.AddCustomToolButton("Save As", "SaveAsDocument()", 1);
        AceoffixCtrl1.AddCustomToolButton("-", "", 0);
        AceoffixCtrl1.AddCustomToolButton("Page Setup", "ShowPageSetupDlg()", 0);
        AceoffixCtrl1.AddCustomToolButton("Print", "ShowPrintDlg()", 6);
        AceoffixCtrl1.AddCustomToolButton("-", "", 0);
        AceoffixCtrl1.AddCustomToolButton("Full-screen Switch", "SwitchFullScreen()", 4);

        AceoffixCtrl1.FileTitle = "Report1";
        AceoffixCtrl1.OpenDocument("../doc/xls-report.xls", Aceoffix.OpenModeType.xlsReadOnly, "John Scott");
    }
}
