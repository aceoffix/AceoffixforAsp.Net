using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class excelTable_excel : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Aceoffix.ExcelWriter.Workbook wb = new Aceoffix.ExcelWriter.Workbook();
        Aceoffix.ExcelWriter.Sheet sheet = wb.OpenSheet("Sheet1");
        Aceoffix.ExcelWriter.Table table = sheet.OpenTable("B4:F13");

        for (int i = 0; i < 50; i++)
        {
            table.DataFields[0].Value="Product:" + i.ToString();
            table.DataFields[1].Value="100";
            table.DataFields[2].Value=(100 + i).ToString();
            table.NextRow();
        }
        table.Close();


        sheet.OpenCell("F4").Value = string.Format("{0:P}", 270.0 / 300);
        AceoffixCtrl1.Bind(wb);

        AceoffixCtrl1.Menubar = false;
        AceoffixCtrl1.CustomToolbar = false;
        AceoffixCtrl1.OpenDocument("doc/test.xls", Aceoffix.OpenModeType.xlsNormalEdit, "John Scott");

    }
}