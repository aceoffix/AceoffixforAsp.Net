using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aceoffix.ExcelWriter;

public partial class definedNameTable_ExcelFill4 : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Workbook wb = new Workbook();
        Sheet sheet = wb.OpenSheet("Sheet1");

        Aceoffix.ExcelWriter.Table table = sheet.OpenTable("B4:F11");
        int rowCount = 12;
        for (int i = 1; i <= rowCount; i++)
        {
            table.DataFields[0].Value = i + "Month";
            table.DataFields[1].Value = "100";
            table.DataFields[2].Value = "120";
            table.DataFields[3].Value = "500";
            table.DataFields[4].Value = "120%";
            table.NextRow();

        }
       
        table.Close();

        Aceoffix.ExcelWriter.Table table2 = sheet.OpenTable("B13:F16");
        int rowCount2 = 12;
        for (int i = 1; i <= rowCount2; i++)
        {
            table2.DataFields[0].Value = i + "Quarter";
            table2.DataFields[1].Value = "300";
            table2.DataFields[2].Value = "300";
            table2.DataFields[3].Value = "300";
            table2.DataFields[4].Value = "100%";
            table2.NextRow();

        }


        table2.Close();
        
        AceoffixCtrl1.Bind(wb);

        AceoffixCtrl1.Menubar = false;

        AceoffixCtrl1.OpenDocument("doc/test4.xls", Aceoffix.OpenModeType.xlsNormalEdit, "John Scott");

    }
}