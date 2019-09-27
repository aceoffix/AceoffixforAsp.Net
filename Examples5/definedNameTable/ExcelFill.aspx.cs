using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aceoffix.ExcelWriter;

public partial class definedNameTable_ExcelFill : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Workbook wb = new Workbook();
        Sheet sheet = wb.OpenSheet("Sheet1");

        Aceoffix.ExcelWriter.Table table = sheet.OpenTableByDefinedName("report", 10, 5, false);
        table.DataFields[0].Value = "Jan";
        table.DataFields[1].Value = "100";
        table.DataFields[2].Value = "120";
        table.DataFields[3].Value = "500";
        table.DataFields[4].Value = "120%";
        table.NextRow();
        table.Close();

        AceoffixCtrl1.Bind(wb);

        AceoffixCtrl1.AddCustomToolButton("Save", "Save()", 1);
        AceoffixCtrl1.Menubar = false;
        AceoffixCtrl1.SaveDataPage = "SaveData.aspx";
        AceoffixCtrl1.OpenDocument("doc/test.xls", Aceoffix.OpenModeType.xlsSubmitForm, "John Scott");


    }
}