using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aceoffix.ExcelWriter;

public partial class definedNameTable_ExcelFill2 : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string tempFileName = Request.QueryString["temp"];

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

       
        Cell cellName = sheet.OpenCellByDefinedName("Name");
        cellName.Value="Tom";

        AceoffixCtrl1.Bind(wb);

        AceoffixCtrl1.Menubar = false;

        AceoffixCtrl1.OpenDocument("doc/" + tempFileName, Aceoffix.OpenModeType.xlsNormalEdit, "John Scott");

    }
}