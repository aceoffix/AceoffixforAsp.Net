using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aceoffix.ExcelWriter;

public partial class clickExcelCell_SubmitExcel : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Workbook wb = new Workbook();
        Sheet sheet = wb.OpenSheet("Sheet1");
        Aceoffix.ExcelWriter.Table table = sheet.OpenTable("B4:D8");
        table.ReadOnly = false;
        AceoffixCtrl1.Bind(wb);
        AceoffixCtrl1.JsFunction_OnExcelCellClick = "OnCellClick()";
        AceoffixCtrl1.AddCustomToolButton("Save", "Save", 1);

        AceoffixCtrl1.Menubar = false;
        AceoffixCtrl1.SaveDataPage = "SaveData.aspx";
        AceoffixCtrl1.OpenDocument("doc/test.xls", Aceoffix.OpenModeType.xlsSubmitForm, "John Scott");


    }
}