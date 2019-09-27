using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aceoffix.ExcelWriter;
using System.Drawing;

public partial class definedNameCell_excel : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Workbook wb = new Workbook();
        Sheet sheet = wb.OpenSheet("Sheet1");
        sheet.OpenCellByDefinedName("testA").Value = "Sales Report (2015)";
        sheet.OpenCellByDefinedName("testB").Value = "100";
        AceoffixCtrl1.Bind(wb);

        AceoffixCtrl1.AddCustomToolButton("Save", "SaveDocument()", 1);
        AceoffixCtrl1.Menubar = false;
        AceoffixCtrl1.SaveDataPage = "savedata.aspx";
        AceoffixCtrl1.OpenDocument("doc/test.xls", Aceoffix.OpenModeType.xlsNormalEdit, "John Scott");

    }
}