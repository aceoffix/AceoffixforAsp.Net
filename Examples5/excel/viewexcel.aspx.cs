using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class excel_viewexcel : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    protected void AceoffixCtrl1_Load(object sender, EventArgs e)
    {
        Aceoffix.ExcelWriter.Workbook wb = new Aceoffix.ExcelWriter.Workbook();
        wb.DisableSheetSelection = true;
        wb.DisableSheetDoubleClick = true;
        AceoffixCtrl1.Caption = "Work Mode: Read Only; [Disabled Edit, Copy, Paste, Save as, Download, PrintScreen, F12, Context Menu and etc.]";
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        // Create custom toolbar
        AceoffixCtrl1.AddCustomToolButton("Page Setup", "ShowPageSetup()", 0);
        AceoffixCtrl1.AddCustomToolButton("Print", "ShowPrintDlg()", 6);
        AceoffixCtrl1.AddCustomToolButton("-", "", 0);
        AceoffixCtrl1.AddCustomToolButton("Full-screen Switch", "SwitchFullScreen()", 4);

        AceoffixCtrl1.AllowCopy = false; // Disabled Copy, PrintScreen, F12, Context Menu and etc.
        AceoffixCtrl1.Bind(wb);
        AceoffixCtrl1.OpenDocument("../doc/viewexcel.xls", Aceoffix.OpenModeType.xlsReadOnly, "John Scott"); // docReadOnly: Disabled Edit, Paste.
    }
}
