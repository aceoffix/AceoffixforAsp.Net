using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class word_editfiledb : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    protected void AceoffixCtrl1_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        // Create custom toolbar
        AceoffixCtrl1.AddCustomToolButton("Save", "SaveDocument()", 1);
        AceoffixCtrl1.AddCustomToolButton("-", "", 0);
        AceoffixCtrl1.AddCustomToolButton("Page Setup", "ShowPageSetupDlg()", 8);
        AceoffixCtrl1.AddCustomToolButton("Print", "ShowPrintDlg()", 6);

        AceoffixCtrl1.SaveFilePage = "savefiledb.aspx?id=1";
        AceoffixCtrl1.OpenDocument("docstream.aspx?id=1", Aceoffix.OpenModeType.docNormalEdit, "John Scott");
    }
}
