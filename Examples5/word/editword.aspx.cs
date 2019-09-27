using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class word_editword : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    protected void AceoffixCtrl1_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        // Create custom menu
        AceoffixCtrl1.CustomMenuCaption = "My &Tools";
        AceoffixCtrl1.AddCustomMenuItem("Show &caption of titlebar", "OnCustomMenuClick()", true);
        AceoffixCtrl1.AddCustomMenuItem("-", "", true);
        AceoffixCtrl1.AddCustomMenuItem("Save As HTML to Web Server", "OnCustomMenuClick()", true);

        // Create custom toolbar
        AceoffixCtrl1.AddCustomToolButton("Save", "SaveDocument()", 1);
        AceoffixCtrl1.AddCustomToolButton("-", "", 0);
        AceoffixCtrl1.AddCustomToolButton("Page Setup", "ShowPageSetupDlg()", 8);
        AceoffixCtrl1.AddCustomToolButton("Print", "ShowPrintDlg()", 6);
        AceoffixCtrl1.AddCustomToolButton("-", "", 0);
        AceoffixCtrl1.AddCustomToolButton("Save As HTML", "SaveAsHTML()", 0);
        AceoffixCtrl1.AddCustomToolButton("-", "", 0);
        AceoffixCtrl1.AddCustomToolButton("Full-screen Switch", "SwitchFullScreen()", 4);
        AceoffixCtrl1.AddCustomToolButton("-", "", 0);
        AceoffixCtrl1.AddCustomToolButton("Close", "Close()", 9);

        AceoffixCtrl1.SaveFilePage = "savefile.aspx";
        AceoffixCtrl1.OpenDocument("../doc/editword.doc", Aceoffix.OpenModeType.docNormalEdit, "John Scott");
    }
}
