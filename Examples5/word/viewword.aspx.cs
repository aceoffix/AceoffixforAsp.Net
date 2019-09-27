using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class word_viewword : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    protected void AceoffixCtrl1_Load(object sender, EventArgs e)
    {
        Aceoffix.WordWriter.WordDocument wd = new Aceoffix.WordWriter.WordDocument();
        wd.DisableWindowSelection = true; // Disabled Selection and Drag.
        AceoffixCtrl1.Caption = "Work Mode: Read Only; [Disabled Edit, Copy, Paste, Save as, Download, PrintScreen, F12, Context Menu, Print and etc.]";
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        AceoffixCtrl1.JsFunction_AfterDocumentOpened = "OnAfterDocumentOpened()";
        AceoffixCtrl1.AllowCopy = false; // Disabled Copy, PrintScreen, F12, Context Menu and etc.
        AceoffixCtrl1.Bind(wd);
        AceoffixCtrl1.OpenDocument("../doc/viewword.doc", Aceoffix.OpenModeType.docReadOnly, "John Scott"); // docReadOnly: Disabled Edit, Paste.
    }
}
