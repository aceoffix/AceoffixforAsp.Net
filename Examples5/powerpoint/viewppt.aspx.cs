using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class powerpoint_viewppt : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    protected void AceoffixCtrl1_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.Caption = "Work Mode: Read Only; [Disabled Edit, Copy, Paste, Save as, Download, PrintScreen, F12, Context Menu and etc.]";
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        AceoffixCtrl1.AllowCopy = false; // Disabled Copy, PrintScreen, F12, Context Menu and etc.
        AceoffixCtrl1.OpenDocument("../doc/viewppt.ppt", Aceoffix.OpenModeType.pptReadOnly, "John Scott"); // docReadOnly: Disabled Edit, Paste.
    }
}
