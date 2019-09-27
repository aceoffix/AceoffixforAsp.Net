using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class word_callvba : System.Web.UI.Page
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
        AceoffixCtrl1.AddCustomToolButton("VBA: paragraph format", "CallVBA1()", 0);
        AceoffixCtrl1.AddCustomToolButton("VBA: insert text", "CallVBA2()", 0);
        AceoffixCtrl1.AddCustomToolButton("VBA: set header/footer", "CallVBA3()", 0);

        AceoffixCtrl1.SaveFilePage = "savefile.aspx";
        AceoffixCtrl1.OpenDocument("../doc/editword.doc", Aceoffix.OpenModeType.docNormalEdit, "John Scott");
    }
}
