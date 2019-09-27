using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class word_revisiononly : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    protected void AceoffixCtrl1_Load(object sender, EventArgs e)
    {
        string strUserName = Request.QueryString["user"];
        
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        // Create custom toolbar
        AceoffixCtrl1.AddCustomToolButton("Save", "SaveDocument()", 1);
        AceoffixCtrl1.AddCustomToolButton("-", "", 0);
        AceoffixCtrl1.AddCustomToolButton("Show/Hide Revisions", "ShowRevisions()", 0);
        AceoffixCtrl1.AddCustomToolButton("Print", "ShowPrintDlg()", 6);

        if ((strUserName != null) && (strUserName.Length > 0))
        {
            AceoffixCtrl1.SaveFilePage = "savefile.aspx";
            AceoffixCtrl1.OpenDocument("../doc/wordrevision.doc", Aceoffix.OpenModeType.docForcedRevision, strUserName);
        }
    }
}
