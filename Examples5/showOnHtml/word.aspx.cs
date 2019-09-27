using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class showOnHtml_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Aceoffix.AceoffixCtrl AceoffixCtrl1 = new Aceoffix.AceoffixCtrl();
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        AceoffixCtrl1.AddCustomToolButton("Save", "SaveDocument()", 1);
        AceoffixCtrl1.Caption = "This content is controlled by Caption property of AceoffixCtrl and can be set as any text you want.";
        AceoffixCtrl1.SaveFilePage = "savefile.aspx";
        AceoffixCtrl1.OpenDocument("doc/introduce.doc", Aceoffix.OpenModeType.docNormalEdit, "John Scott"); // docReadOnly: Disabled Edit, Paste.
        Response.Write(AceoffixCtrl1.GetHtmlCode("AceoffixCtrl1"));
    }
}