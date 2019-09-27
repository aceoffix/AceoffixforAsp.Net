using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;

public partial class createWord_OpenWord : System.Web.UI.Page
{
    string fileName = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        fileName = Request.QueryString["fileName"].Trim();
        Literal_Subject.Text = Request.QueryString["subject"];
    }
    protected void AceoffixCtrl1_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        AceoffixCtrl1.Menubar = false;
        AceoffixCtrl1.SaveFilePage = "SaveFile.aspx";
        AceoffixCtrl1.AddCustomToolButton("Save", "Save()", 1);
        AceoffixCtrl1.AddCustomToolButton("-", "", 0);
        AceoffixCtrl1.AddCustomToolButton("Close", "Close()", 9);

        AceoffixCtrl1.OpenDocument(Server.MapPath("doc/") + fileName, Aceoffix.OpenModeType.docNormalEdit, "Tom");

    }
}