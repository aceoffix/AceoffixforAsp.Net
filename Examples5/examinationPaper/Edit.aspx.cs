using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class examinationPaper_Edit : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        AceoffixCtrl1.Menubar = false;
        AceoffixCtrl1.AddCustomToolButton("Save", "SaveDocument()", 1);
        string id = Request.QueryString["id"];
        AceoffixCtrl1.SaveFilePage = "SaveFile.aspx?id=" + id;

        AceoffixCtrl1.OpenDocument("Openfile.aspx?id=" + id, Aceoffix.OpenModeType.docNormalEdit, "John Scott");
    }
}