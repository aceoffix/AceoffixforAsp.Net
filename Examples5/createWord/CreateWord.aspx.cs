using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class createWord_CreateWord : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        AceoffixCtrl1.Menubar = false;
        AceoffixCtrl1.SaveFilePage = "SaveNewFile.aspx";
        AceoffixCtrl1.CustomToolbar=false;
        AceoffixCtrl1.JsFunction_BeforeDocumentSaved = "BeforeDocumentSaved()";
        AceoffixCtrl1.CreateNewDocument("Tom", Aceoffix.DocumentVersion.Word2003);
      
    }
}