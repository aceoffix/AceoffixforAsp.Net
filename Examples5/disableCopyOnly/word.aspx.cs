﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class disableCopyOnly_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    protected void AceoffixCtrl1_Load(object sender, EventArgs e)
    {
         AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
         AceoffixCtrl1.JsFunction_AfterDocumentOpened = "OnAfterDocumentOpened()";
         AceoffixCtrl1.DisableCopyOnly = true;
        AceoffixCtrl1.SaveFilePage = "savefile.aspx";
        AceoffixCtrl1.OpenDocument("doc/introduce.doc", Aceoffix.OpenModeType.docNormalEdit, "John Scott"); // docReadOnly: Disabled Edit, Paste.

    }
}