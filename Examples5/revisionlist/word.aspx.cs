﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class revisionlist_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        AceoffixCtrl1.JsFunction_AfterDocumentOpened = "AfterDocumentOpened()";
        
        AceoffixCtrl1.AddCustomToolButton("Save", "SaveDocument()", 1);
        AceoffixCtrl1.SaveFilePage = "savefile.aspx";

        AceoffixCtrl1.OfficeToolbars = false;
        AceoffixCtrl1.Menubar = false;

        AceoffixCtrl1.OpenDocument("doc/introduce.doc", Aceoffix.OpenModeType.docForcedRevision, "Tom");

    }
}