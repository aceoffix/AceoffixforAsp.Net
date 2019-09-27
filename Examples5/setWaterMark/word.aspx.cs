﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class setWaterMark_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        AceoffixCtrl1.Theme = Aceoffix.ThemeType.Office2007;
        AceoffixCtrl1.BorderStyle = Aceoffix.BorderStyleType.BorderThin;
        AceoffixCtrl1.CustomToolbar = false;
        AceoffixCtrl1.Menubar = false;

        Aceoffix.WordWriter.WordDocument wd = new Aceoffix.WordWriter.WordDocument();
        wd.WaterMark.Text = "Aceoffix";
        AceoffixCtrl1.Bind(wd);
        AceoffixCtrl1.OpenDocument("doc/introduce.doc", Aceoffix.OpenModeType.docReadOnly, "John Scott");

    }
}