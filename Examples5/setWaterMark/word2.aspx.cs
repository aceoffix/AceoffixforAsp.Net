using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class setWaterMark_word2 : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        AceoffixCtrl1.Theme = Aceoffix.ThemeType.Office2007;
        AceoffixCtrl1.BorderStyle = Aceoffix.BorderStyleType.BorderThin;
        AceoffixCtrl1.Menubar = false;
        AceoffixCtrl1.AddCustomToolButton("Set WaterMark by JS", "AddWaterM", 0);

        AceoffixCtrl1.OpenDocument("doc/introduce.doc", Aceoffix.OpenModeType.docReadOnly, "John");
    }
}