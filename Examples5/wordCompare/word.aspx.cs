using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class wordCompare_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        AceoffixCtrl1.Theme = Aceoffix.ThemeType.Office2007;
        AceoffixCtrl1.BorderStyle = Aceoffix.BorderStyleType.BorderThin;

        AceoffixCtrl1.AddCustomToolButton("Show Document A", "ShowFile1View()", 0);
        AceoffixCtrl1.AddCustomToolButton("Show Document B", "ShowFile2View()", 0);
        AceoffixCtrl1.AddCustomToolButton("Show comparison results", "ShowCompareView()", 0);
        AceoffixCtrl1.Menubar = false;

        AceoffixCtrl1.WordCompare("doc/introduce1.doc", "doc/introduce2.doc", Aceoffix.OpenModeType.docFinalizeDraft, "Tom");

    }
}