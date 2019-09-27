using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class wordTextBox_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        AceoffixCtrl1.Theme = Aceoffix.ThemeType.Office2007;
        AceoffixCtrl1.BorderStyle = Aceoffix.BorderStyleType.BorderThin;
        Aceoffix.WordWriter.WordDocument doc = new Aceoffix.WordWriter.WordDocument();

        Aceoffix.WordWriter.DataRegion data1 = doc.OpenDataRegion("Text1");
        data1.Value = "Aceoffix";

        Aceoffix.WordWriter.DataRegion data2 = doc.OpenDataRegion("Text2");
        data2.Value = "[image]doc/image.png[/image]";
        
        AceoffixCtrl1.Bind(doc);
        AceoffixCtrl1.CustomToolbar = false;
        AceoffixCtrl1.Menubar = false;

        AceoffixCtrl1.OpenDocument("doc/test.doc", Aceoffix.OpenModeType.docNormalEdit, "Tom");


    }
}