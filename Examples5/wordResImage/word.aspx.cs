using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class wordResImage_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        AceoffixCtrl1.Theme = Aceoffix.ThemeType.Office2007;
        AceoffixCtrl1.BorderStyle = Aceoffix.BorderStyleType.BorderThin;
        AceoffixCtrl1.CustomToolbar = false;
        AceoffixCtrl1.Menubar = false;

        Aceoffix.WordWriter.WordDocument wd = new Aceoffix.WordWriter.WordDocument();
        Aceoffix.WordWriter.DataRegion image1 = wd.OpenDataRegion("Image1");
        image1.Value = "[image]doc/image1.png[/image]";
        Aceoffix.WordWriter.DataRegion image2 = wd.OpenDataRegion("Image2");
        image2.Value = "[image]doc/image2.png[/image]";
        Aceoffix.WordWriter.DataRegion data1 = wd.OpenDataRegion("Paragraph1");
        data1.Value = "[word]doc/paragraph1.doc[/word]";
        Aceoffix.WordWriter.DataRegion data2 = wd.OpenDataRegion("Paragraph2");
        data2.Value = "[word]doc/paragraph2.doc[/word]";
        Aceoffix.WordWriter.DataRegion data3 = wd.OpenDataRegion("Paragraph3");
        data3.Value = "[word]doc/paragraph3.doc[/word]";

        AceoffixCtrl1.Bind(wd);
        AceoffixCtrl1.OpenDocument("doc/merge.doc", Aceoffix.OpenModeType.docNormalEdit, "John Scott");

    }
}