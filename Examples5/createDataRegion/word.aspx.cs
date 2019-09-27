using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class createDataRegion_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Aceoffix.WordWriter.WordDocument doc = new Aceoffix.WordWriter.WordDocument();
        Aceoffix.WordWriter.DataRegion dataRegion1 = doc.CreateDataRegion("Paragraph1", Aceoffix.WordWriter.DataRegionInsertType.After, "[home]");
        dataRegion1.Value = "[word]doc/paragraph1.doc[/word]\r\n";
        Aceoffix.WordWriter.DataRegion dataRegion2 = doc.CreateDataRegion("Paragraph2", Aceoffix.WordWriter.DataRegionInsertType.After, "Paragraph1");
        dataRegion2.Value = "[word]doc/paragraph2.doc[/word]\r\n";
        Aceoffix.WordWriter.DataRegion dataRegion3 = doc.CreateDataRegion("Paragraph3", Aceoffix.WordWriter.DataRegionInsertType.After, "Paragraph2");
        dataRegion3.Value = "[word]doc/paragraph3.doc[/word]\r\n";
        AceoffixCtrl1.Bind(doc);
        AceoffixCtrl1.AddCustomToolButton("Save", "SaveDocument()", 1);
        AceoffixCtrl1.SaveFilePage = "savefile.aspx";
        AceoffixCtrl1.OpenDocument("doc/merge.doc", Aceoffix.OpenModeType.docNormalEdit, "John Scott");

    }
}