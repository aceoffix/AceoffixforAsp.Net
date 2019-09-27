using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class mergeWord_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Aceoffix.WordWriter.WordDocument doc = new Aceoffix.WordWriter.WordDocument();
        Aceoffix.WordWriter.DataRegion dataRegion1 = doc.OpenDataRegion("Paragraph1");
        dataRegion1.Value = "[word]doc/paragraph1.doc[/word]";
        Aceoffix.WordWriter.DataRegion dataRegion2 = doc.OpenDataRegion("Paragraph2");
        dataRegion2.Value = "[word]doc/paragraph2.doc[/word]";
        Aceoffix.WordWriter.DataRegion dataRegion3 = doc.OpenDataRegion("Paragraph3");
        dataRegion3.Value = "[word]doc/paragraph3.doc[/word]";
        AceoffixCtrl1.Bind(doc);
        AceoffixCtrl1.AddCustomToolButton("Save", "SaveDocument()", 1);
        AceoffixCtrl1.SaveFilePage = "savefile.aspx";
        AceoffixCtrl1.OpenDocument("doc/merge.doc", Aceoffix.OpenModeType.docNormalEdit, "John Scott");
    }
}