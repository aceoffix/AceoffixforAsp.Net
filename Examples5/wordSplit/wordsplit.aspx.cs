using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class wordSplit_wordsplit : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Aceoffix.WordWriter.WordDocument wd = new Aceoffix.WordWriter.WordDocument();
        Aceoffix.WordWriter.DataRegion dataRegion1 = wd.OpenDataRegion("Paragraph1");
        dataRegion1.SubmitAsFile=true;
        dataRegion1.Editing=true;
        Aceoffix.WordWriter.DataRegion dataRegion2 = wd.OpenDataRegion("Paragraph2");
        dataRegion2.SubmitAsFile = true;
        dataRegion2.Editing = true;
        Aceoffix.WordWriter.DataRegion dataRegion3 = wd.OpenDataRegion("Paragraph3");
        dataRegion3.SubmitAsFile = true;
        dataRegion3.Editing = true;

        AceoffixCtrl1.Bind(wd);
        AceoffixCtrl1.AddCustomToolButton("Save", "SaveDocument()", 1);
        AceoffixCtrl1.SaveDataPage = "savedata.aspx";
        AceoffixCtrl1.OpenDocument("doc/split.doc", Aceoffix.OpenModeType.docSubmitForm, "John Scott");

    }
}