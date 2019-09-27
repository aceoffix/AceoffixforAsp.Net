using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Drawing;

public partial class clickDataRegion_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        AceoffixCtrl1.Theme = Aceoffix.ThemeType.Office2007;
        AceoffixCtrl1.BorderStyle = Aceoffix.BorderStyleType.BorderThin;
        Aceoffix.WordWriter.WordDocument doc = new Aceoffix.WordWriter.WordDocument();

        Aceoffix.WordWriter.DataRegion dataReg = doc.OpenDataRegion("test");
        dataReg.Shading.BackgroundPatternColor=Color.Pink;
       
        AceoffixCtrl1.Bind(doc);
        AceoffixCtrl1.AddCustomToolButton("Save", "SaveDocument()", 1);
        AceoffixCtrl1.JsFunction_OnWordDataRegionClick = "OnWordDataRegionClick()";
        AceoffixCtrl1.OfficeToolbars = false;
        AceoffixCtrl1.SaveFilePage = "savefile.aspx";

        AceoffixCtrl1.OpenDocument("doc/test.doc", Aceoffix.OpenModeType.docSubmitForm, "Tom");

    }
}