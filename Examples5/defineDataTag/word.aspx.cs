using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class defineDataTag_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        AceoffixCtrl1.Theme = Aceoffix.ThemeType.Office2007;
        AceoffixCtrl1.BorderStyle = Aceoffix.BorderStyleType.BorderThin;
        Aceoffix.WordWriter.WordDocument doc = new Aceoffix.WordWriter.WordDocument();

        doc.Template.DefineDataTag("[CONTRACT NO]");
        doc.Template.DefineDataTag("[DATA]");
        doc.Template.DefineDataTag("[THE BUYER]");
        doc.Template.DefineDataTag("[Add]");
        doc.Template.DefineDataTag("{Add}");
        doc.Template.DefineDataTag("[Tel]");
        doc.Template.DefineDataTag("[THE SELLER]");
        doc.Template.DefineDataTag("[Name]");
        doc.Template.DefineDataTag("[Signature]");
        doc.Template.DefineDataTag("{Date}");

        AceoffixCtrl1.Bind(doc);
        AceoffixCtrl1.AddCustomToolButton("Save", "SaveDocument()", 1);
        AceoffixCtrl1.AddCustomToolButton("Define DataTags", "ShowDefineDataTags()", 3);
        AceoffixCtrl1.SaveFilePage = "savefile.aspx";

        AceoffixCtrl1.OpenDocument("doc/test.doc", Aceoffix.OpenModeType.docNormalEdit, "Tom");

    }
}