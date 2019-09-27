using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class defineDataRegion_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        AceoffixCtrl1.Theme = Aceoffix.ThemeType.Office2007;
        AceoffixCtrl1.BorderStyle = Aceoffix.BorderStyleType.BorderThin;
        Aceoffix.WordWriter.WordDocument doc = new Aceoffix.WordWriter.WordDocument();

        doc.Template.DefineDataRegion("Date", "[Date]");
        doc.Template.DefineDataRegion("Payment", "[Payment method]");
        doc.Template.DefineDataRegion("Number", "[Credit card number]");
        doc.Template.DefineDataRegion("Status", "[Status]");
        doc.Template.DefineDataRegion("Name", "[Customer name]");
        doc.Template.DefineDataRegion("Add", "[Add]");
        doc.Template.DefineDataRegion("Tel", "[Tel]");
        doc.Template.DefineDataRegion("Email", "[Email]");


        AceoffixCtrl1.Bind(doc);
        AceoffixCtrl1.AddCustomToolButton("Save", "Save()", 1);
        AceoffixCtrl1.AddCustomToolButton("Define DataRegion", "ShowDefineDataRegions()", 3);
        AceoffixCtrl1.SaveFilePage = "savefile.aspx";

        AceoffixCtrl1.OpenDocument("doc/test.doc", Aceoffix.OpenModeType.docNormalEdit, "Tom");

    }
}