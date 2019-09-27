using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class wordDataTag_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Aceoffix.WordWriter.WordDocument wd = new Aceoffix.WordWriter.WordDocument();
        Aceoffix.WordWriter.DataTag dataTag1 = wd.OpenDataTag("{name}");
        dataTag1.Value = "Tom";
        Aceoffix.WordWriter.DataTag dataTag2 = wd.OpenDataTag("{Number}");
        dataTag2.Value = "201501";
        Aceoffix.WordWriter.DataTag dataTag3 = wd.OpenDataTag("{Department}");
        dataTag3.Value = "Development";
        Aceoffix.WordWriter.DataTag dataTag4 = wd.OpenDataTag("{Manager}");
        dataTag4.Value = "John Scott";
        Aceoffix.WordWriter.DataTag dataTag5 = wd.OpenDataTag("{Salary}");
        dataTag5.Value = "$5000";

        AceoffixCtrl1.Bind(wd);
        AceoffixCtrl1.AddCustomToolButton("Save", "SaveDocument()", 1);
        AceoffixCtrl1.SaveFilePage = "savefile.aspx";
        AceoffixCtrl1.OpenDocument("doc/test.doc", Aceoffix.OpenModeType.docNormalEdit, "John Scott");

    }
}