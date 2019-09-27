using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class saveDataAndFile_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Aceoffix.WordWriter.WordDocument wd = new Aceoffix.WordWriter.WordDocument();
        Aceoffix.WordWriter.DataRegion dataRegion1 = wd.OpenDataRegion("Name");
        dataRegion1.Editing = true;
        dataRegion1.Value = "Tom";
        Aceoffix.WordWriter.DataRegion dataRegion3 = wd.OpenDataRegion("Department");
        dataRegion3.Editing = true;
        dataRegion3.Value = "Development";
        

        AceoffixCtrl1.Bind(wd);
        AceoffixCtrl1.AddCustomToolButton("Save", "SaveDocument()", 1);
        AceoffixCtrl1.SaveDataPage = "savedata.aspx"; //Key code in this demo.
        AceoffixCtrl1.SaveFilePage = "savefile.aspx"; //Key code in this demo.

        AceoffixCtrl1.OpenDocument("doc/test.doc", Aceoffix.OpenModeType.docSubmitForm, "John Scott");

    }
}