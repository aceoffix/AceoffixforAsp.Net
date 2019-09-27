using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class word_wordsubmit : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    protected void AceoffixCtrl1_Load(object sender, EventArgs e)
    {
        Aceoffix.WordWriter.WordDocument wd = new Aceoffix.WordWriter.WordDocument();
        wd.OpenDataRegion("EmployeeName").Editing = true;
        wd.OpenDataRegion("EmployeeNumber").Editing = true;
        wd.OpenDataRegion("Department").Editing = true;
        wd.OpenDataRegion("Manager").Editing = true;

        wd.OpenDataRegion("FromDate").Value = DateTime.Now.AddDays(2).ToLongDateString();
        wd.OpenDataRegion("FromDate").Editing = true;
        wd.OpenDataRegion("ToDate").Value = DateTime.Now.AddDays(2).ToLongDateString();
        wd.OpenDataRegion("ToDate").Editing = true;

        wd.OpenDataRegion("Reason").Editing = true;

        AceoffixCtrl1.Caption = "Word document used as an input form.";
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        // Create custom toolbar
        AceoffixCtrl1.AddCustomToolButton("Save", "SaveDocument()", 1);
        AceoffixCtrl1.AddCustomToolButton("-", "", 0);
        AceoffixCtrl1.AddCustomToolButton("Print", "ShowPrintDlg()", 6);
        AceoffixCtrl1.AddCustomToolButton("-", "", 0);
        AceoffixCtrl1.AddCustomToolButton("Full-screen Switch", "SwitchFullScreen()", 4);

        AceoffixCtrl1.Bind(wd);
        AceoffixCtrl1.SaveDataPage = "saveworddata.aspx";
        AceoffixCtrl1.OpenDocument("../doc/doc-submit.doc", Aceoffix.OpenModeType.docSubmitForm, "John Scott"); 
    }
}
