using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class convertPDFs_Edit : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        String id = Request.QueryString["id"].Trim();
        String filePath = "";
        if ("1".Equals(id))
        {
            filePath = Server.MapPath("doc/AceoffixIntroduction.docx");
        }
        if ("2".Equals(id))
        {
            filePath = Server.MapPath("doc/UsingAceoffixtrialversion.docx");
        }
        if ("3".Equals(id))
        {
            filePath = Server.MapPath("doc/HowtoIntegrateAceoffix.docx");
        }
        if ("4".Equals(id))
        {
            filePath = Server.MapPath("doc/HowtoSupportallBrowsers.docx");
        }


        // Set Aceoffix control server page
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        // Set save file page
        AceoffixCtrl1.SaveFilePage = "SaveFile.aspx";
        // Add a customized button for saving
        AceoffixCtrl1.AddCustomToolButton("Save","Save()",1);
        // Open the document
        AceoffixCtrl1.OpenDocument(filePath,Aceoffix.OpenModeType.docNormalEdit,"Tom");
    }
}