using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class convertPDFs_Convert : System.Web.UI.Page
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
        Console.Write(filePath);
        //Set server page
        FileMakerCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        FileMakerCtrl1.SaveFilePage = "SaveFile.aspx";
        FileMakerCtrl1.JsFunction_OnProgressComplete = "OnProgressComplete()";
        FileMakerCtrl1.FillDocumentAsPDF(filePath, Aceoffix.DocumentOpenType.Word, "aa.pdf");
    }
}