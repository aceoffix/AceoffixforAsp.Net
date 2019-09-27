using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class AceoffixLink_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        AceoffixCtrl1.SaveFilePage = "SaveFile.aspx";
        AceoffixCtrl1.AddCustomToolButton("Save", "Save()", 1);
        AceoffixCtrl1.AddCustomToolButton("SaveAsPDF", "SaveAsPDF()", 1);
        string fileName = "editword.doc";
        string pdfName = fileName.Substring(0, fileName.Length - 4) + ".pdf";
        AceoffixCtrl1.OpenDocument("doc/" + fileName, Aceoffix.OpenModeType.docNormalEdit, "Tom");
    }
}