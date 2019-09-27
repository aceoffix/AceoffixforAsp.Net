using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class openFromDisk_word : System.Web.UI.Page
{
    string filePath = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        filePath = Server.MapPath("doc/introduce.doc");
        Literal_Info.Text = "The value of filePath is" + filePath;
       }

    protected void Aceoffix1_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        AceoffixCtrl1.AddCustomToolButton("Save", "SaveDocument()", 1);
        AceoffixCtrl1.SaveFilePage = "savefile.aspx";
        AceoffixCtrl1.OpenDocument(filePath, Aceoffix.OpenModeType.docNormalEdit, "John Scott");



    }

}