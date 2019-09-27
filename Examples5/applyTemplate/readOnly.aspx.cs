using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class applyTemplate_readOnly : System.Web.UI.Page
{
    protected string fileName = "Official.doc";
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        AceoffixCtrl1.Caption = fileName;

        AceoffixCtrl1.AddCustomToolButton("Save as", "ShowDialog1()", 5);
        AceoffixCtrl1.AddCustomToolButton("Full-screen Switch", "IsFullScreen", 4);
        AceoffixCtrl1.AddCustomToolButton("Print", "ShowDialog3()", 6);
        AceoffixCtrl1.AddCustomToolButton("Page Setup", "ShowDialog2()", 0);
        AceoffixCtrl1.Menubar = false;
        AceoffixCtrl1.OfficeToolbars = false;

        try
        {
            AceoffixCtrl1.OpenDocument(Server.MapPath("doc/" + fileName), Aceoffix.OpenModeType.docReadOnly, "John Scott");
        }
        catch
        {
            throw new Exception("The offical document has not been released. Please go back and click \"Apply Template\" link and then save the document.");
        }

    }
}