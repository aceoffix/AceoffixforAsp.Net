using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class word_editableregions : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    protected void AceoffixCtrl1_Load(object sender, EventArgs e)
    {
        string strUserName = Request.QueryString["user"];
        
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        // Create custom toolbar
        AceoffixCtrl1.AddCustomToolButton("Save", "SaveDocument()", 1);
        AceoffixCtrl1.AddCustomToolButton("-", "", 0);
        AceoffixCtrl1.AddCustomToolButton("Print", "ShowPrintDlg()", 6);

        if ((strUserName != null) && (strUserName.Length > 0))
        {
            //In Word document, developer can define any range of selection as a DataRegion object by inserting a bookmark with the prefix name "ACE_". 
            //In other words, the DataRegion is a bookmark object whose name begins with "ACE_" in Word document.
            Aceoffix.WordWriter.WordDocument wd = new Aceoffix.WordWriter.WordDocument();
            if (strUserName == "John")
            {
                wd.OpenDataRegion("title").Editing = true; // ACE_title
                wd.OpenDataRegion("firstpara").Editing = true; // ACE_firstpara
            }
            else if (strUserName == "Mike")
            {
                wd.OpenDataRegion("func1").Editing = true; // ACE_func1
                wd.OpenDataRegion("func2").Editing = true; // ACE_func2
                wd.OpenDataRegion("lastpara").Editing = true; // ACE_lastpara
            }
            AceoffixCtrl1.Bind(wd);
            AceoffixCtrl1.SaveFilePage = "savefile.aspx";
            AceoffixCtrl1.OpenDocument("../doc/editableregions.doc", Aceoffix.OpenModeType.docSubmitForm, strUserName);
        }
    }
}
