using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class examinationPaper_Compose : System.Web.UI.Page
{

    protected string operateStr = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        int pNum = 1;

        operateStr += "function Create(){\n";

        operateStr += "var obj = document.getElementById('AceoffixCtrl1').DocumentObject.Application;\n";
        operateStr += "obj.Selection.EndKey(6);\n"; 

        for (int i = 10; i > 0; i--)
        {
            string a = "on";
            string c = Request.Form["check" + i.ToString()];

            if (a.Equals(c))
            {
                operateStr += "obj.Selection.TypeParagraph();"; 
                operateStr += "obj.Selection.Range.Text = '" + pNum.ToString() + ".';\n"; 
                
                operateStr += "obj.Selection.EndKey(5,1);\n";
                operateStr += "obj.Selection.MoveRight(1,1);\n";

                operateStr += "document.getElementById('AceoffixCtrl1').InsertDocumentFromURL('Openfile.aspx?id=" + i + "');\n";
                pNum++;

            }
        }
        operateStr += "\n}\n";

        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        AceoffixCtrl1.Menubar = false;
        AceoffixCtrl1.CustomToolbar = false;
        AceoffixCtrl1.JsFunction_AfterDocumentOpened = "Create()";

        AceoffixCtrl1.OpenDocument("doc/test.doc", Aceoffix.OpenModeType.docNormalEdit, "John Scott");
	
    }
}