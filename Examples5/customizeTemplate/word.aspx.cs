using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aceoffix.WordWriter;

public partial class customizeTemplate_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string type = Request.QueryString["type"].ToString();

        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        
        WordDocument doc = new WordDocument();
        doc.EnableAllDataRegionsEditing=true;


        doc.OpenDataRegion("DATE").Value="03/06/2013";
        doc.OpenDataRegion("Payment").Value="First Data";
        doc.OpenDataRegion("Number").Value="************9355";
        doc.OpenDataRegion("Status").Value="Validated (no renewal)";
        doc.OpenDataRegion("Name").Value="Tom";
        doc.OpenDataRegion("Add").Value="XX East XXth Street XXth floor New York NY 10006 USA";
        doc.OpenDataRegion("Tel").Value="+1(888)88882222";
        doc.OpenDataRegion("Email").Value="XXXx@XXXX.com";

        doc.OpenDataTag("{Date}").Value="03/10/2013"; 


        AceoffixCtrl1.Bind(doc);

        if ("1"==type)
        {
            AceoffixCtrl1.AddCustomToolButton("Save", "Save2()", 1);
            AceoffixCtrl1.AddCustomToolButton("-", "", 0);
            AceoffixCtrl1.AddCustomToolButton("Close", "Close()", 9);
            AceoffixCtrl1.OpenDocument("doc/test.doc", Aceoffix.OpenModeType.docNormalEdit, "Tom");
        }
        else if ("2"==type)
        {
            AceoffixCtrl1.Menubar=true;
            AceoffixCtrl1.CustomToolbar=false;
            AceoffixCtrl1.OfficeToolbars=false;
            AceoffixCtrl1.OpenDocument("doc/test.doc", Aceoffix.OpenModeType.docReadOnly, "Tom");
        }
        else if ("3" == type)
        {
            AceoffixCtrl1.AddCustomToolButton("Save", "Save()", 1);
            AceoffixCtrl1.AddCustomToolButton("-", "", 0);
            AceoffixCtrl1.AddCustomToolButton("Close", "Close()", 9);
            AceoffixCtrl1.SaveDataPage="savedata.aspx";
            AceoffixCtrl1.OfficeToolbars=false;
            AceoffixCtrl1.OpenDocument("doc/test.doc", Aceoffix.OpenModeType.docSubmitForm, "Tom");
        }
       
    }
}