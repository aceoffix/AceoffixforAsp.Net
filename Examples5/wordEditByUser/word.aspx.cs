using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class wordEditByUser_word : System.Web.UI.Page
{
    protected string userName = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        userName = Request.QueryString["userName"];

        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        AceoffixCtrl1.Theme = Aceoffix.ThemeType.Office2007;
        AceoffixCtrl1.BorderStyle = Aceoffix.BorderStyleType.BorderThin;


        Aceoffix.WordWriter.WordDocument doc = new Aceoffix.WordWriter.WordDocument();

        Aceoffix.WordWriter.DataRegion data1 = doc.OpenDataRegion("Paragraph1");
        data1.Value = "[word]doc/paragraph1.doc[/word]";
        data1.SubmitAsFile = true;
        Aceoffix.WordWriter.DataRegion data2 = doc.OpenDataRegion("Paragraph2");
        data2.Value = "[word]doc/paragraph2.doc[/word]";
        data2.SubmitAsFile = true;
        Aceoffix.WordWriter.DataRegion data3 = doc.OpenDataRegion("Paragraph3");
        data3.Value = "[word]doc/paragraph3.doc[/word]";

        if (userName.Equals("Tom"))
        {
            data1.Editing=true;
            data2.Editing=false;
            data3.Editing=false;
        }
        else
        {
            data1.Editing=false;
            data2.Editing = true;
            data3.Editing=false;
        }
       
        AceoffixCtrl1.Bind(doc);
        AceoffixCtrl1.AddCustomToolButton("Save", "SaveDocument()", 1);

        AceoffixCtrl1.SaveDataPage = "savedata.aspx?userName="+userName;
        AceoffixCtrl1.OpenDocument("doc/merge.doc", Aceoffix.OpenModeType.docSubmitForm, userName);

    }
}