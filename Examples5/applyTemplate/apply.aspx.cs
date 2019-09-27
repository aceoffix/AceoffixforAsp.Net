using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using Aceoffix.WordWriter;

public partial class applyTemplate_apply : System.Web.UI.Page
{
    public string mb = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        if (!IsPostBack)
        {
            string fileName = "test.doc";
            OpenWord(fileName);
        }
    }
    private void OpenWord(string fileName)
    {
        
        Session.Add("fileName", fileName);
        AceoffixCtrl1.CustomToolbar = false;
        AceoffixCtrl1.SaveFilePage = "savefile.aspx";
        AceoffixCtrl1.OpenDocument("doc/" + fileName, Aceoffix.OpenModeType.docNormalEdit, "Tom");
    }
    protected void BtnApply_Click(object sender, EventArgs e)
    {
        string fileName = "Official.doc";
        mb = Request.Form["templateName"];
        System.IO.File.Copy(Server.MapPath("doc/" + mb),
                Server.MapPath("doc/" + fileName), true);

        WordDocument doc = new WordDocument();
        DataRegion title = doc.OpenDataRegion("Title");
        title.Value="Invitation Letter";
        DataRegion information = doc.OpenDataRegion("Information");
        information.Value="Cooperative Extension is a publicly funded, non-formal educational partnership of the U.S. Department of Agriculture, land-grant universities, and county office.The USDA is an equal opportunity provider and employer.";
        DataRegion text = doc.OpenDataRegion("Text");
        text.Value="[word]doc/test.doc[/word]";
        
        AceoffixCtrl1.Bind(doc);

        OpenWord(fileName);
    }
}