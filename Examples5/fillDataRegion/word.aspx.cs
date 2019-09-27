using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Drawing;

public partial class fillDataRegion_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        AceoffixCtrl1.Theme = Aceoffix.ThemeType.Office2007;
        AceoffixCtrl1.BorderStyle = Aceoffix.BorderStyleType.BorderThin;
        Aceoffix.WordWriter.WordDocument doc = new Aceoffix.WordWriter.WordDocument();

        Aceoffix.WordWriter.DataRegion data1 = doc.OpenDataRegion("Name");
        data1.Value = "Tom";
        data1.Font.Color = Color.Blue;
        data1.Font.Size = 24;
        data1.Font.Name = "Berlin Sans FB Demi";
        data1.Font.Bold = true;

        Aceoffix.WordWriter.DataRegion data2 = doc.OpenDataRegion("Department");
        data2.Value = "Sales Department";
        data2.Font.Color = Color.Red;

        AceoffixCtrl1.Bind(doc);
        AceoffixCtrl1.CustomToolbar=false;
        AceoffixCtrl1.Menubar = false;

        AceoffixCtrl1.OpenDocument("doc/test.doc", Aceoffix.OpenModeType.docNormalEdit, "Tom");

    }
}