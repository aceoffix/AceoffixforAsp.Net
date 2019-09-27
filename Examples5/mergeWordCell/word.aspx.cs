using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Drawing;

public partial class mergeWordCell_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Aceoffix.WordWriter.WordDocument doc = new Aceoffix.WordWriter.WordDocument();
        Aceoffix.WordWriter.Table table = doc.OpenDataRegion("Table1").OpenTable(1);

        table.OpenCellRC(1, 1).MergeTo(1,5);
        table.OpenCellRC(1, 1).Value = "Aceoffix";
        table.OpenCellRC(1, 1).Font.Color=Color.Red;
        table.OpenCellRC(1, 1).Font.Size=24;
        table.OpenCellRC(1, 1).Font.Name = "Andalus";
        table.OpenCellRC(1, 1).ParagraphFormat.Alignment = Aceoffix.WordWriter.WdParagraphAlignment.wdAlignParagraphCenter;

        AceoffixCtrl1.Menubar=false;
        AceoffixCtrl1.CustomToolbar=false;
        AceoffixCtrl1.Bind(doc);
        AceoffixCtrl1.SaveDataPage = "savedata.aspx";
        AceoffixCtrl1.OpenDocument("doc/test.doc", Aceoffix.OpenModeType.docSubmitForm, "John Scott");

    }
}