using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class extractTable_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Aceoffix.WordWriter.WordDocument doc = new Aceoffix.WordWriter.WordDocument();
        Aceoffix.WordWriter.DataRegion dTable = doc.OpenDataRegion("table");
        dTable.Editing = true;
        Aceoffix.WordWriter.Table table1 = doc.OpenDataRegion("Table").OpenTable(1);

        table1.OpenCellRC(1, 2).Value = "Product 1";
        table1.OpenCellRC(1, 3).Value = "Product 2";
        table1.OpenCellRC(2, 1).Value = "A";
        table1.OpenCellRC(3, 1).Value = "B";

        AceoffixCtrl1.AddCustomToolButton("Save", "Save", 1);
        AceoffixCtrl1.AddCustomToolButton("Full-screen Switch", "IsFullScreen", 4);
        AceoffixCtrl1.Bind(doc);
        AceoffixCtrl1.SaveDataPage = "savedata.aspx";
        AceoffixCtrl1.OpenDocument("doc/test.doc", Aceoffix.OpenModeType.docSubmitForm, "John Scott");


    }
}