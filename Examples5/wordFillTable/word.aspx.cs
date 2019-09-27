using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class wordFillTable_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Aceoffix.WordWriter.WordDocument wd = new Aceoffix.WordWriter.WordDocument();


        Aceoffix.WordWriter.DataRegion dataRegion = wd.OpenDataRegion("ACE_Table");

        Aceoffix.WordWriter.Table table = dataRegion.OpenTable(1);

        table.OpenCellRC(3, 1).Value="Tom";
        table.OpenCellRC(3, 2).Value="201501";
        table.OpenCellRC(3, 3).Value="Development";
        table.OpenCellRC(3, 4).Value="John Scott";
        table.OpenCellRC(3, 5).Value="$5000";

        table.InsertRowAfter(table.OpenCellRC(3, 5));

        table.OpenCellRC(4, 1).Value="Jack";
        table.OpenCellRC(4, 2).Value="201502";
        table.OpenCellRC(4, 3).Value="Sales";
        table.OpenCellRC(4, 4).Value="Anna";
        table.OpenCellRC(4, 5).Value = "$5500";

        AceoffixCtrl1.Bind(wd);
        AceoffixCtrl1.OpenDocument("doc/test.doc", Aceoffix.OpenModeType.docNormalEdit, "John Scott");

    }
}