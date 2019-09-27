using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class wordFillTable2_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Aceoffix.WordWriter.WordDocument doc = new Aceoffix.WordWriter.WordDocument();
        Aceoffix.WordWriter.Table table1 = doc.OpenDataRegion("Table1").OpenTable(1);
        table1.OpenCellRC(1, 1).Value = "Aceoffix";
        int dataRowCount = 5;//Insert the number of rows
        int oldRowCount = 3;//The original number of rows in tables
        for (int j = 0; j < dataRowCount - oldRowCount; j++)
        {
            table1.InsertRowAfter(table1.OpenCellRC(2, 5));
        }
        int i = 1;
        while (i <= dataRowCount)
        {
            table1.OpenCellRC(i, 2).Value="AA" +i.ToString();
            table1.OpenCellRC(i, 3).Value = "BB" + i.ToString();
            table1.OpenCellRC(i, 4).Value = "CC" + i.ToString();
            table1.OpenCellRC(i, 5).Value = "DD" + i.ToString();
            i++;
        }

        AceoffixCtrl1.Bind(doc);
        AceoffixCtrl1.OpenDocument("doc/test.doc", Aceoffix.OpenModeType.docNormalEdit, "John Scott");
    }
}