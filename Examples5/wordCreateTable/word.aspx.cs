using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class wordCreateTable_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        AceoffixCtrl1.Theme = Aceoffix.ThemeType.Office2007;
        AceoffixCtrl1.BorderStyle = Aceoffix.BorderStyleType.BorderThin;
        Aceoffix.WordWriter.WordDocument doc = new Aceoffix.WordWriter.WordDocument();

        Aceoffix.WordWriter.Table table1 = doc.OpenDataRegion("Table1").CreateTable(3,5,Aceoffix.WordWriter.WdAutoFitBehavior.wdAutoFitWindow);
        table1.OpenCellRC(1, 1).MergeTo(3,1);
        table1.OpenCellRC(1, 1).Value = "The merged cells";
        for (int i = 1; i < 4; i++)
        {
            table1.OpenCellRC(i, 2).Value = "AA" + i.ToString();
            table1.OpenCellRC(i, 3).Value = "BB" + i.ToString();
            table1.OpenCellRC(i, 4).Value = "CC" + i.ToString();
            table1.OpenCellRC(i, 5).Value = "DD" + i.ToString();
        }
        Aceoffix.WordWriter.DataRegion drTable2 = doc.CreateDataRegion("Table2", Aceoffix.WordWriter.DataRegionInsertType.After, "Table1");
        Aceoffix.WordWriter.Table table2 = doc.OpenDataRegion("Table2").CreateTable(5, 5, Aceoffix.WordWriter.WdAutoFitBehavior.wdAutoFitWindow);
        for (int i = 1; i < 6; i++)
        {
            table2.OpenCellRC(i, 1).Value = "AA" + i.ToString();
            table2.OpenCellRC(i, 2).Value = "BB" + i.ToString();
            table2.OpenCellRC(i, 3).Value = "CC" + i.ToString();
            table2.OpenCellRC(i, 4).Value = "DD" + i.ToString();
            table2.OpenCellRC(i, 5).Value = "EE" + i.ToString();
        }


        AceoffixCtrl1.Bind(doc);
        AceoffixCtrl1.CustomToolbar = false;
        AceoffixCtrl1.Menubar = false;

        AceoffixCtrl1.OpenDocument("doc/test.doc", Aceoffix.OpenModeType.docNormalEdit, "Tom");

    }
}