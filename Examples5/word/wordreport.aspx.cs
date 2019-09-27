using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class word_wordreport : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    protected void AceoffixCtrl1_Load(object sender, EventArgs e)
    {
        string[] strNames = { "John Scott", "Tom Bush", "Steven Nicholls", "Kate Middleton" };
        string[] strGrades = { "C", "B", "A" };
        Aceoffix.WordWriter.WordDocument wd = new Aceoffix.WordWriter.WordDocument();
        System.Random rd = new Random(System.DateTime.Now.Millisecond);
        wd.OpenDataRegion("StudentName1").Value = strNames[rd.Next(4)];
        wd.OpenDataRegion("StudentName2").Value = strNames[rd.Next(4)];

        int iCreditsAttempted = 0;
        int iTemp = 0;
        for (int i = 0; i < 9; i++)
        {
            int iCredits = rd.Next(5) + 1;
            int iScore = rd.Next(3) + 2;
            wd.OpenDataRegion("Table4").OpenTable(1).OpenCellRC(3 + i, 2).Value = iCredits.ToString();
            wd.OpenDataRegion("Table4").OpenTable(1).OpenCellRC(3 + i, 3).Value = iCredits.ToString();
            wd.OpenDataRegion("Table4").OpenTable(1).OpenCellRC(3 + i, 4).Value = strGrades[iScore - 2];
            iCreditsAttempted = iCreditsAttempted + iCredits;
            iTemp = iTemp + iCredits * iScore;
        }
        float fGPA = (float)iTemp / iCreditsAttempted;
        wd.OpenDataRegion("Table4").OpenTable(1).OpenCellRC(12, 1).Value = "Total Credits2:  " + iCreditsAttempted.ToString() + "        GPA:       " + fGPA.ToString("F2");

        wd.OpenDataRegion("GPA").Value = fGPA.ToString("F2");
        wd.OpenDataRegion("CreditsAttempted").Value = iCreditsAttempted.ToString();
        wd.OpenDataRegion("CreditsEarned").Value = iCreditsAttempted.ToString();

        AceoffixCtrl1.Caption = "Generate Word report dynamically.";
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        // Create custom toolbar
        AceoffixCtrl1.AddCustomToolButton("Save As", "SaveAsDocument()", 1);
        AceoffixCtrl1.AddCustomToolButton("-", "", 0);
        AceoffixCtrl1.AddCustomToolButton("Print", "ShowPrintDlg()", 6);
        AceoffixCtrl1.AddCustomToolButton("-", "", 0);
        AceoffixCtrl1.AddCustomToolButton("Full-screen Switch", "SwitchFullScreen()", 4);

        AceoffixCtrl1.FileTitle = "Report1";
        AceoffixCtrl1.Bind(wd);
        AceoffixCtrl1.OpenDocument("../doc/doc-report.doc", Aceoffix.OpenModeType.docReadOnly, "John Scott"); // docReadOnly: Disabled Edit, Paste.
    }
}
