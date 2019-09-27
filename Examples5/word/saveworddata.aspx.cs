using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;

public partial class word_saveworddata : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        // This page is used to receive the data submitted by AceoffixCtrl.
        Aceoffix.WordReader.WordDocument wd = new Aceoffix.WordReader.WordDocument();

        StringBuilder sb = new StringBuilder();
        sb.Append("<p>EmployeeName: <span class='TypeValue'>" + wd.OpenDataRegion("EmployeeName").Value + "</span></p>");
        sb.Append("<p>EmployeeNumber: <span class='TypeValue'>" + wd.OpenDataRegion("EmployeeNumber").Value + "</span></p>");
        sb.Append("<p>Department: <span class='TypeValue'>" + wd.OpenDataRegion("Department").Value + "</span></p>");
        sb.Append("<p>Manager: <span class='TypeValue'>" + wd.OpenDataRegion("Manager").Value + "</span></p>");
        sb.Append("<p>FromDate: <span class='TypeValue'>" + wd.OpenDataRegion("FromDate").Value + "</span></p>");
        sb.Append("<p>ToDate: <span class='TypeValue'>" + wd.OpenDataRegion("ToDate").Value + "</span></p>");
        sb.Append("<p>Reason: <span class='TypeValue'>" + wd.OpenDataRegion("Reason").Value + "</span></p>");
        Label1.Text = sb.ToString();

        // In general, SaveDataPage do not need to dispaly anything. But if you call ShowPage method, this page will be shown in a popup dialog box after saving document. 
        wd.ShowPage(800, 600);
        // Once you saved data, please call this method finally.
        wd.Close();
    }
}
