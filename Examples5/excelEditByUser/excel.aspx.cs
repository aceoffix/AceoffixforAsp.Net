using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aceoffix.ExcelWriter;
using System.Drawing;

public partial class excelEditByUser_excel : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string userName = Request.QueryString["userName"];
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Workbook wb = new Workbook();
        Sheet sheet = wb.OpenSheet("Sheet1");
        Aceoffix.ExcelWriter.Table tableA = sheet.OpenTable("C4:D6");
        Aceoffix.ExcelWriter.Table tableB = sheet.OpenTable("C7:D9");
        if (userName.Equals("Tom"))
        {
            Literal1.Text = "Tom";
            tableA.ReadOnly=false;
            tableA.BackColor = Color.Yellow;
            tableB.ReadOnly=true;
        }
        else
        {
            Literal1.Text = "Jack";
            tableA.ReadOnly=true;
            tableB.BackColor = Color.Yellow;
            tableB.ReadOnly=false;
        }

        AceoffixCtrl1.Bind(wb);
        AceoffixCtrl1.Menubar = false;
        AceoffixCtrl1.AddCustomToolButton("Save", "Save()", 0);
        AceoffixCtrl1.SaveFilePage = "savefile.aspx";
        AceoffixCtrl1.OpenDocument("doc/test.xls", Aceoffix.OpenModeType.xlsSubmitForm, "John Scott");


    }
}