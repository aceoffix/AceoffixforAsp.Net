using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aceoffix.ExcelWriter;

public partial class excelRibbonCtrl_excel : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        AceoffixCtrl1.RibbonBar.SetTabVisible("TabHome", true);//Home
        AceoffixCtrl1.RibbonBar.SetTabVisible("TabFormulas", false);//Formulas
        AceoffixCtrl1.RibbonBar.SetTabVisible("TabInsert", false);//Insert
        AceoffixCtrl1.RibbonBar.SetTabVisible("TabData", false);//Data
        AceoffixCtrl1.RibbonBar.SetTabVisible("TabReview", false);//Review
        AceoffixCtrl1.RibbonBar.SetTabVisible("TabView", false);//View
        AceoffixCtrl1.RibbonBar.SetTabVisible("TabPageLayoutExcel", false);//PageLayout

        AceoffixCtrl1.RibbonBar.SetSharedVisible("FileSave", false);//Office Save

        AceoffixCtrl1.RibbonBar.SetGroupVisible("GroupClipboard", false);//Clipboard

        AceoffixCtrl1.OpenDocument("doc/test.xls", Aceoffix.OpenModeType.xlsNormalEdit, "John Scott");

    }
}