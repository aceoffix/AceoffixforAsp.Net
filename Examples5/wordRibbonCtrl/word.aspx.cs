using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class wordRibbonCtrl_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        AceoffixCtrl1.RibbonBar.SetTabVisible("TabHome", true);//Home
        AceoffixCtrl1.RibbonBar.SetTabVisible("TabPageLayoutWord", false);//Page Layout
        AceoffixCtrl1.RibbonBar.SetTabVisible("TabReferences", false);//References
        AceoffixCtrl1.RibbonBar.SetTabVisible("TabMailings", false);//Mailings
        AceoffixCtrl1.RibbonBar.SetTabVisible("TabReviewWord", false);//Review
        AceoffixCtrl1.RibbonBar.SetTabVisible("TabInsert", false);//Insert
        AceoffixCtrl1.RibbonBar.SetTabVisible("TabView", false);//View

        AceoffixCtrl1.RibbonBar.SetSharedVisible("FileSave", false);//Office save button
        AceoffixCtrl1.RibbonBar.SetGroupVisible("GroupClipboard", false);//Clipboard

        AceoffixCtrl1.Menubar = false;
        AceoffixCtrl1.CustomToolbar = false;
        AceoffixCtrl1.OpenDocument("doc/introduce.doc", Aceoffix.OpenModeType.docNormalEdit, "Tom");
    }
}