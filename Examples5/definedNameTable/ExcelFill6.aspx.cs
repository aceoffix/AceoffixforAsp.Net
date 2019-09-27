using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aceoffix.ExcelWriter;

public partial class definedNameTable_ExcelFill6 : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        AceoffixCtrl1.Menubar = false;

        AceoffixCtrl1.OpenDocument("doc/test4.xls", Aceoffix.OpenModeType.xlsNormalEdit, "John Scott");

    }
}