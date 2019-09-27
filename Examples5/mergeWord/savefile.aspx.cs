using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class mergeWord_savefile : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        // This page is used to receive the file submitted by AceoffixCtrl.
        Aceoffix.FileRequest freq = new Aceoffix.FileRequest();
        freq.SaveToFile(Server.MapPath("doc/") + freq.FileName);
        freq.Close();
    }
}