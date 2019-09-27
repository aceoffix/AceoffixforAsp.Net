using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class word_saveuploaddoc : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        // This page is used to receive the file submitted by AceoffixCtrl.
        Aceoffix.FileRequest freq = new Aceoffix.FileRequest();
        freq.SaveToFile(Server.MapPath("../doc/") + freq.FileName);
        Label1.Text = freq.DocumentText;
        freq.ShowPage(500, 500);
        freq.Close();
    }
}
