using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class saveDataAndFile_savefile : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Aceoffix.FileRequest freq = new Aceoffix.FileRequest();
        freq.SaveToFile(Server.MapPath("doc/") + freq.FileName);
        freq.Close();
    }
}