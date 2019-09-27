using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;

public partial class wordEditByUser_savedata : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Aceoffix.WordReader.WordDocument doc = new Aceoffix.WordReader.WordDocument();

        if (Request.QueryString["userName"] != null && Request.QueryString["userName"].Equals("zhangsan"))
        {
            saveBytesToFile(doc.OpenDataRegion("Paragraph1").FileBytes, Server.MapPath("doc/paragraph1.doc"));
        }
        else
        {
            saveBytesToFile(doc.OpenDataRegion("Paragraph2").FileBytes, Server.MapPath("doc/paragraph2.doc"));
        }

        doc.Close();
    }
    private void saveBytesToFile(byte[] bytes, string filePath)
    {
        FileStream fs = new FileStream(filePath, System.IO.FileMode.OpenOrCreate);
        fs.Write(bytes, 0, bytes.Length);
        fs.Close();
    }

}