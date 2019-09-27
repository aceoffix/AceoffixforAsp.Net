using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.IO;

public partial class wordSplit_savedata : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

     
        Aceoffix.WordReader.WordDocument doc = new Aceoffix.WordReader.WordDocument();
        Byte[] bWord;
        Aceoffix.WordReader.DataRegion dr1 = doc.OpenDataRegion("Paragraph1");
        bWord = dr1.FileBytes;
        Stream s1 = new FileStream(Server.MapPath("doc/") + "Paragraph1.doc", FileMode.Create);
        s1.Write(bWord, 0, bWord.Length);
        s1.Close();


        Aceoffix.WordReader.DataRegion dr2 = doc.OpenDataRegion("Paragraph2");
        bWord = dr2.FileBytes;
        Stream s2 = new FileStream(Server.MapPath("doc/") + "Paragraph2.doc", FileMode.Create);
        s2.Write(bWord, 0, bWord.Length);
        s2.Close();


        Aceoffix.WordReader.DataRegion dr3 = doc.OpenDataRegion("Paragraph3");
        bWord = dr3.FileBytes;
        Stream s3 = new FileStream(Server.MapPath("doc/") + "Paragraph3.doc", FileMode.Create);
        s3.Write(bWord, 0, bWord.Length);
        s3.Close();

        doc.Close();
    }
}