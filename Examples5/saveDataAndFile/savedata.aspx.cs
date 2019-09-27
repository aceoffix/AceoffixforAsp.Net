using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class saveDataAndFile_savedata : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Aceoffix.WordReader.WordDocument doc = new Aceoffix.WordReader.WordDocument();
        string Name = doc.OpenDataRegion("Name").Value;
        
        string Department = doc.OpenDataRegion("Department").Value;
        string companyName = doc.GetFormField("txtCompany");

        //You can save them to the database.

        doc.Close(); 
    }
}