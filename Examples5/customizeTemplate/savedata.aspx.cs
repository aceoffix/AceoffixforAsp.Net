using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class customizeTemplate_savedata : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Aceoffix.WordReader.WordDocument doc = new Aceoffix.WordReader.WordDocument();
	
    try{
        Aceoffix.WordReader.DataRegion poName = doc.OpenDataRegion("Name");
         Response.Write("Get the value of ACE-Name from back end:" + poName.Value);
    }
    catch{
        Response.Write("No DataRegion named ACE-Name submitted from client side.");
    }
	doc.ShowPage(400, 300);
	doc.Close();

    }
}