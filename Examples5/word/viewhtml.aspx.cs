using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class word_viewhtml : System.Web.UI.Page
{
    public string m_htmlURL = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        if (Request.QueryString["mht"] != null)
        {
            System.Random rd = new Random(System.DateTime.Now.Millisecond);
            m_htmlURL = Request.QueryString["mht"] + "?rd=" + rd.Next(1000).ToString();
        }
    }
}
