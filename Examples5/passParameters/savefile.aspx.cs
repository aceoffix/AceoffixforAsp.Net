using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.OleDb;
using System.Data;

public partial class passParameters_savefile : System.Web.UI.Page
{
    public int id = 0;
    public string userName = "";
    public int age = 0;
    public string sex = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        Aceoffix.FileRequest freq = new Aceoffix.FileRequest();
        freq.SaveToFile(Server.MapPath("doc/") + freq.FileName);


        if (Request.QueryString["id"] != null && Request.QueryString["id"].Trim().Length > 0) //Get the parameter
            id = int.Parse(Request.QueryString["id"].Trim());


        if (freq.GetFormField("userName") != null && freq.GetFormField("userName").Trim().Length > 0) //Get the form field
        {
            userName = freq.GetFormField("userName");
        }


        if (freq.GetFormField("age") != null && freq.GetFormField("age").Trim().Length > 0)
        {
            age = int.Parse(freq.GetFormField("age"));
        }


        if (freq.GetFormField("selSex") != null && freq.GetFormField("selSex").Trim().Length > 0)
        {
            sex = freq.GetFormField("selSex");
        }

        string connstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|demo_param.mdb";
        OleDbConnection conn = new OleDbConnection(connstring);
        conn.Open();

        string strsql;
        OleDbCommand cmd = new OleDbCommand();
        cmd.Connection = conn;
        cmd.CommandType = CommandType.Text;
        strsql = "Update Users set UserName = '" + userName + "', age = " + age + ",sex = '" + sex + "' where id = " + id;
        cmd.CommandText = strsql;
        cmd.ExecuteNonQuery();
        conn.Close();

        freq.ShowPage(300, 200);
        freq.Close();
    }
}