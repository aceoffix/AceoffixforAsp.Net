using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.OleDb;
using System.Data;

public partial class examinationPaper_SaveFile : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Aceoffix.FileRequest freq = new Aceoffix.FileRequest();
        
        string id = Request.QueryString["id"];
        if (id != null && id.Length > 0)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|demo_paper.mdb";
            OleDbConnection conn = new OleDbConnection(strConn);
            byte[] aa = freq.FileBytes;

            OleDbCommand cm = new OleDbCommand();
            cm.Connection = conn;
            cm.CommandType = CommandType.Text;
            if (conn.State == 0) conn.Open();
            cm.CommandText = "UPDATE  Stream SET Word=@file WHERE ID=" + id;
            OleDbParameter spFile = new OleDbParameter("@file", OleDbType.Binary);
            spFile.Value = aa;
            cm.Parameters.Add(spFile);
            cm.ExecuteNonQuery();
            conn.Close();

            freq.CustomSaveResult = "ok";

        }
        else
        {
            Page.ClientScript.RegisterStartupScript(this.GetType(), "script1", "alert('Save Failure');location.href='Default.aspx';", true);
        }

        freq.Close();


    }
}