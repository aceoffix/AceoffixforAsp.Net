using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;

public partial class word_docstream : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string strID = Request.QueryString["id"];
        if (strID == null)
        {
            Response.Clear();
            Response.End();
            return;
        }

        Response.Clear(); // Required
        Response.ContentType = "application/msword"; // Required
        Response.AddHeader("Content-Disposition", "attachment; filename=down.doc"); // Required

        string connstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|demo.mdb";
        OleDbConnection conn = new OleDbConnection(connstring);
        conn.Open();
        OleDbCommand cmd = new OleDbCommand();
        cmd.Connection = conn;
        cmd.CommandType = CommandType.Text;
        cmd.CommandText = "select FileBin from documents where ID = @id";
        OleDbParameter spID = new OleDbParameter("@id", OleDbType.Integer);
        spID.Value = strID;
        cmd.Parameters.Add(spID);
        OleDbDataReader dr = cmd.ExecuteReader();
        int iFileSize = 0;
        if (dr.Read())
        {
            int FileCol = 0; // the column # of the BLOB field
            Byte[] b = new Byte[(dr.GetBytes(FileCol, 0, null, 0, int.MaxValue))];
            dr.GetBytes(FileCol, 0, b, 0, b.Length);
            iFileSize = b.Length;
            System.IO.Stream fs = Response.OutputStream;
            fs.Write(b, 0, b.Length);
            fs.Close();
        }
        dr.Close();
        conn.Close();
        Response.AppendHeader("Content-Length", iFileSize.ToString()); // Required
        Response.End(); // Required
    }
}
