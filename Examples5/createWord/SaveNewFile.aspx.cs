using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.OleDb;
using System.Data;

public partial class createWord_SaveNewFile : System.Web.UI.Page
{
    string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|demo_creword.mdb";
    protected void Page_Load(object sender, EventArgs e)
    {
        Aceoffix.FileRequest freq = new Aceoffix.FileRequest();
        string subject = freq.GetFormField("FileSubject");
        string newFileName = Insert(subject);
        freq.SaveToFile(Server.MapPath("doc/") + newFileName);
        freq.CustomSaveResult = "ok";
        freq.Close();

    }
    string Insert(string subject)
    {
        string newID = "0";
        string strSql = "select Max(ID) from word";
        OleDbConnection conn = new OleDbConnection(connectionString);
        OleDbCommand cmd = new OleDbCommand(strSql, conn);
        conn.Open();
        cmd.CommandType = CommandType.Text;
        OleDbDataReader Reader = cmd.ExecuteReader();
        if (Reader.Read() && Reader[0].ToString().Trim().Length > 0)
        {
            newID = ((int)Reader[0] + 1).ToString();
        }
        Reader.Close();

        string fileName = "test" + newID + ".doc";
        string strsql = "Insert into word(ID,FileName,Subject,SubmitTime) values(" + newID
            + ",'" + fileName + "','" + subject + "','" + DateTime.Now.ToString() + "')";
        cmd.CommandText = strsql;
        cmd.ExecuteNonQuery();
        conn.Close();

        return fileName;
    }
}