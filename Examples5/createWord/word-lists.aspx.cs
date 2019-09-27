using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.OleDb;
using System.Data;
using System.Text;

public partial class createWord_word_lists : System.Web.UI.Page
{
    string strSql = "";
    string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|demo_creword.mdb";
    string newID = "1";
    string fileName = "";
    string FileSubject = "";
    protected string strHtmls = "";
    protected void Page_Load(object sender, EventArgs e)
    {

          string strSql = "select * from word order by ID desc ";

        DataSet ds = new DataSet();
        using (OleDbConnection connection = new OleDbConnection(connectionString))
        {
            try
            {
                connection.Open();
                OleDbDataAdapter command = new OleDbDataAdapter(strSql, connection);
                command.Fill(ds, "ds");
                connection.Close();
            }
            catch (System.Data.OleDb.OleDbException ex)
            {
                throw new Exception(ex.Message);
            }
        }

        if (ds != null && ds.Tables[0].Rows.Count > 0)
        {
            DataTable dt = ds.Tables[0];
            StringBuilder strHtml = new StringBuilder();

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                strHtml.Append("<tr onmouseover='onColor(this)' onmouseout='offColor(this)'>\n");
                strHtml.Append("<td><a href = \"javascript:AceBrowser.openWindowModeless('OpenWord.aspx?filename=" + dt.Rows[i]["FileName"] + "&subject=" + dt.Rows[i]["Subject"] + "','width=1200px;height=800px;');\">" + dt.Rows[i]["Subject"] + "</a></td>\n");
                if (dt.Rows[i]["SubmitTime"] != null && dt.Rows[i]["SubmitTime"].ToString().Trim().Length > 0)
                {
                    strHtml.Append("<td>" + DateTime.Parse(dt.Rows[i]["SubmitTime"].ToString()).ToString("MM/dd/yyyy") + "</td>\n");
                }
                else
                {
                    strHtml.Append("<td>&nbsp;</td>\n");
                }
                strHtml.Append(" </tr>\n");
            }

            strHtmls = strHtml.ToString();
        }
       
    }


    protected void btnCreate0_Click(object sender, EventArgs e)
    {
     
        Insert();
      
        System.IO.File.Copy(Server.MapPath("doc/template.doc"),
            Server.MapPath("doc/" + fileName), true);
    
        Response.Redirect("word-lists.aspx");
    }

    void Insert()
    {
        strSql = "select Max(ID) from word";
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

        fileName = "test" + newID + ".doc";
        FileSubject = "Enter a Filename";
        if (Request.Form["FileSubject"] != "")
            FileSubject = Request.Form["FileSubject"];
        string strsql = "Insert into word(ID,FileName,Subject,SubmitTime) values(" + newID
            + ",'" + fileName + "','" + FileSubject + "','" + DateTime.Now.ToString() + "')";
        cmd.CommandText = strsql;
        cmd.ExecuteNonQuery();
        conn.Close();
    }
}