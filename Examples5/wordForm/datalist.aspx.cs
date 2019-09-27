using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data.OleDb;

public partial class wordForm_datalist : System.Web.UI.Page
{
    public string docID;
    public string docFile;
    public string docSubject;
    public string docName;
    public string docDept;
    public string docCause;
    public string docDays;
    public string docDate;
    public DateTime docSubmitTime;
    protected void Page_Load(object sender, EventArgs e)
    {
        docID = Request.QueryString["ID"];

        string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|demo2.mdb";
        OleDbConnection conn = new OleDbConnection(connString);
        conn.Open();

        string sql = "select * from leaveRecord where ID=" + docID;
        OleDbCommand cmd = new OleDbCommand(sql, conn);
        cmd.CommandType = CommandType.Text;
        OleDbDataReader Reader = cmd.ExecuteReader();
        if (Reader.Read())
        {
            docSubject = Reader["Subject"].ToString();
            docName = Reader["Name"].ToString();
            docDept = Reader["Dept"].ToString();
            docCause = Reader["Cause"].ToString();
            docDays = Reader["Num"].ToString();

            docSubmitTime = DateTime.Parse(Reader["SubmitTime"].ToString());
            docDate = docSubmitTime.ToLongDateString();
        }
        Reader.Close();
        conn.Close();

    }
}