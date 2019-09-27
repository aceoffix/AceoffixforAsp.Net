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

public partial class wordForm_GenDoc : System.Web.UI.Page
{
    public string docID;
    public string docFile;
    public string docName;
    public string docDept;
    public string docCause;
    public string docDays;
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
            docFile = Reader["FileName"].ToString();
            docName = Reader["Name"].ToString();
            docDept = Reader["Dept"].ToString();
            docCause = Reader["Cause"].ToString();
            docDays = Reader["Num"].ToString();
            docSubmitTime = DateTime.Parse(Reader["SubmitTime"].ToString());
        }
        Reader.Close();
        conn.Close();

        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        AceoffixCtrl1.BorderStyle = Aceoffix.BorderStyleType.BorderThin;
        Aceoffix.WordWriter.WordDocument doc = new Aceoffix.WordWriter.WordDocument();
        doc.DisableWindowRightClick=true;
        doc.OpenDataRegion("name").Value = docName;
        doc.OpenDataRegion("dept").Value = docDept;
        doc.OpenDataRegion("cause").Value = docCause;
        doc.OpenDataRegion("days").Value = docDays;
        doc.OpenDataRegion("date").Value = docSubmitTime.ToLongDateString();
        doc.OpenDataRegion("ToDate").Value = docSubmitTime.ToLongDateString();
        doc.OpenDataRegion("tip").Value = "";
        AceoffixCtrl1.Bind(doc);
        AceoffixCtrl1.Menubar = false;
        AceoffixCtrl1.AddCustomToolButton("Print", "poPrint", 1);
        AceoffixCtrl1.AddCustomToolButton("Full-screen Switch", "poSetFullScreen", 4);
        
        AceoffixCtrl1.SaveDataPage="SaveData.aspx?ID=" + docID;
        AceoffixCtrl1.OpenDocument("doc/template.doc", Aceoffix.OpenModeType.docReadOnly, "John Scott");


     

    }
}