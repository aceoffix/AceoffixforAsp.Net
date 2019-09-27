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

public partial class wordForm_SaveData : System.Web.UI.Page
{
    public string ErrorMsg = "";
    public string BaseUrl = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        string strID = Request.QueryString["ID"];

        string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|demo2.mdb";
        OleDbConnection conn = new OleDbConnection(connString);
        conn.Open();

        string strsql;
        OleDbCommand cmd = new OleDbCommand();
        cmd.Connection = conn;
        cmd.CommandType = CommandType.Text;

        Aceoffix.WordReader.WordDocument doc = new  Aceoffix.WordReader.WordDocument();
        string sName = doc.OpenDataRegion("name").Value;
        string sDept = doc.OpenDataRegion("dept").Value;
        string sCause = doc.OpenDataRegion("cause").Value;
        string sDays = doc.OpenDataRegion("days").Value;
        string sDate = doc.OpenDataRegion("date").Value;

        if (sName=="")
        {
            ErrorMsg = ErrorMsg + "<li>Name</li>";
        }
        if (sDept=="")
        {
            ErrorMsg = ErrorMsg + "<li>Department</li>";
        }
        if (sCause=="")
        {
            ErrorMsg = ErrorMsg + "<li>Cause</li>";
        }
        if (sDate=="")
        {
            ErrorMsg = ErrorMsg + "<li>Date</li>";
        }
        try
        {
            if (sDays != "")
            {
                if (Int32.Parse(sDays) < 0)
                {
                    ErrorMsg = ErrorMsg + "<li>The number of the days can't be minus.</li>";
                }
            }
            else
            {
                ErrorMsg = ErrorMsg + "<li>Days</li>";
            }
        }
        catch
        {
            ErrorMsg = ErrorMsg + "<li><font color=red>Notice:</font>You should fill the 'Days' with numbers.</li>";
        }

        if (ErrorMsg == "")
        {
             strsql = "update leaveRecord set Name='" + sName
                    + "', Dept='" + sDept + "', Cause='" + sCause
                    + "', Num=" + sDays + ", SubmitTime='" + sDate
                    + "' where  ID=" + strID;
            cmd.CommandText = strsql;
            cmd.ExecuteNonQuery();
        }
        else
        {
            doc.ShowPage(578, 380);
        }
        doc.Close();
        conn.Close();


        string mScriptName = "savedata.aspx";
        string mHttpUrl = "http://" + Request.ServerVariables["HTTP_HOST"] + Request.ServerVariables["SCRIPT_NAME"];
        BaseUrl = mHttpUrl.Substring(0, mHttpUrl.Length - mScriptName.Length); 
    }
}