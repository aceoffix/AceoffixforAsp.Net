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
using System.Drawing;

public partial class wordForm_SubmitDataOfDoc : System.Web.UI.Page
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



        Aceoffix.WordWriter.WordDocument doc = new Aceoffix.WordWriter.WordDocument();
        Aceoffix.WordWriter.DataRegion drName = doc.OpenDataRegion("name");
        drName.Value=docName;
        drName.Editing=true;
        Aceoffix.WordWriter.DataRegion drDept = doc.OpenDataRegion("dept");
        drDept.Value=docDept;
        drDept.Shading.BackgroundPatternColor=Color.Gray;
        //drDept.setEditing(true);  
        Aceoffix.WordWriter.DataRegion drCause = doc.OpenDataRegion("cause");
        drCause.Value=docCause;
        drCause.Editing=true;
        Aceoffix.WordWriter.DataRegion drNum = doc.OpenDataRegion("days");
        drNum.Value=docDays;
        drNum.Editing=true;
        Aceoffix.WordWriter.DataRegion drDate = doc.OpenDataRegion("date");
        drDate.Value=docSubmitTime.ToLongDateString();
        drDate.Shading.BackgroundPatternColor=Color.Pink;
        Aceoffix.WordWriter.DataRegion toDate = doc.OpenDataRegion("ToDate");
        toDate.Editing=true;
        toDate.Value=docSubmitTime.ToLongDateString();

        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";
        AceoffixCtrl1.BorderStyle=Aceoffix.BorderStyleType.BorderThin;

        AceoffixCtrl1.AddCustomToolButton("Save", "poSave()", 1);
        AceoffixCtrl1.AddCustomToolButton("Full-screen Switch", "poSetFullScreen()", 4);

        AceoffixCtrl1.OfficeToolbars=false;
        AceoffixCtrl1.Menubar=false;

        AceoffixCtrl1.JsFunction_OnWordDataRegionClick="OnWordDataRegionClick()";

        AceoffixCtrl1.SaveDataPage="SaveData.aspx?ID=" + docID;

        AceoffixCtrl1.Bind(doc);

        AceoffixCtrl1.OpenDocument("doc/template.doc", Aceoffix.OpenModeType.docSubmitForm, "Tom");
       
	

    }
}