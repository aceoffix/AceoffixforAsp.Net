using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.OleDb;
using System.Data;
using Aceoffix.WordWriter;

public partial class examinationPaper_Compose2 : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|demo_paper.mdb";
        string strSql = "select * from stream ";

        OleDbConnection conn = new OleDbConnection(strConn);
        OleDbCommand cmd = new OleDbCommand(strSql, conn);
        conn.Open();
        cmd.CommandType = CommandType.Text;
        OleDbDataReader reader = cmd.ExecuteReader();

        int id = 0;
        string temp = "ACE_begin1";

        WordDocument doc = new WordDocument();
        if (reader.HasRows)
        {
            int num = 1;
            while (reader.Read())
            {
                

                id = int.Parse(reader["ID"].ToString());
                string chk = Request.Form["check" + id];
                if (chk != null && chk.Equals("on", StringComparison.OrdinalIgnoreCase))
                {
                    if (id == 1)
                    {
                        DataRegion dataNum = doc.OpenDataRegion("ACE_begin1");
                        dataNum.Value = "1.\t";
                        DataRegion dataReg = doc.CreateDataRegion("begin" + (id + 1), DataRegionInsertType.After, "begin1");
                        dataReg.Value = "[word]Openfile.aspx?id=" + id + "[/word]";
                    }
                    else
                    {
                        DataRegion dataNum = doc.CreateDataRegion("ACE_" + num, DataRegionInsertType.After, temp);
                        dataNum.Value = num + ".\t";
                        DataRegion dataRegion = doc.CreateDataRegion("begin" + (id + 1), DataRegionInsertType.After, "ACE_" + num);
                        dataRegion.Value = "[word]Openfile.aspx?id=" + id + "[/word]";
                    }
                    
                    temp = "ACE_begin" + (id + 1);
                    num++;
                }
            }

        }

        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        AceoffixCtrl1.Menubar = false;
        AceoffixCtrl1.CustomToolbar = false;
        AceoffixCtrl1.Bind(doc);
       
        AceoffixCtrl1.OpenDocument("doc/test.doc", Aceoffix.OpenModeType.docNormalEdit, "John Scott");


    }
}