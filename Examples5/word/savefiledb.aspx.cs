using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.OleDb;

public partial class word_savefiledb : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        // This page is used to receive the file submitted by AceoffixCtrl.
        string strID = Request.QueryString["id"];
        if (strID == null)
        {
            Response.Write("The ID querystring of this page is not found.");
            return;
        }

        Aceoffix.FileRequest freq = new Aceoffix.FileRequest();

        string connstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|demo.mdb";
        OleDbConnection conn = new OleDbConnection(connstring);
        conn.Open();
        OleDbCommand cmd = new OleDbCommand();
        cmd.Connection = conn;
        cmd.CommandType = CommandType.Text;
        cmd.CommandText = "update documents set FileType = @filetype, FileName = @filename, FileExtName = @fileextname, FileSize = @filesize, FileBin = @file where ID = @id";
        //cmd.CommandText = "insert into documents(FileType, FileName, FileExtName, FileSize, FileBin) values(@filetype, @filename, @fileextname, @filesize, @file)"; // for adding new file.
        OleDbParameter spFileType = new OleDbParameter("@filetype", OleDbType.VarChar, 50);
        spFileType.Value = "Word";
        cmd.Parameters.Add(spFileType);
        OleDbParameter spFileName = new OleDbParameter("@filename", OleDbType.VarChar, 50);
        spFileName.Value = freq.FileName;
        cmd.Parameters.Add(spFileName);
        OleDbParameter spFileExtName = new OleDbParameter("@fileextname", OleDbType.VarChar, 50);
        spFileExtName.Value = freq.FileExtName;
        cmd.Parameters.Add(spFileExtName);
        OleDbParameter spFileSize = new OleDbParameter("@filesize", OleDbType.Integer);
        spFileSize.Value = freq.FileSize;
        cmd.Parameters.Add(spFileSize);
        OleDbParameter spFile = new OleDbParameter("@file", OleDbType.Binary);
        spFile.Value = freq.FileBytes;
        cmd.Parameters.Add(spFile);
        OleDbParameter spID = new OleDbParameter("@id", OleDbType.Integer);
        spID.Value = strID;
        cmd.Parameters.Add(spID);
        cmd.ExecuteNonQuery();
        conn.Close();

        freq.Close();

    }
}
