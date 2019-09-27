using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;

public partial class extractTable_savedata : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Aceoffix.WordReader.WordDocument doc=new Aceoffix.WordReader.WordDocument();
        Aceoffix.WordReader.DataRegion dataReg=doc.OpenDataRegion("table");
        Aceoffix.WordReader.Table table=dataReg.OpenTable(1);
        Response.Write("Table Data:<br/><br/>");
        StringBuilder dataStr = new StringBuilder();
        for (int i = 1; i <= table.RowsCount; i++)
        {
            dataStr.Append("<div style='width:220px;'>");
            for (int j = 1; j <= table.ColumnsCount; j++)
            {
                dataStr.Append("<div style='float:left;width:70px;border:1px solid red;'>"+table.OpenCellRC(i,j).Value+"</div>");
            }
            dataStr.Append("</div>");
        }
        Response.Write(dataStr.ToString());
        doc.ShowPage(350, 300);
        doc.Close();

    }
}