using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aceoffix.ExcelReader;

public partial class definedNameTable_SaveData : System.Web.UI.Page
{
    protected string content = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        Workbook workBook = new Workbook();
        Sheet sheet = workBook.OpenSheet("Sheet1");

        Aceoffix.ExcelReader.Table table = sheet.OpenTableByDefinedName("report");
        
        int result = 0;
        while (!table.EOF)
        {

            if (!table.DataFields.IsEmpty)
            {
                content += "<br/>Month:"
                        + table.DataFields[0].Text;
                content += "<br/>Plan:"
                        + table.DataFields[1].Text;
                content += "<br/>Reality:"
                        + table.DataFields[2].Text;
                content += "<br/>Accumulative:"
                        + table.DataFields[3].Text;

                if (string.IsNullOrEmpty(table.DataFields[2].Text) || !int.TryParse(table.DataFields[2].Text, out result) ||
                    !int.TryParse(table.DataFields[1].Text, out result))
                {
                    content += "<br/>Order Fill Rate:0%";
                }
                else
                {
                    float f = int.Parse(table.DataFields[2].Text);
                    f = f / int.Parse(table.DataFields[1].Text);

                    content += "<br/>Order Fill Rate:" +string.Format("{0:P}",f);
                }
                content += "<br/>*********************************************";
            }

            table.NextRow();
        }
        table.Close();
        workBook.ShowPage(500, 400);
        workBook.Close();
    }
}