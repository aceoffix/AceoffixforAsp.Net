using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aceoffix.ExcelReader;

public partial class clickExcelCell_SaveData : System.Web.UI.Page
{
    protected string content = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        Workbook workBook = new Workbook();
        Sheet sheet = workBook.OpenSheet("Sheet1");
        Aceoffix.ExcelReader.Table table = sheet.OpenTable("B4:D8");
       
        while (!table.EOF)
        {

            if (!table.DataFields.IsEmpty)
            {
                content += "<br/>Month:" + table.DataFields[0].Text;
                content += "<br/>Plan:" + table.DataFields[1].Text;
                content += "<br/>Reality:" + table.DataFields[2].Text;

                content += "<br/>*********************************************";
            }

            table.NextRow();
        }
        table.Close();


        workBook.ShowPage(500, 400);
        workBook.Close();
    }
}