using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aceoffix.ExcelReader;

public partial class definedNameCell_savedata : System.Web.UI.Page
{
    protected string content = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        Workbook workBook = new Workbook();
        Sheet sheet = workBook.OpenSheet("Sheet1");

       
        content += "testA:" + sheet.OpenCellByDefinedName("testA").Value + "<br/>";
        content += "testB:" + sheet.OpenCellByDefinedName("testB").Value + "<br/>";

        workBook.ShowPage(500, 400);
        workBook.Close();
    }
}