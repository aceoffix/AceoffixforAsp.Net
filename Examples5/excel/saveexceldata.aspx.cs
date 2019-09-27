using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;

public partial class excel_saveexceldata : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        // This page is used to receive the data submitted by AceoffixCtrl.
        Aceoffix.ExcelReader.Workbook wb = new Aceoffix.ExcelReader.Workbook();
        Aceoffix.ExcelReader.Sheet sheet1 = wb.OpenSheet("Purchase Order");

        StringBuilder sb = new StringBuilder();
        sb.Append("<p>OrderNumber: <span class='TypeValue'>" + sheet1.OpenCell("D8").Value + "</span></p>");

        Aceoffix.ExcelReader.Table table1 = sheet1.OpenTable("B18:J23");
        sb.Append("<p>ProductItems: <br><table border=1><tr><td>QTY</td><td>DESCRIPTION</td><td>UNIT PRICE</td><td>AMOUNT</td></tr>");
        while (!table1.EOF)
        {
            string strCells = "";
            if (!table1.DataFields.IsEmpty)
            {
                strCells = strCells + "<td>" + table1.DataFields[0].Value + "</td>";
                strCells = strCells + "<td>" + table1.DataFields[1].Value + "</td>";
                strCells = strCells + "<td>" + table1.DataFields[7].Value + "</td>"; // Skip the merged columns.
                strCells = strCells + "<td>" + table1.DataFields[8].Value + "</td>";
                sb.Append("<tr>" + strCells + "</tr>");
            }
            table1.NextRow();
        }
        table1.Close();
        sb.Append("</table></p>");

        sb.Append("<p>Subtotal: <span class='TypeValue'>" + sheet1.OpenCell("J24").Value + "</span></p>");
        sb.Append("<p>Shipping: <span class='TypeValue'>" + sheet1.OpenCell("J25").Value + "</span></p>");
        sb.Append("<p>Tax: <span class='TypeValue'>" + sheet1.OpenCell("J26").Value + "</span></p>");
        sb.Append("<p>Other: <span class='TypeValue'>" + sheet1.OpenCell("J27").Value + "</span></p>");
        sb.Append("<p>Total: <span class='TypeValue'>" + sheet1.OpenCell("J28").Value + "</span></p>");
        Label1.Text = sb.ToString();

        // In general, SaveDataPage do not need to dispaly anything. But if you call ShowPage method, this page will be shown in a popup dialog box after saving document. 
        wb.ShowPage(800, 600);
        // Once you saved data, please call this method finally.
        wb.Close();
    }
}
