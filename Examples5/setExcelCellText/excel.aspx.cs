using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aceoffix.ExcelWriter;
using System.Drawing;

public partial class setExcelCellText_excel : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Workbook wb = new Workbook();
        Sheet sheet = wb.OpenSheet("Sheet1");
        Aceoffix.ExcelWriter.Cell cC3 = sheet.OpenCell("C3");
        cC3.BackColor = Color.LightGray;
        cC3.Value = "Jan";
        cC3.ForeColor = Color.White;
        cC3.HorizontalAlignment = Aceoffix.ExcelWriter.XlHAlign.xlHAlignCenter;
        Aceoffix.ExcelWriter.Cell cD3 = sheet.OpenCell("D3");
        cD3.BackColor = Color.LightGray;
        cD3.Value = "Feb";
        cD3.ForeColor = Color.White;
        cD3.HorizontalAlignment = Aceoffix.ExcelWriter.XlHAlign.xlHAlignCenter;
        Aceoffix.ExcelWriter.Cell cE3 = sheet.OpenCell("E3");
        cE3.BackColor = Color.LightGray;
        cE3.Value = "Mar";
        cE3.ForeColor = Color.White;

        cE3.HorizontalAlignment = Aceoffix.ExcelWriter.XlHAlign.xlHAlignCenter;
        Aceoffix.ExcelWriter.Cell cB4 = sheet.OpenCell("B4");
        cB4.BackColor = Color.Blue;
        cB4.Value = "Lodging";
        cB4.Font.Size = 20;
        cB4.ForeColor = Color.Yellow;
        cB4.HorizontalAlignment = Aceoffix.ExcelWriter.XlHAlign.xlHAlignCenter;
        Aceoffix.ExcelWriter.Cell cB5 = sheet.OpenCell("B5");
        cB5.BackColor = Color.Teal;
        cB5.Value = "Three Meals";
        cB5.ForeColor = Color.Wheat;
        cB5.HorizontalAlignment = Aceoffix.ExcelWriter.XlHAlign.xlHAlignCenter;
        Aceoffix.ExcelWriter.Cell cB6 = sheet.OpenCell("B6");
        cB6.BackColor = Color.MediumPurple;
        cB6.Value = "Car Fare";
        cB6.ForeColor = Color.Wheat;
        cB6.HorizontalAlignment = Aceoffix.ExcelWriter.XlHAlign.xlHAlignCenter;
        Aceoffix.ExcelWriter.Cell cB7 = sheet.OpenCell("B7");
        cB7.BackColor = Color.SteelBlue;
        cB7.Value = "Communication";
        cB7.ForeColor = Color.Wheat;
        cB7.HorizontalAlignment = Aceoffix.ExcelWriter.XlHAlign.xlHAlignCenter;
        Aceoffix.ExcelWriter.Table titleTable = sheet.OpenTable("B3:E10");
        titleTable.Border.Weight =Aceoffix.ExcelWriter.XlBorderWeight.xlThick;
        titleTable.Border.LineColor = Color.FromArgb(255, 64, 0);
        titleTable.Border.BorderType=Aceoffix.ExcelWriter.XlBorderType.xlAllEdges;
        sheet.OpenTable("B1:E2").Merge();
        sheet.OpenTable("B1:E2").RowHeight=30;
        Aceoffix.ExcelWriter.Cell B1 = sheet.OpenCell("B1");
        B1.HorizontalAlignment = Aceoffix.ExcelWriter.XlHAlign.xlHAlignCenter ;
        B1.VerticalAlignment =Aceoffix.ExcelWriter.XlVAlign.xlVAlignCenter;
        B1.ForeColor = Color.FromArgb(255, 64, 0);
        B1.Value = "Consumption";
        B1.Font.Bold = true;
        B1.Font.Size = 25;
        
        AceoffixCtrl1.Bind(wb);
        AceoffixCtrl1.Menubar = false;
        AceoffixCtrl1.CustomToolbar = false;
        AceoffixCtrl1.SaveFilePage = "savefile.aspx";
        AceoffixCtrl1.OpenDocument("doc/test.xls", Aceoffix.OpenModeType.xlsNormalEdit, "John Scott");


    }
}