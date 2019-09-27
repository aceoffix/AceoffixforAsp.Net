using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aceoffix.ExcelWriter;
using System.Drawing;

public partial class drawExcel_excel : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Workbook wb = new Workbook();

        Aceoffix.ExcelWriter.Table backGroundTable = wb.OpenSheet("Sheet1").OpenTable("A1:P200");
        backGroundTable.Border.LineColor=Color.White;

        wb.OpenSheet("Sheet1").OpenTable("A1:H2").Merge();
        wb.OpenSheet("Sheet1").OpenTable("A1:H2").RowHeight=30;
        Cell A1 = wb.OpenSheet("Sheet1").OpenCell("A1");
        A1.HorizontalAlignment=Aceoffix.ExcelWriter.XlHAlign.xlHAlignCenter;
        A1.VerticalAlignment=Aceoffix.ExcelWriter.XlVAlign.xlVAlignCenter;
        A1.ForeColor = Color.FromArgb(255, 64, 0);
        A1.Value="Consumption";

        wb.OpenSheet("Sheet1").OpenTable("A1:A1").Font.Bold=true;
        wb.OpenSheet("Sheet1").OpenTable("A1:A1").Font.Size=25;


        Aceoffix.ExcelWriter.Table titleTable = wb.OpenSheet("Sheet1").OpenTable("B4:H5");
        titleTable.Border.BorderType = Aceoffix.ExcelWriter.XlBorderType.xlAllEdges;
        titleTable.Border.Weight=Aceoffix.ExcelWriter.XlBorderWeight.xlThick;
        titleTable.Border.LineColor = Color.FromArgb(255, 64, 0);

        Aceoffix.ExcelWriter.Table bodyTable = wb.OpenSheet("Sheet1").OpenTable("B6:H15");
        bodyTable.Border.LineColor=Color.Gray;
        bodyTable.Border.Weight = Aceoffix.ExcelWriter.XlBorderWeight.xlHairline;

        Border B7Border = wb.OpenSheet("Sheet1").OpenTable("B7:B7").Border;
        B7Border.LineColor=Color.White;

        Border B9Border = wb.OpenSheet("Sheet1").OpenTable("B9:B9").Border;
        B9Border.BorderType=Aceoffix.ExcelWriter.XlBorderType.xlBottomEdge;
        B9Border.LineColor=Color.White;

        Border C6C15BorderLeft = wb.OpenSheet("Sheet1").OpenTable("C6:C15").Border;
        C6C15BorderLeft.LineColor=Color.White;
        C6C15BorderLeft.BorderType=Aceoffix.ExcelWriter.XlBorderType.xlLeftEdge;

        Border C6C15BorderRight = wb.OpenSheet("Sheet1").OpenTable("C6:C15").Border;
        C6C15BorderRight.LineColor=Color.Yellow;
        C6C15BorderRight.LineStyle=Aceoffix.ExcelWriter.XlBorderLineStyle.xlDot;
        C6C15BorderRight.BorderType=Aceoffix.ExcelWriter.XlBorderType.xlRightEdge;

        Border E6E15Border = wb.OpenSheet("Sheet1").OpenTable("E6:E15").Border;
        E6E15Border.LineStyle=Aceoffix.ExcelWriter.XlBorderLineStyle.xlDot;
        E6E15Border.BorderType=Aceoffix.ExcelWriter.XlBorderType.xlAllEdges;
        E6E15Border.LineColor=Color.Yellow;

        Border G6G15BorderRight = wb.OpenSheet("Sheet1").OpenTable("G6:G15").Border;
        G6G15BorderRight.BorderType=Aceoffix.ExcelWriter.XlBorderType.xlRightEdge;
        G6G15BorderRight.LineColor=Color.White;

        Border G6G15BorderLeft = wb.OpenSheet("Sheet1").OpenTable("G6:G15").Border;
        G6G15BorderLeft.LineStyle=Aceoffix.ExcelWriter.XlBorderLineStyle.xlDot;
        G6G15BorderLeft.BorderType=XlBorderType.xlLeftEdge;
        G6G15BorderLeft.LineColor=Color.Yellow;

        Aceoffix.ExcelWriter.Table bodyTable2 = wb.OpenSheet("Sheet1").OpenTable("B6:H15");
        bodyTable2.Border.Weight=Aceoffix.ExcelWriter.XlBorderWeight.xlThick;
        bodyTable2.Border.LineColor = Color.FromArgb(255, 64, 0);
        bodyTable2.Border.BorderType=Aceoffix.ExcelWriter.XlBorderType.xlAllEdges;

        Border H16H17Border = wb.OpenSheet("Sheet1").OpenTable("H16:H17").Border;
        H16H17Border.LineColor = Color.FromArgb(204, 255, 204);

        Border E16G17Border = wb.OpenSheet("Sheet1").OpenTable("E16:G17").Border;
        E16G17Border.LineColor = Color.FromArgb(255, 64, 0);

        Aceoffix.ExcelWriter.Table footTable = wb.OpenSheet("Sheet1").OpenTable("B16:H17");
        footTable.Border.Weight=Aceoffix.ExcelWriter.XlBorderWeight.xlThick;
        footTable.Border.LineColor = Color.FromArgb(255, 64, 0);
        footTable.Border.BorderType=Aceoffix.ExcelWriter.XlBorderType.xlAllEdges;

        wb.OpenSheet("Sheet1").OpenTable("A1:A1").ColumnWidth=1;
        wb.OpenSheet("Sheet1").OpenTable("B1:B1").ColumnWidth=20;
        wb.OpenSheet("Sheet1").OpenTable("C1:C1").ColumnWidth = 15;
        wb.OpenSheet("Sheet1").OpenTable("D1:D1").ColumnWidth = 10;
        wb.OpenSheet("Sheet1").OpenTable("E1:E1").ColumnWidth = 8;
        wb.OpenSheet("Sheet1").OpenTable("F1:F1").ColumnWidth = 3;
        wb.OpenSheet("Sheet1").OpenTable("G1:G1").ColumnWidth = 12;
        wb.OpenSheet("Sheet1").OpenTable("H1:H1").ColumnWidth = 20;

        wb.OpenSheet("Sheet1").OpenTable("A16:A16").RowHeight=20;
        wb.OpenSheet("Sheet1").OpenTable("A17:A17").RowHeight = 20;

        for (int i = 0; i < 12; i++)
        {
            for (int j = 0; j < 7; j++)
            {
                wb.OpenSheet("Sheet1").OpenCellRC(4 + i, 2 + j).Font.Size=10;
            }
        }


        for (int i = 0; i < 10; i++)
        {
            wb.OpenSheet("Sheet1").OpenCell("H" + (6 + i)).BackColor=Color.FromArgb(255, 255, 153);
        }

        wb.OpenSheet("Sheet1").OpenCell("E16").BackColor = Color.FromArgb(255, 64, 0);
        wb.OpenSheet("Sheet1").OpenCell("F16").BackColor = Color.FromArgb(255, 64, 0);
        wb.OpenSheet("Sheet1").OpenCell("G16").BackColor = Color.FromArgb(255, 64, 0);
        wb.OpenSheet("Sheet1").OpenCell("E17").BackColor = Color.FromArgb(255, 64, 0);
        wb.OpenSheet("Sheet1").OpenCell("F17").BackColor = Color.FromArgb(255, 64, 0);
        wb.OpenSheet("Sheet1").OpenCell("G17").BackColor = Color.FromArgb(255, 64, 0);
        wb.OpenSheet("Sheet1").OpenCell("H16").BackColor=Color.FromArgb(204, 255, 204);
        wb.OpenSheet("Sheet1").OpenCell("H17").BackColor=Color.FromArgb(204, 255, 204);

        Aceoffix.ExcelWriter.Cell B4 = wb.OpenSheet("Sheet1").OpenCell("B4");
        B4.Font.Bold=true;
        B4.Value="Consumption";
        Aceoffix.ExcelWriter.Cell H5 = wb.OpenSheet("Sheet1").OpenCell("H5");
        H5.Font.Bold=true;
        H5.Value="total";
        H5.HorizontalAlignment=Aceoffix.ExcelWriter.XlHAlign.xlHAlignCenter;
        Aceoffix.ExcelWriter.Cell B6 = wb.OpenSheet("Sheet1").OpenCell("B6");
        B6.Font.Bold=true;
        B6.Value="airfare";
        Aceoffix.ExcelWriter.Cell B9 = wb.OpenSheet("Sheet1").OpenCell("B9");
        B9.Font.Bold=true;
        B9.Value="hotel";
        Aceoffix.ExcelWriter.Cell B11 = wb.OpenSheet("Sheet1").OpenCell("B11");
        B11.Font.Bold=true;
        B11.Value="repast";
        Aceoffix.ExcelWriter.Cell B12 = wb.OpenSheet("Sheet1").OpenCell("B12");
        B12.Font.Bold=true;
        B12.Value="transportation fee";
        Aceoffix.ExcelWriter.Cell B13 = wb.OpenSheet("Sheet1").OpenCell("B13");
        B13.Font.Bold=true;
        B13.Value="entertainment";
        Aceoffix.ExcelWriter.Cell B14 = wb.OpenSheet("Sheet1").OpenCell("B14");
        B14.Font.Bold=true;
        B14.Value="gift";
        Aceoffix.ExcelWriter.Cell B15 = wb.OpenSheet("Sheet1").OpenCell("B15");
        B15.Font.Bold=true;
        B15.Font.Size=10;
        B15.Value="other";

        wb.OpenSheet("Sheet1").OpenCell("C6").Value="air fare";
        wb.OpenSheet("Sheet1").OpenCell("C7").Value = "air fare(back)";
        wb.OpenSheet("Sheet1").OpenCell("C8").Value="other";
        wb.OpenSheet("Sheet1").OpenCell("C9").Value="Daily National Consumption";
        wb.OpenSheet("Sheet1").OpenCell("C10").Value="other";
        wb.OpenSheet("Sheet1").OpenCell("C11").Value="Daily National Consumption";
        wb.OpenSheet("Sheet1").OpenCell("C12").Value="Daily National Consumption";
        wb.OpenSheet("Sheet1").OpenCell("C13").Value="total";
        wb.OpenSheet("Sheet1").OpenCell("C14").Value="total";
        wb.OpenSheet("Sheet1").OpenCell("C15").Value="total";

        wb.OpenSheet("Sheet1").OpenCell("G6");
        wb.OpenSheet("Sheet1").OpenCell("G7");
        wb.OpenSheet("Sheet1").OpenCell("G9");
        wb.OpenSheet("Sheet1").OpenCell("G10");
        wb.OpenSheet("Sheet1").OpenCell("G11").Value=" day";
        wb.OpenSheet("Sheet1").OpenCell("G12").Value = " day";

        wb.OpenSheet("Sheet1").OpenCell("H6").Formula="=D6*F6";
        wb.OpenSheet("Sheet1").OpenCell("H7").Formula="=D7*F7";
        wb.OpenSheet("Sheet1").OpenCell("H8").Formula="=D8*F8";
        wb.OpenSheet("Sheet1").OpenCell("H9").Formula="=D9*F9";
        wb.OpenSheet("Sheet1").OpenCell("H10").Formula="=D10*F10";
        wb.OpenSheet("Sheet1").OpenCell("H11").Formula="=D11*F11";
        wb.OpenSheet("Sheet1").OpenCell("H12").Formula="=D12*F12";
        wb.OpenSheet("Sheet1").OpenCell("H13").Formula="=D13*F13";
        wb.OpenSheet("Sheet1").OpenCell("H14").Formula="=D14*F14";
        wb.OpenSheet("Sheet1").OpenCell("H15").Formula="=D15*F15";

        for (int i = 0; i < 10; i++)
        {

            wb.OpenSheet("Sheet1").OpenCell("D" + (6 + i)).NumberFormatLocal="$#,##0.00;$-#,##0.00";
            wb.OpenSheet("Sheet1").OpenCell("H" + (6 + i)).NumberFormatLocal="$#,##0.00;$-#,##0.00";
        }

        Cell E16 = wb.OpenSheet("Sheet1").OpenCell("E16");
        E16.Font.Bold=true;
        E16.Font.Size=11;
        E16.ForeColor=Color.White;
        E16.Value="The total amount";
        E16.VerticalAlignment = Aceoffix.ExcelWriter.XlVAlign.xlVAlignCenter;
        Cell E17 = wb.OpenSheet("Sheet1").OpenCell("E17");
        E17.Font.Bold = true;
        E17.Font.Size = 11;
        E17.ForeColor = Color.White;
        E17.Formula="=IF(C4>H16,\"budget\",\"budget\")";
        E17.VerticalAlignment=Aceoffix.ExcelWriter.XlVAlign.xlVAlignCenter;
        Cell H16 = wb.OpenSheet("Sheet1").OpenCell("H16");
        H16.VerticalAlignment=Aceoffix.ExcelWriter.XlVAlign.xlVAlignCenter;
        H16.NumberFormatLocal="$#,##0.00;$-#,##0.00";
        H16.Font.Name="Arial";
        H16.Font.Size=11;
        H16.Font.Bold=true;
        H16.Formula="=SUM(H6:H15)";
        Cell H17 = wb.OpenSheet("Sheet1").OpenCell("H17");
        H17.VerticalAlignment=Aceoffix.ExcelWriter.XlVAlign.xlVAlignCenter;
        H17.NumberFormatLocal="$#,##0.00;$-#,##0.00";
        H17.Font.Name="Arial";
        H17.Font.Size=11;
        H17.Font.Bold=true;
        H17.Formula="=(C4-H16)";

        Cell C4 = wb.OpenSheet("Sheet1").OpenCell("C4");
        C4.NumberFormatLocal="$#,##0.00;$-#,##0.00";
        C4.Value="2500";
        Cell D6 = wb.OpenSheet("Sheet1").OpenCell("D6");
        D6.NumberFormatLocal="$#,##0.00;$-#,##0.00";
        D6.Value="1200";
        wb.OpenSheet("Sheet1").OpenCell("F6").Font.Size=10;
        wb.OpenSheet("Sheet1").OpenCell("F6").Value="1";
        Cell D7 = wb.OpenSheet("Sheet1").OpenCell("D7");
        D7.NumberFormatLocal="$#,##0.00;$-#,##0.00";
        D7.Value="875";
        wb.OpenSheet("Sheet1").OpenCell("F7").Value="1";
        AceoffixCtrl1.Bind(wb);
        AceoffixCtrl1.Menubar = false;
        AceoffixCtrl1.CustomToolbar = false;

        AceoffixCtrl1.OpenDocument("doc/test.xls", Aceoffix.OpenModeType.xlsNormalEdit, "John Scott");


    }
}