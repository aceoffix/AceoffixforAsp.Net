using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Drawing;

public partial class dataRegionFont_word : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        AceoffixCtrl1.ServerPage = "../aceoffix-runtime/server.aspx";

        Aceoffix.WordWriter.WordDocument doc = new Aceoffix.WordWriter.WordDocument();
        Aceoffix.WordWriter.DataRegion paragraph1 = doc.OpenDataRegion("Paragraph1");
        paragraph1.Font.Color = Color.Blue;//Font Color
        paragraph1.Font.Name = "Arial Rounded MT Bold";// Font Style
        paragraph1.Font.Size = 16;
        paragraph1.ParagraphFormat.Alignment = Aceoffix.WordWriter.WdParagraphAlignment.wdAlignParagraphCenter;//Alignment

        Aceoffix.WordWriter.DataRegion paragraph2 = doc.OpenDataRegion("Paragraph2");
        paragraph2.Font.Color = Color.Orange;//Font Color
        paragraph2.Font.Name = "Algerian";// Font Style
        paragraph2.Font.Size = 14;
        paragraph2.ParagraphFormat.Alignment = Aceoffix.WordWriter.WdParagraphAlignment.wdAlignParagraphLeft;//Alignment

        Aceoffix.WordWriter.DataRegion paragraph3 = doc.OpenDataRegion("Paragraph3");
        paragraph3.Font.Color = Color.Magenta;//Font Color
        paragraph3.Font.Name = "Blackadder ITC";// Font Style
        paragraph3.Font.Size = 12;
        paragraph3.ParagraphFormat.Alignment = Aceoffix.WordWriter.WdParagraphAlignment.wdAlignParagraphRight;//Alignment
        
        AceoffixCtrl1.Bind(doc);

        AceoffixCtrl1.OpenDocument("doc/introduce.doc", Aceoffix.OpenModeType.docReadOnly, "John Scott");

    }
}