using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;

namespace mytest
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Button3_Click(object sender, RibbonControlEventArgs e)
        {
            // To save as PDF
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.FileName = "*";
            dlg.DefaultExt = "pdf";
            dlg.ValidateNames = true;
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Globals.ThisAddIn.Application.ActiveDocument.ExportAsFixedFormat(dlg.FileName, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, OpenAfterExport: true);
            }
        }

        private void Button2_Click(object sender, RibbonControlEventArgs e)
        {
            // To insert a table
            Globals.ThisAddIn.Application.ActiveDocument.Tables.Add(Globals.ThisAddIn.Application.ActiveDocument.Range(0, 0), 3, 4);
            Globals.ThisAddIn.Application.ActiveDocument.Tables[1].Range.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorSeaGreen;
            Globals.ThisAddIn.Application.ActiveDocument.Tables[1].Range.Font.Size = 12;
            Globals.ThisAddIn.Application.ActiveDocument.Tables[1].Rows.Borders.Enable = 1;
        }

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {

           

        

        }
    }
}

