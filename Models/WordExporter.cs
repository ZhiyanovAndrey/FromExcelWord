using Aspose.Cells;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;

namespace FromExcelWord.Models
{
    public class WordExporter
    {

        public static void WordExport(DataGrid myGrid)
        {
            if (myGrid == null || myGrid.Items.Count <= 0)
            {
                MessageBox.Show("Данные для экспорта не обнаружены.");
                return;
            }
            var wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Visible = true;
            var wordDoc = wordApp.Documents.Add();



            try
            {
                InsertDataWord(wordDoc, myGrid);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при экспорте: {ex.Message}");
            }
            finally
            {
                Marshal.ReleaseComObject(wordDoc);
                Marshal.ReleaseComObject(wordApp);
            }
        }

        private static void InsertDataWord(Document doc, DataGrid myGrid)
        {
            //var table = doc.Tables.Add(doc.Range(), myGrid.Items.Count + 1, myGrid.Columns.Count);

            //for (int j = 0; j < myGrid.Columns.Count; j++)
            //{
            //    table.Rows[1].Cells[j + 1].Range.Text = myGrid.Columns[j].HeaderText;
            //}

            //for (int i = 0; i < myGrid.Items.Count; i++)
            //{
            //    for (int j = 0; j < myGrid.Columns.Count; j++)
            //    {
            //        if (myGrid[j, i].Value != null)
            //        {
            //            table.Rows[i + 2].Cells[j + 1].Range.Text = myGrid[j, i].Value.ToString();
            //        }
            //    }
            //}

            //table.Rows[1].Range.Font.Bold = 1;
            //table.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            //table.Range.ParagraphFormat.SpaceAfter = 6;
            ////    table.Borders.Enable = 1;
        }


        //Workbook wb = new Workbook(path);

        //    foreach (Worksheet worksheet in wb.Worksheets)
        //    {
        //        MessageBox.Show(worksheet.Name);
        //    }


}
}

