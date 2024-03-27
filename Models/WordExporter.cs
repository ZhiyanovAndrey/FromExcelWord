using Microsoft.Office.Interop.Word;
using System;
using System.Windows;
using System.Windows.Controls;

namespace FromExcelWord.Models
{
    public class WordExporter
    {

        public void WordExport(DataGrid myGrid)
        {
            if (myGrid == null || myGrid.Items.Count <= 0)
            {
                MessageBox.Show("Данные для экспорта не обнаружены.");
                return;
            }
            var app = new Microsoft.Office.Interop.Word.Application();
            app.Visible = true;
            var doc = app.Documents.Add();



            try
            {
                InsertDataWord(doc, myGrid);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при экспорте: {ex.Message}");
            }
            finally
            {
                //Marshal.ReleaseComObject(wordDoc);
                //Marshal.ReleaseComObject(wordApp);

                // Сохраняем документ Word и закрываем приложение
                doc.Save();
                doc.Close();
                app.Quit();
            }
        }

        private void InsertDataWord(Document doc, DataGrid myGrid)
        {
            // создаем таблицу
            var table = doc.Tables.Add(doc.Range(), myGrid.Items.Count + 1, myGrid.Columns.Count);
            //Формат таблицы

            //table.Application.Selection.Tables[1].Rows[1].Select();
            //table.Application.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            //table.Application.Selection.Tables[1].Select();
            //table.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
            //table.Application.Selection.Tables[1].Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
            //table.Application.Selection.Tables[1].Rows[1].Select();
            //table.Application.Selection.InsertRowsAbove(1);
            //table.Application.Selection.Tables[1].Rows[1].Select();



            // заголовки
            for (int j = 0; j < myGrid.Columns.Count; j++)
            {

                table.Cell(1, 1).Range.Text = "Отдел";
                table.Cell(1, 2).Range.Text = "Количество задач";
            }



            for (int i = 0; i < myGrid.Items.Count; i++)
            {
                DataGridRow row = (DataGridRow)myGrid.ItemContainerGenerator.ContainerFromIndex(i);
                for (int j = 0; j < myGrid.Columns.Count; j++)
                {
                    if (myGrid.Columns[j] != null)
                    {

                        TextBlock cellContent = myGrid.Columns[j].GetCellContent(row) as TextBlock;
                        string cellValue = cellContent == null ? "" : cellContent.Text;
                        table.Cell(i + 2, j + 1).Range.Text = cellValue;
                    }

                }

            }

            table.Application.Selection.Tables[1].Borders.Enable = 1; // включаем все границы
            //Стиль заголовка таблицы
            table.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
            table.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Calibri";
            table.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 11;
            table.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            //table.Range.ParagraphFormat.SpaceAfter = 6; // отступ

            //table.Application.Selection.Tables[1].Rows[1].Alignment= WdRowAlignment.wdAlignRowCenter;

            //table.Rows[1].Range.Font.Bold = 1;
            //table.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            //table.Borders.Enable = 1;
        }


        //Workbook wb = new Workbook(path);

        //    foreach (Worksheet worksheet in wb.Worksheets)
        //    {
        //        MessageBox.Show(worksheet.Name);
        //    }


    }
}

